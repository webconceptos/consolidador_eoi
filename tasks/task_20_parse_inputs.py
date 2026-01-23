# tasks/task_20_parse_inputs.py
# -*- coding: utf-8 -*-
"""
Task 20: parse_inputs
"""

import argparse
import csv
import json
import re
from pathlib import Path
from datetime import datetime, date
from typing import Dict, Any, List, Optional


import sys
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))


from parsers.eoi_excel import parse_eoi_excel
from parsers.eoi_pdf import parse_eoi_pdf


IN_FOLDER_NAME = "009. EDI RECIBIDAS"
OUT_FOLDER_NAME = "011. INSTALACIÓN DE COMITÉ"

FILES_SELECTED = "files_selected.csv"
LAYOUT_FILE = "config_layout.json"

OUT_JSONL = "consolidado.jsonl"
OUT_CSV = "consolidado.csv"
OUT_PARSE_LOG = "parse_log.csv"
OUT_DEBUG_LOG = "debug_parse_inputs.log"


# -------------------------
# Helpers
# -------------------------
def ts():
    return datetime.now().isoformat(timespec="seconds")


def norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())


def ensure_dir(p: Path):
    p.mkdir(parents=True, exist_ok=True)


def log_append(path: Path, msg: str):
    ensure_dir(path.parent)
    with path.open("a", encoding="utf-8") as f:
        f.write(f"[{ts()}] {msg}\n")


def load_global_config() -> dict:
    p = Path("configs/config.json")
    if not p.exists():
        return {}
    return json.loads(p.read_text(encoding="utf-8"))


def read_selected_csv(path: Path) -> List[Dict[str, str]]:
    """
    Espera columnas:
      n, carpeta_postulante, archivo, tipo, ruta
    """
    rows: List[Dict[str, str]] = []
    if not path.exists():
        return rows
    with path.open("r", encoding="utf-8") as f:
        r = csv.DictReader(f)
        for row in r:
            rows.append({k: (row.get(k) or "").strip() for k in (r.fieldnames or [])})
    return rows


def write_jsonl(path: Path, items: List[Dict[str, Any]]):
    ensure_dir(path.parent)
    with path.open("w", encoding="utf-8") as f:
        for obj in items:
            f.write(json.dumps(json_safe(obj), ensure_ascii=False) + "\n")


def write_csv(path: Path, header: List[str], rows: List[List[Any]]):
    ensure_dir(path.parent)
    with path.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


def safe_int(x, default=0):
    try:
        return int(x)
    except Exception:
        return default


def to_iso(x):
    if isinstance(x, datetime):
        return x.isoformat(timespec="seconds")
    if isinstance(x, date):
        return x.isoformat()
    return x


def json_safe(obj):
    if isinstance(obj, (datetime, date)):
        return to_iso(obj)
    if isinstance(obj, Path):
        return str(obj)
    if isinstance(obj, dict):
        return {str(k): json_safe(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple, set)):
        return [json_safe(v) for v in obj]
    if obj is None or isinstance(obj, (str, int, float, bool)):
        return obj
    return str(obj)


def normalize_phone(s: str) -> str:
    s = norm(s)
    digits = re.sub(r"\D+", "", s)
    return digits


def normalize_dni(s: str) -> str:
    s = norm(s)
    m = re.search(r"\b(\d{8})\b", s)
    return m.group(1) if m else s


def normalize_email(s: str) -> str:
    s = norm(s)
    m = re.search(r"([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})", s)
    return m.group(1) if m else s


def flatten_summary_row(proceso: str, meta: Dict[str, str], data: Dict[str, Any]) -> List[Any]:
    return [
        proceso,
        meta.get("carpeta_postulante", ""),
        meta.get("archivo", ""),
        meta.get("tipo", ""),
        meta.get("ruta", ""),
        data.get("dni", ""),
        data.get("nombre_full", ""),
        data.get("email", ""),
        data.get("celular", ""),
        safe_int(data.get("exp_general_dias", 0)),
        safe_int(data.get("exp_especifica_dias", 0)),
        len(data.get("experiencias", []) or []),
        len(data.get("cursos", []) or []),
        data.get("_parse_warnings", ""),
    ]


def attach_meta(proceso: str, meta: Dict[str, str], data: Dict[str, Any], ftype: str) -> Dict[str, Any]:
    out = dict(data) if isinstance(data, dict) else {}
    out["_meta"] = {
        "proceso": proceso,
        "carpeta_postulante": meta.get("carpeta_postulante", ""),
        "archivo": meta.get("archivo", ""),
        "tipo": ftype,
        "ruta": meta.get("ruta", ""),
        "parsed_at": ts(),
    }
    return out


def post_normalize(data: Dict[str, Any]) -> Dict[str, Any]:
    data = dict(data)

    dni = normalize_dni(str(data.get("dni", "") or ""))
    email = normalize_email(str(data.get("email", "") or ""))
    cel = normalize_phone(str(data.get("celular", "") or ""))

    data["dni"] = dni
    data["email"] = email
    data["celular"] = cel
    data["nombre_full"] = norm(str(data.get("nombre_full", "") or ""))

    cursos = data.get("cursos", []) or []
    if isinstance(cursos, list):
        data["cursos"] = [norm(str(x)) for x in cursos if norm(str(x))]
    else:
        data["cursos"] = []

    exps = data.get("experiencias", []) or []
    if not isinstance(exps, list):
        exps = []
    clean_exps = []
    for e in exps:
        if not isinstance(e, dict):
            continue
        clean_exps.append({
            "fi": e.get("fi"),
            "ff": e.get("ff"),
            "entidad": norm(str(e.get("entidad", "") or "")),
            "cargo": norm(str(e.get("cargo", "") or "")),
            "proyecto": norm(str(e.get("proyecto", "") or "")),
            "funciones": norm(str(e.get("funciones", "") or "")),
        })
    data["experiencias"] = clean_exps

    data["java_ok"] = bool(data.get("java_ok", False))
    data["oracle_ok"] = bool(data.get("oracle_ok", False))

    warnings = []
    if dni and not re.fullmatch(r"\d{8}", dni):
        warnings.append(f"DNI_NO_8_DIGITOS:{dni}")
    if email and "@" not in email:
        warnings.append(f"EMAIL_RARO:{email}")
    if cel and len(cel) < 7:
        warnings.append(f"CEL_RARO:{cel}")

    data["_parse_warnings"] = " | ".join(warnings)
    return data


# -------------------------
# Layout dinámico desde Task_00
# -------------------------
def load_layout_json(layout_path: Path) -> Dict[str, Any]:
    if not layout_path.exists():
        return {}
    try:
        return json.loads(layout_path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _int_or_none(x) -> Optional[int]:
    try:
        if x is None:
            return None
        return int(x)
    except Exception:
        return None


def build_parser_layout_from_config_layout(layout_cfg: Dict[str, Any]) -> Dict[str, Any]:
    """
    Convierte config_layout.json (Task_00) en layout para parse_eoi_excel(layout=...).

    Regla:
    - DP fijo 12..23 (porque ese formato es estable)
    - EG dinámico:
        start = label_rows_detectados.exp_general + 1 (típicamente debajo del título)
        end = antes de exp_especifica/entrevista/puntaje_total (el primero que aparezca)
      Si algo falta, fallback a 101..145
    """
    label = (layout_cfg or {}).get("label_rows_detectados") or {}

    exp_general_label = _int_or_none(label.get("exp_general"))
    exp_especifica_label = _int_or_none(label.get("exp_especifica"))
    entrevista_label = _int_or_none(label.get("entrevista"))
    puntaje_total_label = _int_or_none(label.get("puntaje_total"))

    # DP estable
    dp_layout = {"start_row": 12, "end_row": 23, "max_cols": 12}

    # EG dinámico
    if exp_general_label:
        eg_start = exp_general_label + 1

        # el fin es la primera sección posterior
        candidates = [x for x in [exp_especifica_label, puntaje_total_label, entrevista_label] if x]
        eg_end = (min(candidates) - 1) if candidates else (eg_start + 60)

        # sanidad
        if eg_end < eg_start:
            eg_end = eg_start + 40

        eg_layout = {"start_row": eg_start, "end_row": eg_end}
    else:
        eg_layout = {"start_row": 101, "end_row": 145}

    return {
        "datos_personales": dp_layout,
        "experiencia_general": eg_layout,
    }


# -------------------------
# Main
# -------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", type=str, default="", help="Carpeta raíz: contiene los procesos")
    ap.add_argument("--limit", type=int, default=0, help="Limitar N archivos por proceso (0 = sin límite)")
    ap.add_argument("--only-proc", type=str, default="", help="Procesar solo procesos cuyo nombre contenga este texto")
    ap.add_argument("--use-ocr", action="store_true", help="Forzar OCR en PDF (si tu parser lo soporta)")
    args = ap.parse_args()

    cfg = load_global_config()
    root = Path(args.root) if args.root.strip() else Path(cfg.get("input_root", ""))

    if not str(root).strip():
        raise SystemExit("Debes pasar --root o definir input_root en configs/config.json")
    if not root.exists():
        raise SystemExit(f"No existe root: {root}")

    use_ocr = bool(cfg.get("pdf", {}).get("use_ocr", False)) or bool(args.use_ocr)

    procesos = [p for p in root.iterdir() if p.is_dir()]
    procesos.sort(key=lambda x: x.name.lower())

    only_filter = norm(args.only_proc).lower()

    print(f"[parse_inputs] root = {root}")
    print(f"[parse_inputs] procesos detectados = {len(procesos)} | use_ocr={use_ocr}")

    ok_proc = 0
    skip_proc = 0

    for proc_dir in procesos:
        proceso = proc_dir.name
        if only_filter and only_filter not in proceso.lower():
            continue

        in_dir = proc_dir / IN_FOLDER_NAME
        out_dir = proc_dir / OUT_FOLDER_NAME
        selected_path = out_dir / FILES_SELECTED
        layout_path = out_dir / LAYOUT_FILE

        if not selected_path.exists():
            skip_proc += 1
            print(f"  - SKIP: {proceso} (no existe {FILES_SELECTED} en 011; ejecuta Task 10)")
            continue

        selected = read_selected_csv(selected_path)
        if not selected:
            skip_proc += 1
            print(f"  - SKIP: {proceso} ({FILES_SELECTED} vacío)")
            continue

        if args.limit and args.limit > 0:
            selected = selected[: args.limit]

        ensure_dir(out_dir)

        debug_log = out_dir / OUT_DEBUG_LOG
        if debug_log.exists():
            debug_log.unlink(missing_ok=True)

        # cargar layout por proceso (Task_00)
        layout_cfg = load_layout_json(layout_path)
        parser_layout = build_parser_layout_from_config_layout(layout_cfg)

        log_append(debug_log, f"== PROCESO: {proceso} ==")
        log_append(debug_log, f"selected_count: {len(selected)}")
        log_append(debug_log, f"use_ocr: {use_ocr}")
        log_append(debug_log, f"selected_file: {selected_path}")
        log_append(debug_log, f"layout_file_exists: {layout_path.exists()}")
        log_append(debug_log, f"parser_layout: {json.dumps(parser_layout, ensure_ascii=False)}")

        jsonl_items: List[Dict[str, Any]] = []
        parse_log_rows: List[List[str]] = []
        consolidado_rows: List[List[Any]] = []

        consolidado_header = [
            "proceso", "carpeta_postulante", "archivo", "tipo", "ruta",
            "dni", "nombre_full", "email", "celular",
            "exp_general_dias", "exp_especifica_dias",
            "n_experiencias", "n_cursos",
            "warnings"
        ]

        for idx, meta in enumerate(selected, start=1):
            ruta = meta.get("ruta", "")
            if not ruta:
                parse_log_rows.append([ts(), proceso, "", meta.get("archivo", ""), meta.get("tipo", ""), "ERROR", "RUTA_VACIA"])
                continue

            fp = Path(ruta)
            ftype = (meta.get("tipo") or "").upper().strip()

            log_append(debug_log, f"[{idx}/{len(selected)}] START {fp}")

            if not fp.exists():
                log_append(debug_log, f"[{idx}] ERROR FILE_NOT_FOUND: {fp}")
                parse_log_rows.append([ts(), proceso, str(fp), meta.get("archivo", ""), ftype, "ERROR", "FILE_NOT_FOUND"])
                continue

            try:
                if fp.suffix.lower() in (".xlsx", ".xlsm", ".xls") or ftype == "EXCEL":
                    # ✅ Excel con layout dinámico
                    data = parse_eoi_excel(fp, layout=parser_layout)
                    ftype2 = "EXCEL"
                else:
                    data = parse_eoi_pdf(fp, use_ocr=use_ocr)
                    ftype2 = "PDF"

                if not isinstance(data, dict):
                    raise ValueError("PARSER_NO_DEVOLVIO_DICT")

                data = post_normalize(data)
                data = attach_meta(proceso, meta, data, ftype2)

                jsonl_items.append(data)
                consolidado_rows.append(flatten_summary_row(proceso, meta, data))

                parse_log_rows.append([ts(), proceso, str(fp), meta.get("archivo", ""), ftype2, "OK", ""])
                log_append(
                    debug_log,
                    f"[{idx}] OK dni='{data.get('dni','')}' nombre='{data.get('nombre_full','')}' "
                    f"expG={data.get('exp_general_dias',0)} expE={data.get('exp_especifica_dias',0)} "
                    f"warnings='{data.get('_parse_warnings','')}'"
                )

            except Exception as e:
                log_append(debug_log, f"[{idx}] EXCEPTION {repr(e)}")
                parse_log_rows.append([ts(), proceso, str(fp), meta.get("archivo", ""), ftype, "ERROR", repr(e)])

        # escribir outputs por proceso
        write_jsonl(out_dir / OUT_JSONL, jsonl_items)
        write_csv(out_dir / OUT_CSV, consolidado_header, consolidado_rows)
        write_csv(
            out_dir / OUT_PARSE_LOG,
            ["fecha", "proceso", "ruta", "archivo", "tipo", "estado", "detalle"],
            parse_log_rows,
        )

        log_append(debug_log, f"[DONE] jsonl={OUT_JSONL} rows={len(jsonl_items)}")
        log_append(debug_log, f"[DONE] csv={OUT_CSV} rows={len(consolidado_rows)}")
        log_append(debug_log, f"[DONE] parse_log={OUT_PARSE_LOG} rows={len(parse_log_rows)}")

        ok_proc += 1
        print(f"  - OK: {proceso} -> {OUT_JSONL} ({len(jsonl_items)} postulantes) | errores={sum(1 for r in parse_log_rows if r[5]=='ERROR')}")

    print("")
    print(f"[parse_inputs] resumen: OK_PROCESOS={ok_proc} SKIP_PROCESOS={skip_proc}")


if __name__ == "__main__":
    main()
