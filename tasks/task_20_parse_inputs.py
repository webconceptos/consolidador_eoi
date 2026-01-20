# tasks/task_20_parse_inputs.py
# -*- coding: utf-8 -*-
"""
Task 20: parse_inputs

Lee el listado de archivos elegidos por proceso (Task 10: files_selected.csv),
parsea cada PDF/Excel y genera salidas normalizadas por proceso:

En <Proceso>/011. INSTALACIÓN DE COMITÉ/
  - consolidado.jsonl       (1 JSON por postulante)
  - consolidado.csv         (resumen plano)
  - parse_log.csv           (OK/ERROR por archivo)
  - debug_parse_inputs.log  (log detallado)

Uso:
  python tasks/task_20_parse_inputs.py
  python tasks/task_20_parse_inputs.py --root "D:\...\ProcesoSelección"
  python tasks/task_20_parse_inputs.py --limit 10
  python tasks/task_20_parse_inputs.py --only-proc "SCI N° 069-2025 ANALISTA ..."
"""

import argparse
import csv
import json
import re
from pathlib import Path
#from datetime import datetime
from datetime import datetime, date
from typing import Dict, Any, List, Optional, Tuple

from parsers.eoi_excel import parse_eoi_excel
from parsers.eoi_pdf import parse_eoi_pdf

IN_FOLDER_NAME = "009. EDI RECIBIDA"
OUT_FOLDER_NAME = "011. INSTALACIÓN DE COMITÉ"

FILES_SELECTED = "files_selected.csv"

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
            rows.append({k: (row.get(k) or "").strip() for k in r.fieldnames or []})
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
    # date o datetime => isoformat
    if isinstance(x, datetime):
        return x.isoformat(timespec="seconds")
    if isinstance(x, date):
        return x.isoformat()
    return x

def json_safe(obj):
    """
    Convierte recursivamente dict/list/tuplas y tipos no serializables.
    """
    # fechas
    if isinstance(obj, (datetime, date)):
        return to_iso(obj)

    # Path
    if isinstance(obj, Path):
        return str(obj)

    # dict
    if isinstance(obj, dict):
        return {str(k): json_safe(v) for k, v in obj.items()}

    # list/tuple/set
    if isinstance(obj, (list, tuple, set)):
        return [json_safe(v) for v in obj]

    # primitivos OK
    if obj is None or isinstance(obj, (str, int, float, bool)):
        return obj

    # fallback
    return str(obj)


def normalize_phone(s: str) -> str:
    s = norm(s)
    digits = re.sub(r"\D+", "", s)
    # En Perú suelen ser 9 dígitos móvil; igual devolvemos todo lo numérico si hay.
    return digits


def normalize_dni(s: str) -> str:
    s = norm(s)
    # DNI peruano típico 8 dígitos
    m = re.search(r"\b(\d{8})\b", s)
    return m.group(1) if m else s


def normalize_email(s: str) -> str:
    s = norm(s)
    m = re.search(r"([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})", s)
    return m.group(1) if m else s


def flatten_summary_row(proceso: str, meta: Dict[str, str], data: Dict[str, Any]) -> List[Any]:
    """
    Fila resumen para consolidado.csv
    """
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
    """
    Normaliza campos clave SIN inventar información.
    """
    data = dict(data)

    dni = normalize_dni(str(data.get("dni", "") or ""))
    email = normalize_email(str(data.get("email", "") or ""))
    cel = normalize_phone(str(data.get("celular", "") or ""))

    data["dni"] = dni
    data["email"] = email
    data["celular"] = cel

    # nombre_full: solo normaliza espacios
    data["nombre_full"] = norm(str(data.get("nombre_full", "") or ""))

    # cursos: limpia vacíos
    cursos = data.get("cursos", []) or []
    if isinstance(cursos, list):
        data["cursos"] = [norm(str(x)) for x in cursos if norm(str(x))]
    else:
        data["cursos"] = []

    # experiencias: asegurar lista de dict
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

    # flags tecnológicos (si los parsers ya lo ponen, lo respetamos)
    data["java_ok"] = bool(data.get("java_ok", False))
    data["oracle_ok"] = bool(data.get("oracle_ok", False))

    # warnings
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

        log_append(debug_log, f"== PROCESO: {proceso} ==")
        log_append(debug_log, f"selected_count: {len(selected)}")
        log_append(debug_log, f"use_ocr: {use_ocr}")
        log_append(debug_log, f"selected_file: {selected_path}")

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
                    data = parse_eoi_excel(fp)
                    ftype2 = "EXCEL"
                else:
                    data = parse_eoi_pdf(fp, use_ocr=use_ocr)
                    ftype2 = "PDF"
                
                #####LOG - FECHA - FGARCIAA
                import pprint
                log_append(debug_log, "[DEBUG] tipos sospechosos:")
                for k,v in (data or {}).items():
                    if isinstance(v, (date, datetime)):
                        log_append(debug_log, f"  TOP_LEVEL_DATE: {k}={v}")
                ###################

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
