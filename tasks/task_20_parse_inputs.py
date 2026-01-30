# tasks/task_20_parse_inputs.py
# -*- coding: utf-8 -*-
"""
Task 20 definitivo — Parse Inputs (ETL)

- Lee 011/files_selected.csv
- Parsea cada EDI con parsers definitivos (Excel/PDF)
- Exporta:
    011/parsed_postulantes.jsonl
    011/parsed_postulantes.csv
    011/parse_log.csv
    011/debug_parse_inputs.log
"""

import argparse
import csv
import json
import re
from pathlib import Path
from datetime import datetime

import datetime as dt
from datetime import datetime, date, timedelta
from typing import Dict, Any, List, Tuple, Optional

from parsers.eoi_excel import parse_eoi_excel
from parsers.eoi_pdf import parse_eoi_pdf  # asumiendo que ya lo tienes
from parsers.eoi_pdf_pro import parse_eoi_pdf_pro
import sys
import pytesseract

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")




pytesseract.pytesseract.tesseract_cmd = (
    r"C:\Users\67733\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
)


OUT_FOLDER_NAME = "011. INSTALACIÓN DE COMITÉ"
FILES_SELECTED = "files_selected.csv"

OUT_JSONL = "parsed_postulantes.jsonl"
OUT_CSV = "parsed_postulantes.csv"
OUT_PARSE_LOG = "parse_log.csv"
OUT_DEBUG_LOG = "debug_parse_inputs.log"

_DATE_FMT = "%d/%m/%Y"
_CAL_ANCHOR = date(2000, 1, 1)  # ancla fija para convertir días -> (y,m,d) real

try:
    from dateutil.relativedelta import relativedelta
except Exception:
    relativedelta = None  # si no está dateutil instalado

def ts() -> str:
    return datetime.now().isoformat(timespec="seconds")


def norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())


def ensure_dir(p: Path) -> None:
    p.mkdir(parents=True, exist_ok=True)


def log_append(path: Path, msg: str) -> None:
    ensure_dir(path.parent)
    with path.open("a", encoding="utf-8") as f:
        f.write(f"[{ts()}] {msg}\n")


def normalize_phone(s: str) -> str:
    return re.sub(r"\D+", "", norm(s))


def normalize_dni(s: str) -> str:
    m = re.search(r"\b(\d{8})\b", norm(s))
    return m.group(1) if m else norm(s)


def normalize_email(s: str) -> str:
    s = norm(s)
    m = re.search(r"([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})", s)
    return m.group(1) if m else s


def read_selected_csv(path: Path) -> List[Dict[str, str]]:
    rows = []
    with path.open("r", encoding="utf-8") as f:
        r = csv.DictReader(f)
        for row in r:
            rows.append({k: (row.get(k) or "").strip() for k in (r.fieldnames or [])})
    return rows

def _json_sanitize(o):
    # Pathlib (WindowsPath, PosixPath)
    if isinstance(o, Path):
        return str(o)
    # datetime/date
    if isinstance(o, (dt.datetime, dt.date)):
        return o.isoformat()
    # set/tuple
    if isinstance(o, (set, tuple)):
        return list(o)
    # fallback: que explote con TypeError si sigue raro
    raise TypeError(f"Object of type {o.__class__.__name__} is not JSON serializable")

def deep_sanitize(x):
    if isinstance(x, Path):
        return str(x)
    if isinstance(x, (dt.datetime, dt.date)):
        return x.isoformat()
    if isinstance(x, dict):
        return {k: deep_sanitize(v) for k, v in x.items()}
    if isinstance(x, list):
        return [deep_sanitize(v) for v in x]
    if isinstance(x, tuple):
        return [deep_sanitize(v) for v in x]
    if isinstance(x, set):
        return [deep_sanitize(v) for v in x]
    return x

def write_jsonl(path, items):
    with open(path, "w", encoding="utf-8") as f:
        for obj in items:
            f.write(json.dumps(obj, ensure_ascii=False, default=_json_sanitize) + "\n")

def write_csv(path: Path, header: List[str], rows: List[List[Any]]) -> None:
    ensure_dir(path.parent)
    with path.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)

def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())

def _parse_date(s: str) -> Optional[date]:
    s = (s or "").strip()
    if not s:
        return None
    try:
        return datetime.strptime(s, _DATE_FMT).date()
    except Exception:
        return None

def _merge_intervals(intervals: List[Tuple[date, date]]) -> List[Tuple[date, date]]:
    """
    intervals: lista de (start, end) con end INCLUSIVO.
    Mergea superposiciones y adyacentes.
    """
    if not intervals:
        return []
    intervals = sorted(intervals, key=lambda x: (x[0], x[1]))
    merged = [intervals[0]]
    for s, e in intervals[1:]:
        ps, pe = merged[-1]
        # si se superpone o es adyacente (pe + 1 día >= s), unir
        if s <= pe + timedelta(days=1):
            merged[-1] = (ps, max(pe, e))
        else:
            merged.append((s, e))
    return merged

##a task20
def _days_inclusive(s: date, e: date) -> int:
    return (e - s).days + 1

##a task20
def _days_to_ymd_calendar_real(total_days: int, anchor: date = _CAL_ANCHOR) -> Tuple[int, int, int]:
    """
    Convierte días -> (años, meses, días) en calendario real usando un anchor fijo.
    Requiere python-dateutil.
    """
    if relativedelta is None:
        raise RuntimeError("Falta python-dateutil. Instala con: pip install python-dateutil")

    end = anchor + timedelta(days=total_days)
    rd = relativedelta(end, anchor)
    return rd.years, rd.months, rd.days


def compute_experience_summary_and_total_calendar_real(
    exp_block: Dict[str, Any],
    anchor: date = _CAL_ANCHOR
) -> Tuple[str, Tuple[int, int, int], int, List[Tuple[date, date]],str]:
    """
    Retorna:
      - resumen_text
      - (años, meses, días) calendario real (con anchor fijo)
      - total_days_unicos (sin superposición)
      - merged_intervals (para depuración)
    """
    items = exp_block.get("items") or []
    if not isinstance(items, list):
        items = []

    resumen_parts: List[str] = []
    detalle_parts: List[str] = []
    raw_intervals: List[Tuple[date, date]] = []

    for it in items:
        if not isinstance(it, dict):
            continue

        entidad = _norm(it.get("entidad", ""))
        cargo = _norm(it.get("cargo", ""))
        f1s = _norm(it.get("fecha_inicio", ""))
        f2s = _norm(it.get("fecha_fin", ""))

        d1 = _parse_date(f1s)
        d2 = _parse_date(f2s)

        # Resumen
        header = f"{entidad} - {cargo}".strip(" -")
        if f1s or f2s:
            header += f" | {f1s or '?'} a {f2s or '?'}"

        desc = (it.get("descripcion") or "").strip()
        if desc:
            resumen_parts.append(f"{header}\n  Desc: {desc}")
        else:
            resumen_parts.append(header)

        detalle_parts.append(f"- {header}")

        # Intervalos (solo si hay fechas válidas)
        if d1 and d2:
            if d2 < d1:
                d1, d2 = d2, d1
            raw_intervals.append((d1, d2))

    resumen_text = "\n\n".join([p for p in resumen_parts if p.strip()]).strip()
    detalle_text = "\n".join([p for p in detalle_parts if p.strip()]).strip()    

    merged = _merge_intervals(raw_intervals)
    total_days = sum(_days_inclusive(s, e) for s, e in merged)
    y, m, d = _days_to_ymd_calendar_real(total_days, anchor=anchor)

    return resumen_text, (y, m, d), total_days, merged, detalle_text

def format_ymd(y: int, m: int, d: int) -> str:
    return f"{y} año(s), {m} mes(es), {d} día(s)"

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True, help="Carpeta raíz con procesos")
    ap.add_argument("--only-proc", default="", help="Procesar solo procesos cuyo nombre contenga este texto")
    ap.add_argument("--use-ocr", action="store_true", help="Usar OCR para PDF (si tu parser lo soporta)")
    args = ap.parse_args()

    root = Path(args.root)
    only_filter = norm(args.only_proc).lower()
    use_ocr = bool(args.use_ocr)

    procesos = [p for p in root.iterdir() if p.is_dir()]
    procesos.sort(key=lambda x: x.name.lower())

    print(f"[task_20_parse_inputs] root={root} procesos={len(procesos)} use_ocr={use_ocr}")

    ok, skip, fail = 0, 0, 0

    for proc_dir in procesos:
        proceso = proc_dir.name
        if only_filter and only_filter not in proceso.lower():
            continue

        out_dir = proc_dir / OUT_FOLDER_NAME
        selected_path = out_dir / FILES_SELECTED

        if not selected_path.exists():
            print(f"  - SKIP: {proceso} (falta 011/{FILES_SELECTED})")
            skip += 1
            continue

        selected = read_selected_csv(selected_path)
        if not selected:
            print(f"  - SKIP: {proceso} (files_selected vacío)")
            skip += 1
            continue

        ensure_dir(out_dir)
        dbg = out_dir / OUT_DEBUG_LOG
        if dbg.exists():
            dbg.unlink(missing_ok=True)

        items_jsonl: List[Dict[str, Any]] = []
        parse_log_rows: List[List[Any]] = []

        for i, meta in enumerate(selected, start=1):
            ruta = meta.get("ruta", "")
            fp = Path(ruta) if ruta else None
            if not fp or not fp.exists():
                parse_log_rows.append([ts(), proceso, ruta, meta.get("archivo",""), meta.get("tipo",""), "ERROR", "FILE_NOT_FOUND"])
                continue

            try:
                if fp.suffix.lower() in (".xlsx", ".xlsm", ".xls") or (meta.get("tipo","").upper() == "EXCEL"):
                    data = parse_eoi_excel(fp)
                    tipo = "EXCEL"
                else:
                    #data = parse_eoi_pdf(fp, use_ocr=use_ocr)
                    data = parse_eoi_pdf_pro(fp, use_ocr=True)
                    tipo = "PDF"

                # normalizaciones finales (consistentes)
                data["dni"] = normalize_dni(str(data.get("dni","")))
                data["email"] = normalize_email(str(data.get("email","")))
                data["celular"] = normalize_phone(str(data.get("celular","")))
                data["nombre_full"] = norm(str(data.get("nombre_full","")))

                print(data["dni"])
                resumen_exp_general, (y, m, d), total_days, merged, detalle_exp_general = compute_experience_summary_and_total_calendar_real(data.get("exp_general") or {})
                #total_exp_general_texto= y + "Año(s)" + m + "Mes(es)" + d + "día(s)"                
                total_exp_general=(format_ymd(y, m, d))
                print(total_exp_general)

                resumen_exp_especifica, (y, m, d), total_days, merged , detalle_exp_especifica= compute_experience_summary_and_total_calendar_real(data.get("exp_especifica") or {})
                #total_exp_especifica_texto= y + "Año(s)" + m + "Mes(es)" + d + "día(s)"
                total_exp_especifica=(format_ymd(y, m, d))
                print(total_exp_especifica)
                

                # payload listo para Task 40 (solo valores)
                data["_fill_payload"] = {
                    "dni": data.get("dni",""),
                    "nombre_full": data.get("nombre_full",""),
                    "email": data.get("email",""),
                    "celular": data.get("celular",""),
                    "formacion_obligatoria_resumen": (data.get("formacion_obligatoria") or {}).get("resumen",""),
                    "estudios_complementarios_resumen": (data.get("estudios_complementarios") or {}).get("resumen",""),
                    "exp_general_detalle_text": resumen_exp_general,
                    "exp_general_resumen_text": detalle_exp_general,
                    "exp_general_total_text": total_exp_general,
                    "exp_general_dias": int(data.get("exp_general_dias",0) or 0),
                    "exp_especifica_detalle_text": resumen_exp_especifica,
                    "exp_especifica_resumen_text": detalle_exp_especifica,
                    "exp_especifica_total_text": total_exp_especifica,
                    "exp_especifica_dias": int(data.get("exp_especifica_dias",0) or 0),
                }

                # meta
                data["_meta"] = {
                    "proceso": proceso,
                    "carpeta_postulante": meta.get("carpeta_postulante",""),
                    "archivo": meta.get("archivo",""),
                    "tipo": tipo,
                    "ruta": ruta,
                    "parsed_at": ts(),
                }

                items_jsonl.append(data)
                parse_log_rows.append([ts(), proceso, ruta, meta.get("archivo",""), tipo, "OK", ""])
                log_append(dbg, f"[{i}/{len(selected)}] OK {meta.get('archivo','')} dni={data.get('dni','')}")

            except Exception as e:
                parse_log_rows.append([ts(), proceso, ruta, meta.get("archivo",""), meta.get("tipo",""), "ERROR", repr(e)])
                log_append(dbg, f"[{i}/{len(selected)}] ERROR {meta.get('archivo','')} {repr(e)}")

        # Export
        write_jsonl(out_dir / OUT_JSONL, items_jsonl)

        header = [
            "proceso","carpeta_postulante","archivo","tipo","ruta",
            "dni","nombre_full","email","celular",
            "formacion_obligatoria_resumen",
            "exp_general_dias","exp_especifica_dias"
        ]
        rows = []
        for d in items_jsonl:
            m = d.get("_meta", {}) or {}
            fp = d.get("_fill_payload", {}) or {}
            rows.append([
                proceso,
                m.get("carpeta_postulante",""),
                m.get("archivo",""),
                m.get("tipo",""),
                m.get("ruta",""),
                d.get("dni",""),
                d.get("nombre_full",""),
                d.get("email",""),
                d.get("celular",""),
                fp.get("formacion_obligatoria_resumen",""),
                fp.get("exp_general_dias",0),
                fp.get("exp_especifica_dias",0),
            ])

        write_csv(out_dir / OUT_CSV, header, rows)
        write_csv(out_dir / OUT_PARSE_LOG, ["fecha","proceso","ruta","archivo","tipo","estado","detalle"], parse_log_rows)

        errs = sum(1 for r in parse_log_rows if r[5] == "ERROR")
        print(f"  - OK: {proceso} -> {OUT_JSONL} ({len(items_jsonl)} postulantes) | errores={errs}")
        ok += 1

    print(f"\n[task_20_parse_inputs] resumen OK={ok} SKIP={skip} FAIL={fail}")


if __name__ == "__main__":
    main()
