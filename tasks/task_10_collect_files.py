# tasks/task_10_collect_files.py
# -*- coding: utf-8 -*-
"""
TASK 10 — Collect Files (por proceso)

Qué hace (en el pipeline actualizado)
-------------------------------------
Para cada proceso dentro de --root:

1) Ubica la carpeta de postulantes:
      <PROCESO>/009. EDI RECIBIDAS
   (soporta variantes: "009. EDI RECIBIDA" y "009. EDI RECIBIDAS")

2) Define "postulante" como:
      carpeta DIRECTA (primer nivel) dentro de 009

   OJO: Esto corrige el error típico de agrupar por p.parent cuando hay subcarpetas,
        que puede fragmentar un postulante en múltiples "postulantes".

3) Dentro de cada carpeta de postulante, busca de forma recursiva archivos elegibles:
      .xlsx, .xlsm, .xls, .pdf

   y elige SOLO UNO (1) por postulante, aplicando un scoring:
   - Prefiere Excel sobre PDF.
   - Penaliza PDFs que parezcan "correo/presentación" (por default se omiten).

4) Genera salidas en:
      <PROCESO>/011. INSTALACIÓN DE COMITÉ/

   - files_selected.csv  (1 fila por postulante con el archivo elegido)
   - files_skipped.csv   (postulantes sin archivo elegible o solo PDF tipo correo)
   - files_manifest.csv  (opcional pero recomendado: inventario completo por postulante)
   - debug_collect_files.log
   - collect_summary.json (resumen + validaciones)

Compatibilidad
--------------
- Mantiene el formato esperado por Task 20:
    files_selected.csv columnas: n, carpeta_postulante, archivo, tipo, ruta

Integración con Task 00
-----------------------
- Lee (si existe) el config_layout.json generado por Task 00 en 011
  para auditoría (no bloquea la ejecución si falta, pero registra warning).
- Compara conteos: carpetas en 009 vs runtime.total_postulantes (si está disponible).

Uso
---
python tasks/task_10_collect_files.py --root "C:\\IA_Investigacion\\ProcesoSelección"

Opcionales:
--only-proc "SCI N° 068-2025"
--dry-run
--allow-bad-pdf   (incluye PDFs tipo correo si es lo único que hay)
"""

import argparse
import csv
import json
import re
import sys
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Tuple, Optional, Any

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

# -------------------------
# Convenciones de carpetas
# -------------------------
OUT_FOLDER_NAME = "011. INSTALACIÓN DE COMITÉ"
IN_FOLDER_CANDIDATES = ["009. EDI RECIBIDA", "009. EDI RECIBIDAS"]

LAYOUT_FILE = "config_layout.json"

OUT_SELECTED = "files_selected.csv"
OUT_SKIPPED = "files_skipped.csv"
OUT_MANIFEST = "files_manifest.csv"
OUT_LOG = "debug_collect_files.log"
OUT_SUMMARY = "collect_summary.json"

ELIGIBLE_EXTS = {".xlsx", ".xlsm", ".xls", ".pdf"}


# -------------------------
# Helpers
# -------------------------
def ts() -> str:
    return datetime.now().isoformat(timespec="seconds")


def norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())


def ensure_dir(p: Path):
    p.mkdir(parents=True, exist_ok=True)


def log_append(path: Path, msg: str):
    ensure_dir(path.parent)
    with path.open("a", encoding="utf-8") as f:
        f.write(f"[{ts()}] {msg}\n")


def write_csv(path: Path, header: List[str], rows: List[List[Any]]):
    ensure_dir(path.parent)
    with path.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


def write_json(path: Path, obj: dict):
    ensure_dir(path.parent)
    path.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")


def load_global_config() -> dict:
    p = Path("configs/config.json")
    if not p.exists():
        return {}
    return json.loads(p.read_text(encoding="utf-8"))


def safe_int(x, default=0) -> int:
    try:
        return int(x)
    except Exception:
        return default


def find_009_dir(proc_dir: Path) -> Optional[Path]:
    for name in IN_FOLDER_CANDIDATES:
        p = proc_dir / name
        if p.exists() and p.is_dir():
            return p
    return None


def list_postulante_folders(in_dir_009: Path) -> List[Path]:
    """
    Postulante = carpeta DIRECTA dentro de 009.
    """
    return sorted(
        [p for p in in_dir_009.iterdir() if p.is_dir() and not p.name.startswith(".")],
        key=lambda x: x.name.lower()
    )


# -------------------------
# Scoring de archivos
# -------------------------
def is_bad_pdf_name(name: str) -> bool:
    n = name.lower()
    # típicos adjuntos que NO son la EDI real
    return any(k in n for k in ("correo", "presentacion", "presentación", "mail", "mensaje", "email"))


def score_excel(f: Path) -> int:
    """
    Puntuación simple:
    - preferencia por extensión (xlsx > xlsm > xls)
    - bonus por keywords esperables
    - penalización por "plantilla/ejemplo"
    """
    name = f.name.lower()
    ext = f.suffix.lower()
    ext_score = {".xlsx": 50, ".xlsm": 40, ".xls": 20}.get(ext, 0)

    bonus = 0
    if any(k in name for k in ("formatocv", "formato", "cv", "edi", "expresion", "expresión", "exp_int", "expinteres")):
        bonus += 15
    if any(k in name for k in ("plantilla", "template", "blank", "ejemplo", "sample")):
        bonus -= 30

    # bonus leve si está en raíz del postulante (no en subcarpetas)
    # (esto se aplicará desde el caller)
    return ext_score + bonus


def score_pdf(f: Path, allow_bad_pdf: bool = False) -> int:
    name = f.name.lower()
    bonus = 0
    if any(k in name for k in ("formatocv", "cv", "expresion", "expresión", "edi", "exp_int", "expinteres")):
        bonus += 10
    if is_bad_pdf_name(name) and not allow_bad_pdf:
        bonus -= 80
    return bonus


def choose_best_file_for_postulante(post_dir: Path, allow_bad_pdf: bool = False) -> Tuple[Optional[Path], str]:
    """
    Retorna (path_elegido, motivo)
    motivo:
      - "OK_EXCEL"
      - "OK_PDF"
      - "SOLO_PDF_TIPO_CORREO"
      - "SIN_ARCHIVO_ELEGIBLE"
    """
    files: List[Path] = []
    for p in post_dir.rglob("*"):
        if not p.is_file():
            continue
        if p.name.startswith("~$"):
            continue
        if p.suffix.lower() not in ELIGIBLE_EXTS:
            continue
        files.append(p)

    if not files:
        return None, "SIN_ARCHIVO_ELEGIBLE"

    excels = [f for f in files if f.suffix.lower() in (".xlsx", ".xlsm", ".xls")]
    pdfs = [f for f in files if f.suffix.lower() == ".pdf"]

    # 1) Excel siempre primero
    if excels:
        scored = []
        for f in excels:
            s = score_excel(f)
            # bonus si está directamente en la carpeta del postulante
            if f.parent.resolve() == post_dir.resolve():
                s += 5
            scored.append((s, f))
        scored.sort(key=lambda t: (t[0], t[1].name.lower()), reverse=True)
        return scored[0][1], "OK_EXCEL"

    # 2) Si no hay excel, PDF
    if pdfs:
        scored = []
        for f in pdfs:
            s = score_pdf(f, allow_bad_pdf=allow_bad_pdf)
            if f.parent.resolve() == post_dir.resolve():
                s += 3
            scored.append((s, f))
        scored.sort(key=lambda t: (t[0], t[1].name.lower()), reverse=True)
        best = scored[0][1]
        if is_bad_pdf_name(best.name) and not allow_bad_pdf:
            return None, "SOLO_PDF_TIPO_CORREO"
        return best, "OK_PDF"

    return None, "SIN_ARCHIVO_ELEGIBLE"


# -------------------------
# Layout (Task 00) lectura para auditoría
# -------------------------
def load_layout_json(out_dir_011: Path) -> Optional[dict]:
    p = out_dir_011 / LAYOUT_FILE
    if not p.exists():
        return None
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return None


# -------------------------
# MAIN
# -------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", type=str, default="", help="Carpeta raíz: contiene los procesos")
    ap.add_argument("--only-proc", type=str, default="", help="Procesar solo procesos cuyo nombre contenga este texto")
    ap.add_argument("--dry-run", action="store_true", help="No escribe CSV/log, solo imprime resumen")
    ap.add_argument("--allow-bad-pdf", action="store_true",help="Permite elegir PDF tipo correo si no hay alternativa (NO recomendado salvo emergencia).")
    args = ap.parse_args()

    cfg = load_global_config()
    root = Path(args.root) if norm(args.root) else Path(cfg.get("input_root", ""))

    if not norm(str(root)):
        raise SystemExit("Debes pasar --root o definir input_root en configs/config.json")
    root = root.resolve()
    if not root.exists():
        raise SystemExit(f"No existe root: {root}")

    only_filter = norm(args.only_proc).lower()

    procesos = [p for p in root.iterdir() if p.is_dir()]
    procesos.sort(key=lambda x: x.name.lower())

    print(f"[task_10_collect_files] root = {root}")
    print(f"[task_10_collect_files] procesos detectados = {len(procesos)}")

    ok_proc = 0
    skip_proc = 0
    fail_proc = 0

    for proc_dir in procesos:
        proceso = proc_dir.name
        if only_filter and only_filter not in proceso.lower():
            continue

        in_dir_009 = find_009_dir(proc_dir)
        out_dir_011 = proc_dir / OUT_FOLDER_NAME

        if not in_dir_009:
            skip_proc += 1
            print(f"  - SKIP: {proceso} (no existe 009. EDI RECIBIDA(S))")
            continue

        # Carga config_layout.json si existe (para auditoría)
        layout = load_layout_json(out_dir_011)
        layout_warns: List[str] = []
        runtime_total_expected = None
        if layout is None:
            layout_warns.append("MISSING_config_layout_json")
        else:
            runtime = layout.get("runtime") or {}
            runtime_total_expected = safe_int(runtime.get("total_postulantes"), default=None) if runtime.get("total_postulantes") is not None else None

        # Postulantes = subcarpetas directas dentro de 009
        postulante_dirs = list_postulante_folders(in_dir_009)
        total_post_dirs = len(postulante_dirs)

        selected_rows: List[List[str]] = []
        skipped_rows: List[List[str]] = []
        manifest_rows: List[List[str]] = []

        chosen_count = 0
        skipped_count = 0

        # log
        if not args.dry_run:
            ensure_dir(out_dir_011)
            log_path = out_dir_011 / OUT_LOG
            if log_path.exists():
                log_path.unlink(missing_ok=True)
            log_append(log_path, f"== PROCESO: {proceso} ==")
            log_append(log_path, f"in_dir_009: {in_dir_009}")
            log_append(log_path, f"out_dir_011: {out_dir_011}")
            if layout_warns:
                for w in layout_warns:
                    log_append(log_path, f"[WARN] {w}")
            if runtime_total_expected is not None:
                log_append(log_path, f"runtime.total_postulantes (Task00): {runtime_total_expected}")
            log_append(log_path, f"carpetas_postulante_en_009: {total_post_dirs}")

        # Procesa postulantes
        for idx, post_dir in enumerate(postulante_dirs, start=1):
            chosen, reason = choose_best_file_for_postulante(post_dir, allow_bad_pdf=args.allow_bad_pdf)

            # manifest completo (inventario de elegibles)
            # (no es pesado: solo registra elegibles por postulante)
            eligibles = []
            for p in post_dir.rglob("*"):
                if p.is_file() and not p.name.startswith("~$") and p.suffix.lower() in ELIGIBLE_EXTS:
                    eligibles.append(p)
            eligibles.sort(key=lambda x: x.name.lower())
            if not eligibles:
                manifest_rows.append([post_dir.name, "", "", "", str(post_dir)])
            else:
                for f in eligibles:
                    ftype = "EXCEL" if f.suffix.lower() in (".xlsx", ".xlsm", ".xls") else "PDF"
                    manifest_rows.append([post_dir.name, f.name, ftype, str(f), str(post_dir)])

            if chosen is None:
                skipped_count += 1
                skipped_rows.append([post_dir.name, reason, str(post_dir)])
                if not args.dry_run:
                    log_append(out_dir_011 / OUT_LOG, f"[SKIP] {post_dir.name} | {reason}")
                continue

            chosen_count += 1
            ftype = "EXCEL" if chosen.suffix.lower() in (".xlsx", ".xlsm", ".xls") else "PDF"
            selected_rows.append([str(chosen_count), post_dir.name, chosen.name, ftype, str(chosen)])
            if not args.dry_run:
                log_append(out_dir_011 / OUT_LOG, f"[CHOSEN] {post_dir.name} -> {chosen.name} ({ftype})")

        # Validación: chosen + skipped == carpetas en 009
        if chosen_count + skipped_count != total_post_dirs:
            layout_warns.append("COUNT_MISMATCH_internal")  # debería no pasar

        # Validación con Task 00 (si existe)
        if runtime_total_expected is not None and runtime_total_expected != total_post_dirs:
            layout_warns.append(f"COUNT_MISMATCH_vs_Task00_expected={runtime_total_expected}_found={total_post_dirs}")

        # Dry run
        if args.dry_run:
            print(
                f"  - OK(dry): {proceso} | postulantes={total_post_dirs} | chosen={chosen_count} skipped={skipped_count} "
                f"| warnings={layout_warns if layout_warns else '[]'}"
            )
            ok_proc += 1
            continue

        # Escribe outputs
        ensure_dir(out_dir_011)

        write_csv(out_dir_011 / OUT_SELECTED,
                  ["n", "carpeta_postulante", "archivo", "tipo", "ruta"],
                  selected_rows)

        write_csv(out_dir_011 / OUT_SKIPPED,
                  ["carpeta_postulante", "motivo", "ruta_carpeta"],
                  skipped_rows)

        write_csv(out_dir_011 / OUT_MANIFEST,
                  ["carpeta_postulante", "archivo", "tipo", "ruta_archivo", "ruta_carpeta_postulante"],
                  manifest_rows)

        summary = {
            "generated_at": ts(),
            "process": proceso,
            "paths": {
                "process_dir": str(proc_dir),
                "in_dir_009": str(in_dir_009),
                "out_dir_011": str(out_dir_011),
                "layout_file": str((out_dir_011 / LAYOUT_FILE)),
            },
            "counts": {
                "postulantes_dirs_en_009": total_post_dirs,
                "chosen": chosen_count,
                "skipped": skipped_count,
            },
            "warnings": layout_warns,
            "notes": {
                "allow_bad_pdf": bool(args.allow_bad_pdf),
                "postulante_definition": "direct_subfolder_of_009",
                "selected_csv_schema": ["n", "carpeta_postulante", "archivo", "tipo", "ruta"],
            }
        }
        write_json(out_dir_011 / OUT_SUMMARY, summary)

        print(
            f"  - OK: {proceso} -> {OUT_SELECTED}({chosen_count}) / {OUT_SKIPPED}({skipped_count}) "
            f"| warnings={layout_warns if layout_warns else '[]'}"
        )
        ok_proc += 1

    print("")
    print(f"[task_10_collect_files] resumen: OK_PROCESOS={ok_proc} SKIP_PROCESOS={skip_proc} FAIL_PROCESOS={fail_proc}")


if __name__ == "__main__":
    main()
