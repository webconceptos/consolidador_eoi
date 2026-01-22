
# tasks/task_10_collect_files.py
# -*- coding: utf-8 -*-
"""
Task 10: collect_files

Recolecta y elige 1 archivo por carpeta de postulante dentro de:
  <Proceso>/009. EDI RECIBIDA/<EDI postulante...>/(xlsx|xlsm|xls|pdf)

Regla:
- Si hay Excel => usar Excel (mejor score)
- Si NO hay Excel => usar PDF (mejor score), pero descartar PDFs tipo correo/presentación
- Genera reportes en:
  <Proceso>/011. INSTALACIÓN DE COMITÉ/
    - files_selected.csv
    - files_skipped.csv
    - debug_collect_files.log

Uso:
  python tasks/task_10_collect_files.py --root "D:\...\ProcesoSelección"
  python tasks/task_10_collect_files.py --dry-run
"""

import argparse
import csv
import json
import re
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Tuple

import sys
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))


IN_FOLDER_NAME = "009. EDI RECIBIDAS"
OUT_FOLDER_NAME = "011. INSTALACIÓN DE COMITÉ"

OUT_SELECTED = "files_selected.csv"
OUT_SKIPPED = "files_skipped.csv"
OUT_LOG = "debug_collect_files.log"


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


def save_csv(path: Path, rows: List[List[str]], header: List[str]):
    ensure_dir(path.parent)
    with path.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


def load_global_config() -> dict:
    p = Path("configs/config.json")
    if not p.exists():
        return {}
    return json.loads(p.read_text(encoding="utf-8"))


def is_bad_pdf_name(name: str) -> bool:
    n = name.lower()
    return any(k in n for k in ("correo", "presentacion", "presentación", "mail", "mensaje", "email"))


def score_excel(f: Path) -> int:
    name = f.name.lower()
    ext = f.suffix.lower()
    ext_score = {".xlsx": 30, ".xlsm": 20, ".xls": 10}.get(ext, 0)
    bonus = 0
    if any(k in name for k in ("formatocv", "formato", "cv", "edi", "expresion", "expresión")):
        bonus += 10
    if any(k in name for k in ("plantilla", "template", "blank", "ejemplo")):
        bonus -= 10
    return ext_score + bonus


def score_pdf(f: Path) -> int:
    name = f.name.lower()
    bonus = 0
    if any(k in name for k in ("formatocv", "cv", "expresion", "expresión", "edi")):
        bonus += 10
    if is_bad_pdf_name(name):
        bonus -= 50
    return bonus


def choose_one_file_per_postulante_folder(in_dir: Path):
    if not in_dir.exists():
        return [], []

    by_folder: Dict[Path, List[Path]] = {}
    for p in in_dir.rglob("*"):
        if not p.is_file():
            continue
        if p.name.startswith("~$"):
            continue
        ext = p.suffix.lower()
        if ext not in (".xlsx", ".xlsm", ".xls", ".pdf"):
            continue
        by_folder.setdefault(p.parent, []).append(p)

    chosen: List[Path] = []
    skipped: List[Tuple[Path, str]] = []

    for folder, files in by_folder.items():
        excels = [f for f in files if f.suffix.lower() in (".xlsx", ".xlsm", ".xls")]
        pdfs = [f for f in files if f.suffix.lower() == ".pdf"]

        if excels:
            excels.sort(key=score_excel, reverse=True)
            chosen.append(excels[0])
            continue

        if pdfs:
            pdfs.sort(key=score_pdf, reverse=True)
            best = pdfs[0]
            if is_bad_pdf_name(best.name):
                skipped.append((folder, "SOLO_PDF_TIPO_CORREO"))
                continue
            chosen.append(best)
        else:
            skipped.append((folder, "SIN_ARCHIVO_ELEGIBLE"))

    chosen.sort(key=lambda x: str(x).lower())
    return chosen, skipped


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", type=str, default="", help="Carpeta raíz: contiene los procesos")
    ap.add_argument("--dry-run", action="store_true", help="No escribe CSV/log, solo imprime resumen")
    args = ap.parse_args()

    cfg = load_global_config()
    root = Path(args.root) if args.root.strip() else Path(cfg.get("input_root", ""))

    if not str(root).strip():
        raise SystemExit("Debes pasar --root o definir input_root en configs/config.json")
    if not root.exists():
        raise SystemExit(f"No existe root: {root}")

    procesos = [p for p in root.iterdir() if p.is_dir()]
    procesos.sort(key=lambda x: x.name.lower())

    print(f"[collect_files] root = {root}")
    print(f"[collect_files] procesos detectados = {len(procesos)}")

    ok_proc = 0
    skip_proc = 0

    for proc_dir in procesos:
        proceso = proc_dir.name
        in_dir = proc_dir / IN_FOLDER_NAME
        out_dir = proc_dir / OUT_FOLDER_NAME

        if not in_dir.exists():
            skip_proc += 1
            print(f"  - SKIP: {proceso} (no existe 009)")
            continue

        chosen, skipped = choose_one_file_per_postulante_folder(in_dir)

        if args.dry_run:
            print(f"  - OK(dry): {proceso} | chosen={len(chosen)} skipped={len(skipped)}")
            ok_proc += 1
            continue

        ensure_dir(out_dir)
        log_path = out_dir / OUT_LOG
        if log_path.exists():
            log_path.unlink(missing_ok=True)

        log_append(log_path, f"== PROCESO: {proceso} ==")
        log_append(log_path, f"in_dir: {in_dir}")
        log_append(log_path, f"out_dir: {out_dir}")
        log_append(log_path, f"chosen: {len(chosen)} | skipped: {len(skipped)}")

        selected_rows: List[List[str]] = []
        for i, fp in enumerate(chosen, start=1):
            ftype = "EXCEL" if fp.suffix.lower() in (".xlsx", ".xlsm", ".xls") else "PDF"
            selected_rows.append([str(i), fp.parent.name, fp.name, ftype, str(fp)])
            log_append(log_path, f"[CHOSEN] {fp.parent.name} -> {fp.name} ({ftype})")

        skipped_rows: List[List[str]] = []
        for folder, reason in skipped:
            skipped_rows.append([folder.name, reason, str(folder)])
            log_append(log_path, f"[SKIP] {folder.name} | {reason}")

        save_csv(out_dir / OUT_SELECTED, selected_rows, ["n", "carpeta_postulante", "archivo", "tipo", "ruta"])
        save_csv(out_dir / OUT_SKIPPED, skipped_rows, ["carpeta_postulante", "motivo", "ruta_carpeta"])

        ok_proc += 1
        print(f"  - OK: {proceso} -> {OUT_SELECTED} ({len(chosen)} elegidos) / {OUT_SKIPPED} ({len(skipped)} omitidos)")

    print("")
    print(f"[collect_files] resumen: OK_PROCESOS={ok_proc} SKIP_PROCESOS={skip_proc}")


if __name__ == "__main__":
    main()
