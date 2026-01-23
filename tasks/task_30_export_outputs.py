
# tasks/task_30_export_outputs.py
# -*- coding: utf-8 -*-
"""
Task 30: export_outputs
"""

import argparse
import json
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, List

import sys
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))


from exporters.excel_exporter import export_consolidado_to_excel


OUT_FOLDER_NAME = "011. INSTALACIÓN DE COMITÉ"
IN_JSONL = "consolidado.jsonl"

OUT_XLSX = "consolidado_export.xlsx"


def ts():
    return datetime.now().isoformat(timespec="seconds")


def load_jsonl(path: Path) -> List[Dict[str, Any]]:
    items = []
    if not path.exists():
        return items
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line:
            continue
        items.append(json.loads(line))
    return items


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", type=str, default="", help="Carpeta raíz: contiene los procesos")
    args = ap.parse_args()

    root = Path(args.root) if args.root.strip() else None
    if not root or not root.exists():
        raise SystemExit("Debes pasar --root válido")

    procesos = [p for p in root.iterdir() if p.is_dir()]
    procesos.sort(key=lambda x: x.name.lower())

    ok = 0
    skip = 0

    print(f"[task_30] root={root} procesos={len(procesos)}")

    for proc_dir in procesos:
        proceso = proc_dir.name
        out_dir = proc_dir / OUT_FOLDER_NAME
        in_jsonl = out_dir / IN_JSONL

        if not in_jsonl.exists():
            print(f"  - SKIP: {proceso} (no existe {IN_JSONL})")
            skip += 1
            continue

        data = load_jsonl(in_jsonl)
        if not data:
            print(f"  - SKIP: {proceso} ({IN_JSONL} vacío)")
            skip += 1
            continue

        out_xlsx = out_dir / OUT_XLSX
        export_consolidado_to_excel(data, out_xlsx)

        ok += 1
        print(f"  - OK: {proceso} -> {OUT_XLSX} ({len(data)} filas)")

    print("")
    print(f"[task_30] resumen: OK={ok} SKIP={skip}")


if __name__ == "__main__":
    main()
