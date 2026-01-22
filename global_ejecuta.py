
# global_ejecuta.py
# -*- coding: utf-8 -*-
"""
Ejecutor global de tasks.

Permite correr:
  task_00_config_layout
  task_10_collect_files
  task_20_parse_inputs
  task_30_export_outputs

Uso:
  python global_ejecuta.py --root "D:\...\ProcesoSelección" --from 0 --to 30
  python global_ejecuta.py --root "D:\...\ProcesoSelección" --only-proc "SCI N° 069"
  python global_ejecuta.py --root "D:\...\ProcesoSelección" --limit 5
"""

import argparse
import subprocess
import sys
from pathlib import Path


def run(cmd):
    print("")
    print(">>", " ".join(cmd))
    r = subprocess.run(cmd, check=False)
    if r.returncode != 0:
        raise SystemExit(r.returncode)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", type=str, required=True)
    ap.add_argument("--from", dest="from_task", type=int, default=0)
    ap.add_argument("--to", dest="to_task", type=int, default=30)
    ap.add_argument("--only-proc", type=str, default="")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--use-ocr", action="store_true")
    args = ap.parse_args()

    root = Path(args.root)
    if not root.exists():
        raise SystemExit(f"No existe root: {root}")

    py = sys.executable

    if args.from_task <= 0 <= args.to_task:
        cmd = [py, "tasks/task_00_config_layout.py", "--root", str(root)]
        if args.only_proc:
            cmd += ["--only-proc", args.only_proc]
        run(cmd)

    if args.from_task <= 10 <= args.to_task:
        cmd = [py, "tasks/task_10_collect_files.py", "--root", str(root)]
        if args.dry_run:
            cmd += ["--dry-run"]
        run(cmd)

    if args.from_task <= 20 <= args.to_task:
        cmd = [py, "tasks/task_20_parse_inputs.py", "--root", str(root)]
        if args.only_proc:
            cmd += ["--only-proc", args.only_proc]
        if args.limit and args.limit > 0:
            cmd += ["--limit", str(args.limit)]
        if args.use_ocr:
            cmd += ["--use-ocr"]
        run(cmd)

    if args.from_task <= 30 <= args.to_task:
        cmd = [py, "tasks/task_30_export_outputs.py", "--root", str(root)]
        run(cmd)

    print("\n[global_ejecuta] DONE")


if __name__ == "__main__":
    main()
