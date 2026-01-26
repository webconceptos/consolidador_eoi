# global_ejecuta.py
# -*- coding: utf-8 -*-

import argparse
import json
import os
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict, Any, List, Union

# -----------------------------
# Config general
# -----------------------------
OUT_011 = "011. INSTALACIÓN DE COMITÉ"
PROCESADOS = "procesados"

DEFAULT_TASKS_DIR = "tasks"

FILLED_PREFIX = "Cuadro_Evaluacion_LLENO_"
FILLED_EXT = ".xlsx"


# Estos nombres deben coincidir con tu repo actual:
TASK_LAYOUT = "task_00_layout_final.py"        # si existe
TASK_COLLECT_FILES = "task_10_collect_files.py"      # recopila archivos de postulantes
TASK_INIT_TEMPLATE = "task_15_init_cuadro_evaluacion.py"  # crea Cuadro_Evaluacion_<PROCESO>.xlsx
TASK_CRITERIA = "task_16_detect_criteria.py"           # genera criteria_evaluacion.json (antes task_05)
TASK_PARSE_POST = "task_20_parse_inputs.py"            # genera parsed_postulantes.jsonl
TASK_FILL_EVAL = "task_40_fill_cuadro_final.py"   # llena plantilla con postulantes
TASK_EVAL_LLM = "task_41_eval_postulantes_openai.py"   # eval FA/EC/etc (por ahora lo pausas, pero queda)

CFG_LAYOUT_NAME = "config_layout.json"
INIT_SUMMARY_NAME = "init_cuadro_summary.json"
PARSED_JSONL_NAME = "parsed_postulantes.jsonl"
CRITERIA_JSON_NAME = "criteria_evaluacion.json"
FILE_SELECTED_CSV = "files_selected.csv"
COLLECT_SUMMARY_JSON = "collect_summary.json"

# -----------------------------
# Helpers
# -----------------------------
def ts() -> str:
    return datetime.now().isoformat(timespec="seconds")


def log(msg: str):
    print(f"[global] {msg}")


def run_cmd(cmd: List[str], cwd: Optional[Path] = None, env: Optional[Dict[str, str]] = None,
            dry_run: bool = False, fail_fast: bool = True) -> int:
    """
    Ejecuta un comando con captura de stdout/err en tiempo real.
    """
    log(f"CMD => {' '.join(cmd)}")
    if dry_run:
        return 0

    p = subprocess.Popen(
        cmd,
        cwd=str(cwd) if cwd else None,
        env=env,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
    )

    assert p.stdout is not None
    for line in p.stdout:
        print(line.rstrip())

    rc = p.wait()
    if rc != 0 and fail_fast:
        raise SystemExit(f"[global] ERROR: comando falló (rc={rc}) => {' '.join(cmd)}")
    return rc

def run_cmd_m(cmd: List[str], cwd: Optional[Path] = None, env: Optional[Dict[str, str]] = None,
            dry_run: bool = False, fail_fast: bool = True) -> int:
    """
    Ejecuta un comando con captura de stdout/err en tiempo real.
    Si cmd apunta a un archivo .py, lo ejecuta como módulo:
        python -m paquete.modulo ...
    (para evitar problemas de imports entre carpetas)
    """
    def _log(msg: str) -> None:
        try:
            log(msg)  # si existe tu función log()
        except NameError:
            print(msg)

    def _to_module(py_file: str, project_root: Path) -> str:
        """
        Convierte ruta a .py dentro del project_root en ruta de módulo con puntos.
        Ej:
            project_root=/repo
            py_file="core/tasks/task_40.py"  -> "core.tasks.task_40"
        """
        p = Path(py_file)

        # Resolver a absoluto usando project_root/cwd
        if not p.is_absolute():
            p = (project_root / p).resolve()
        else:
            p = p.resolve()

        root = project_root.resolve()

        try:
            rel = p.relative_to(root)
        except ValueError:
            # Si no cuelga del root, igual lo intentamos con el padre del archivo
            # (último recurso, pero mejor fallar explícito para no ejecutar mal)
            raise ValueError(
                f"El script {p} no está dentro del project_root/cwd={root}. "
                f"Para usar python -m debe estar bajo el root del paquete."
            )

        if rel.suffix.lower() != ".py":
            raise ValueError(f"No es archivo .py: {rel}")

        # Quita .py y convierte separadores a '.'
        mod = ".".join(rel.with_suffix("").parts)

        # Edge-case: ejecutar __init__.py como módulo apunta al paquete
        if rel.name == "__init__.py":
            mod = ".".join(rel.parent.parts) if rel.parent.parts else ""
            if not mod:
                raise ValueError("No se puede ejecutar el __init__.py del root como módulo.")

        return mod

    if not cmd:
        raise ValueError("cmd no puede estar vacío")

    # Si ya viene con -m, no tocamos nada
    if "-m" not in cmd:
        # Asumimos formato usual: [python_exe, target, ...]
        if len(cmd) >= 2:
            python_exe = cmd[0]
            target = cmd[1]

            # Root del proyecto: si te pasan cwd, úsalo; si no, el directorio actual
            project_root = Path(cwd).resolve() if cwd else Path.cwd().resolve()

            # Detectar si target es ruta a archivo .py (o ruta con separadores)
            is_py = str(target).lower().endswith(".py")
            looks_like_path = ("/" in str(target)) or ("\\" in str(target))

            if is_py or looks_like_path:
                module = _to_module(str(target), project_root)
                # Reemplazamos: python path/to/file.py args...
                # por:          python -m paquete.modulo args...
                cmd = [python_exe, "-m", module, *cmd[2:]]
            else:
                # Si NO es ruta, asumimos que ya es módulo (core.tasks.task_40)
                cmd = [python_exe, "-m", target, *cmd[2:]]

    _log(f"CMD => {' '.join(cmd)}")
    if dry_run:
        return 0

    merged_env = os.environ.copy()
    if env:
        merged_env.update(env)

    p = subprocess.Popen(
        cmd,
        cwd=str(cwd) if cwd else None,
        env=merged_env,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
    )

    assert p.stdout is not None
    for line in p.stdout:
        print(line.rstrip())

    rc = p.wait()
    if rc != 0 and fail_fast:
        raise SystemExit(f"[global] ERROR: comando falló (rc={rc}) => {' '.join(cmd)}")
    return rc

def read_json(path: Path) -> Dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def exists_or_warn(p: Path, label: str, strict: bool) -> bool:
    if p.exists():
        return True
    msg = f"{label} no existe: {p}"
    if strict:
        raise SystemExit(f"[global] ERROR: {msg}")
    log(f"WARN: {msg}")
    return False


@dataclass
class ProcPaths:
    proc_dir: Path
    out_011: Path
    proc_out: Path
    cfg_layout: Path
    init_summary: Path
    parsed_jsonl: Path
    criteria_json: Path
    files_selected: Path 
    collect_summary: Path

    def as_debug(self) -> Dict[str, str]:
        return {
            "proc_dir": str(self.proc_dir),
            "out_011": str(self.out_011),
            "proc_out": str(self.proc_out),
            "cfg_layout": str(self.cfg_layout),
            "init_summary": str(self.init_summary),
            "parsed_jsonl": str(self.parsed_jsonl),
            "criteria_json": str(self.criteria_json),  
            "files_selected": str(self.files_selected),
            "collect_summary": str(self.collect_summary),
        }


def build_paths(proc_dir: Path) -> ProcPaths:
    out_011 = proc_dir / OUT_011
    proc_out = out_011 / PROCESADOS

    return ProcPaths(
        proc_dir=proc_dir,
        out_011=out_011,
        proc_out=proc_out,
        cfg_layout=out_011 / CFG_LAYOUT_NAME,
        init_summary=out_011 / INIT_SUMMARY_NAME,
        parsed_jsonl=out_011 / PARSED_JSONL_NAME,
        criteria_json=proc_out / CRITERIA_JSON_NAME,
        files_selected=out_011 / FILE_SELECTED_CSV,
        collect_summary=out_011 / COLLECT_SUMMARY_JSON,
    )


def find_process_dirs(root: Path, only_proc: str = "") -> List[Path]:
    if not root.exists():
        raise SystemExit(f"[global] ERROR: root no existe: {root}")

    procs = [p for p in root.iterdir() if p.is_dir()]
    procs.sort(key=lambda p: p.name.lower())

    if only_proc:
        procs = [p for p in procs if p.name == only_proc]
    return procs


def resolve_tasks_dir(root: Path, tasks_dir: str) -> Path:
    td = Path(tasks_dir)
    if td.is_dir():
        return td
    # fallback: relativo al root de ejecución
    td2 = (Path.cwd() / tasks_dir)
    if td2.is_dir():
        return td2
    # fallback: relativo al root de datos
    td3 = (root / tasks_dir)
    if td3.is_dir():
        return td3
    raise SystemExit(f"[global] ERROR: no encuentro tasks_dir: {tasks_dir}")


def script_path(tasks_dir: Path, name: str) -> Path:
    p = tasks_dir / name
    return p


def py() -> str:
    return sys.executable


def find_filled_excel(paths: ProcPaths) -> Optional[Path]:
    """
    Busca el Excel final lleno en 011/procesados:
    - preferencia exacta: Cuadro_Evaluacion_LLENO_<PROCESO>.xlsx
    - fallback: cualquier archivo que empiece con Cuadro_Evaluacion_LLENO_
    """
    target = paths.proc_out / f"{FILLED_PREFIX}{paths.proc_dir.name}{FILLED_EXT}"
    if target.exists():
        return target

    if not paths.proc_out.exists():
        return None

    cands = sorted(
        [p for p in paths.proc_out.glob(f"{FILLED_PREFIX}*{FILLED_EXT}") if p.is_file() and not p.name.startswith("~$")],
        key=lambda p: p.stat().st_mtime,
        reverse=True
    )
    return cands[0] if cands else None


# -----------------------------
# Pipeline Steps
# -----------------------------
def step_init_template(tasks_dir: Path, proc_dir: Path, dry_run: bool, debug: bool = False):
    """
    Task_15: crea el Cuadro_Evaluacion_<PROCESO>.xlsx dentro de 011.
    """
    script = script_path(tasks_dir, TASK_INIT_TEMPLATE)
    if not script.exists():
        raise SystemExit(f"[global] ERROR: no existe {script}")

    cmd = [py(), str(script), "--root", str(proc_dir.parent), "--only-proc", proc_dir.name]
    if debug:
        cmd.append("--debug")
    run_cmd(cmd, dry_run=dry_run)

def step_detect_layout(tasks_dir: Path, proc_dir: Path, dry_run: bool, debug: bool = False):
    """
    Task_00: genera config_layout.json (si lo estás usando).
    """
    script = script_path(tasks_dir, TASK_LAYOUT)
    if not script.exists():
        log(f"WARN: {TASK_LAYOUT} no existe, salto este paso.")
        return

    cmd = [py(), str(script), "--root", str(proc_dir.parent), "--only-proc", proc_dir.name]
    if debug:
        cmd.append("--debug")
    run_cmd(cmd, dry_run=dry_run)

def step_collect_files(tasks_dir: Path, proc_dir: Path, dry_run: bool, debug: bool, allow_bad_pdf: bool = False ):
    #script = tasks_dir / "task_10_collect_files.py"
    script = script_path(tasks_dir, TASK_COLLECT_FILES)
    cmd = [
        sys.executable, str(script),
        "--root", str(proc_dir.parent),         # root de procesos
        "--only-proc", proc_dir.name,           # filtra a este proceso
    ]
    if dry_run:
        cmd.append("--dry-run")
    if allow_bad_pdf:
        cmd.append("--allow-bad-pdf")

    run_cmd(cmd, dry_run=dry_run)

def step_parse_postulantes(tasks_dir: Path, proc_dir: Path, dry_run: bool, debug: bool):
    """
    Task_20: genera parsed_postulantes.jsonl dentro de 011/procesados.
    """
    script = script_path(tasks_dir, TASK_PARSE_POST)
    if not script.exists():
        raise SystemExit(f"[global] ERROR: no existe {script}")

    cmd = [py(), str(script), "--root", str(proc_dir.parent), "--only-proc", proc_dir.name]
    if debug:
        cmd.append("--debug")
    run_cmd_m(cmd, dry_run=dry_run)


def step_fill_cuadro(tasks_dir: Path, proc_dir: Path, dry_run: bool, debug: bool, limit: int = 0):
    """
    Task_40: llena el cuadro con los postulantes parseados.
    """
    script = script_path(tasks_dir, TASK_FILL_EVAL)
    if not script.exists():
        raise SystemExit(f"[global] ERROR: no existe {script}")

    cmd = [py(), str(script), "--root", str(proc_dir.parent), "--only-proc", proc_dir.name]
    if limit > 0:
        cmd += ["--limit", str(limit)]
    if debug:
        cmd.append("--debug")
    run_cmd_m(cmd, dry_run=dry_run)

def step_detect_criteria(tasks_dir: Path, proc_dir: Path, dry_run: bool, debug: bool):
    """
    Task_16: detecta criterios y genera criteria_evaluacion.json en 011/procesados.
    """
    script = script_path(tasks_dir, TASK_CRITERIA)
    if not script.exists():
        raise SystemExit(f"[global] ERROR: no existe {script}")

    cmd = [py(), str(script), "--root", str(proc_dir.parent), "--only-proc", proc_dir.name]
    if debug:
        cmd.append("--debug")
    run_cmd_m(cmd, dry_run=dry_run)

def step_eval_llm(tasks_dir: Path, paths: ProcPaths, dry_run: bool, debug: bool, limit: int = 1):
    """
    Task_41: evalúa con OpenAI (FA/EC...). Por ahora lo tienes en pausa,
    pero queda listo si lo reactivas.
    """
    script = script_path(tasks_dir, TASK_EVAL_LLM)
    if not script.exists():
        raise SystemExit(f"[global] ERROR: no existe {script}")

    cmd = [
        py(), str(script),
        "--criteria", str(paths.criteria_json),
        "--postulantes", str(paths.parsed_jsonl),
        "--limit", str(limit),
    ]
    if debug:
        cmd.append("--debug")
    run_cmd_m(cmd, dry_run=dry_run)


# -----------------------------
# Main Orchestrator
# -----------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True, help="Ruta raíz donde están los procesos (carpetas SCI...)")
    ap.add_argument("--only-proc", default="", help="Nombre exacto del proceso (opcional)")
    ap.add_argument("--tasks-dir", default=DEFAULT_TASKS_DIR, help="Carpeta donde están los tasks")
    ap.add_argument("--dry-run", action="store_true", help="No ejecuta, solo imprime comandos")
    ap.add_argument("--debug", action="store_true", help="Pasa flags debug a tasks (si existen)")

    # Control por pasos
    ap.add_argument("--do-init-template", action="store_true", help="Ejecuta task_15_init...")
    ap.add_argument("--do-layout", action="store_true", help="Ejecuta task_00_create_config_layout (si existe)")
    ap.add_argument("--do-collect", action="store_true", help="Ejecuta task_10_collect_files")
    ap.add_argument("--do-parse", action="store_true", help="Ejecuta task_20_parse_inputs")
    ap.add_argument("--do-fill", action="store_true", help="Ejecuta task_40_fill_cuadro_evaluacion")
    ap.add_argument("--do-criteria", action="store_true", help="Ejecuta task_16_detect_criteria")
    ap.add_argument("--do-eval", action="store_true", help="Ejecuta task_41_eval_postulantes_openai")
    ap.add_argument("--allow-bad-pdf", action="store_true", help="Permite PDFs corruptos en collect_files")
    ap.add_argument("--limit", type=int, default=0, help="Limita postulantes en fill/eval (0=no limita)")
    ap.add_argument("--strict", action="store_true", help="Falla si falta algún archivo esperado")

    args = ap.parse_args()
    hora_inicio = datetime.now()
    print(f"[global] Iniciando pipeline global: {ts()}")

    root = Path(args.root)
    tasks_dir = resolve_tasks_dir(root, args.tasks_dir)

    # Si no selecciona pasos explícitos => modo “pipeline mínimo”
    if not any([args.do_init_template, args.do_layout, args.do_parse, args.do_fill, args.do_criteria, args.do_eval]):
        # “Default sensato”: parse + fill + criteria (sin LLM)
        args.do_init_template = True
        args.do_layout = True
        args.do_collect = True  
        args.do_parse = True
        args.do_fill = True
        args.do_criteria = True
        args.do_eval = False

    procs = find_process_dirs(root, args.only_proc)
    if not procs:
        raise SystemExit("[global] ERROR: no hay procesos para ejecutar")

    log(f"Root: {root}")
    log(f"TasksDir: {tasks_dir}")
    log(f"Procesos: {len(procs)}")
    log(f"Pasos => init={args.do_init_template} layout={args.do_layout} parse={args.do_parse} fill={args.do_fill} criteria={args.do_criteria} eval={args.do_eval}")
    log(f"limit={args.limit} strict={args.strict} dry_run={args.dry_run}")

    for proc_dir in procs:
        paths = build_paths(proc_dir)
        log("-" * 80)
        log(f"PROCESO: {proc_dir.name}")
        log(f"Paths: {paths.as_debug()}")



        # 1) layout (config_layout.json) TASK_00
        if args.do_layout:
            print(f"[global] Ejecutando task_00_detect_layout para {proc_dir.name}")
            step_detect_layout(tasks_dir, proc_dir, args.dry_run, args.debug)
            print(f"[global] task_00_detect_layout completado para {proc_dir.name}")

        # 2) collect files (task_10) TASK_10
        if args.do_collect:
            print(f"[global] Ejecutando task_10_collect_files para {proc_dir.name}")
            step_collect_files(tasks_dir, proc_dir, args.dry_run, args.debug, allow_bad_pdf=args.allow_bad_pdf)
            print(f"[global] task_10_collect_files completado para {proc_dir.name}")

        # 3) init template (genera 011 y el excel) TASK_15
        if args.do_init_template:
            print(f"[global] Ejecutando task_15_init_template para {proc_dir.name}")
            step_init_template(tasks_dir, proc_dir, args.dry_run)
            print(f"[global] task_15_init_template completado para {proc_dir.name}")

        # 4) criteria TASK_16
        if args.do_criteria:
            print(f"[global] Ejecutando task_16_detect_criteria para {proc_dir.name}")
            step_detect_criteria(tasks_dir, proc_dir, args.dry_run, args.debug)
            print(f"[global] task_16_detect_criteria completado para {proc_dir.name}")

        # 5) parse postulantes TASK_20
        if args.do_parse:
            print(f"[global] Ejecutando task_20_parse_postulantes para {proc_dir.name}")
            step_parse_postulantes(tasks_dir, proc_dir, args.dry_run, args.debug)
            print(f"[global] task_20_parse_postulantes completado para {proc_dir.name}")

        # 6) fill cuadro TASK_40
        if args.do_fill:
            print(f"[global] Ejecutando task_40_fill_cuadro_final para {proc_dir.name}")
            step_fill_cuadro(tasks_dir, proc_dir, args.dry_run, args.debug, limit=args.limit)


            # Validación del Excel lleno (después de fill)
            if args.do_fill and not args.dry_run:
                filled = find_filled_excel(paths)
                if not filled:
                    msg = f"No se generó el Excel lleno esperado en {paths.proc_out} (prefijo: {FILLED_PREFIX})"
                    if args.strict:
                        raise SystemExit(f"[global] ERROR: {msg}")
                    log(f"WARN: {msg}")
                else:
                    log(f"Excel LLENO detectado: {filled.name}")

            # (Opcional) resumen rápido por proceso
                filled = find_filled_excel(paths)
                if filled: log(f"[global] OUTPUT => {filled}")
            
            print(f"[global] task_40_fill_cuadro_final completado para {proc_dir.name}")

        # Validaciones post
        if not args.dry_run:
            exists_or_warn(paths.out_011, "Carpeta 011", args.strict)
            exists_or_warn(paths.proc_out, "Carpeta 011/procesados", args.strict)

        if args.do_collect:
            exists_or_warn(paths.files_selected, "files_selected.csv", args.strict)
            exists_or_warn(paths.collect_summary, "collect_summary.json", False)

            exists_or_warn(paths.parsed_jsonl, "parsed_postulantes.jsonl", args.strict)
            exists_or_warn(paths.criteria_json, "criteria_evaluacion.json", args.strict)

        # 7) eval llm (opcional) TASK_41
        if args.do_eval:
            # por defecto: 1 postulante si no seteas limit
            lim = args.limit if args.limit > 0 else 1
            if not args.dry_run:
                # estas dos son indispensables
                exists_or_warn(paths.parsed_jsonl, "parsed_postulantes.jsonl", True)
                exists_or_warn(paths.criteria_json, "criteria_evaluacion.json", True)
            step_eval_llm(tasks_dir, paths, args.dry_run, args.debug, limit=lim)

    print(f"[global] Finalizando pipeline global: {ts()}")
    hora_fin = datetime.now()
    duracion = hora_fin - hora_inicio
    
    log(f"Duración total: {duracion}")
    log("-" * 80)
    log(" Pipeline global terminado.")


if __name__ == "__main__":
    main()
