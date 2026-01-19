
import re
from pathlib import Path
from datetime import datetime, date
from openpyxl import load_workbook

from utils.experience import to_date, total_days


def norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())


def read_cell(ws, addr: str) -> str:
    try:
        v = ws[addr].value
        return norm("" if v is None else str(v))
    except Exception:
        return ""


def extract_courses(ws):
    """
    El formato tiene bloques de cursos/capacitaciones repetibles.
    Estrategia robusta: recorre filas del bloque y toma textos en columna F (Capacitación).
    """
    courses = []
    for row in range(55, 90):
        cap = ws.cell(row=row, column=6).value  # F
        if cap:
            cap = norm(str(cap))
            up = cap.upper()
            # filtra rótulos comunes del formato
            if up and ("CAPACIT" not in up) and ("DESEABLE" not in up) and (up != "N°"):
                courses.append(cap)

    # dedupe conservando orden
    seen = set()
    out = []
    for c in courses:
        k = c.upper()
        if k not in seen:
            seen.add(k)
            out.append(c)
    return out


def extract_experience_records(ws, base_rows):
    records = []
    for r in base_rows:
        ent = norm(ws.cell(row=r, column=4).value or "")  # D
        proj = norm(ws.cell(row=r, column=6).value or "") # F
        cargo = norm(ws.cell(row=r, column=7).value or "")# G
        fi = to_date(ws.cell(row=r, column=8).value)       # H
        ff = to_date(ws.cell(row=r, column=9).value)       # I
        if ent or proj or cargo or fi or ff:
            records.append({"entidad": ent, "proyecto": proj, "cargo": cargo, "fi": fi, "ff": ff})
    return records


def parse_eoi_excel(path: Path) -> dict:
    wb = load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]  # "Formt Exp Int" en tu formato

    data = {
        "ap_paterno": read_cell(ws, "C13"),
        "ap_materno": read_cell(ws, "G13"),
        "nombres": read_cell(ws, "C15"),
        "dni": read_cell(ws, "H17"),
        "direccion": read_cell(ws, "C19"),
        "email": read_cell(ws, "F23"),
        "telefono": read_cell(ws, "C23"),
        "celular": read_cell(ws, "E23"),
        # Formación (bloque combinado: la esquina sup-izq guarda el valor)
        "titulo": read_cell(ws, "C47"),
        "bachiller": "",  # si deseas, mapear a otra celda si existe en tu formato
        "egresado": "",
    }

    # Cursos
    data["cursos"] = extract_courses(ws)

    # Experiencia general: filas base del formato
    exp_general_rows = [96, 102, 108, 114, 120]
    exps = extract_experience_records(ws, exp_general_rows)
    data["experiencias"] = exps

    # CALCULO EFECTIVO (SIN DUPLICAR SUPERPOSICIONES)
    intervals = [(rec["fi"], rec["ff"]) for rec in exps if rec.get("fi") and rec.get("ff")]
    gen_days = total_days(intervals, inclusive=True)
    data["exp_general_dias"] = int(gen_days)
    data["exp_general_anios"] = round(gen_days / 365.25, 2)

    # Experiencia específica: filtrado por keywords en cargo/proyecto (ajustable vía config, aquí heurístico)
    spec_intervals = []
    java_ok = False
    oracle_ok = False
    for rec in exps:
        txt = f"{rec.get('cargo','')} {rec.get('proyecto','')}".upper()
        if "JAVA" in txt:
            java_ok = True
        if "ORACLE" in txt:
            oracle_ok = True
        if any(k in txt for k in ("DESARROL", "PROGRAM", "ANALISTA", "SISTEM", "SOFTWARE")):
            if rec.get("fi") and rec.get("ff"):
                spec_intervals.append((rec["fi"], rec["ff"]))

    spec_days = total_days(spec_intervals, inclusive=True)
    data["exp_especifica_dias"] = int(spec_days)
    data["exp_especifica_anios"] = round(spec_days / 365.25, 2)
    data["java_ok"] = java_ok
    data["oracle_ok"] = oracle_ok

    data["nombre_full"] = norm(f"{data['nombres']} {data['ap_paterno']} {data['ap_materno']}")

    return data
