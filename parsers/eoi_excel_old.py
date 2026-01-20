import re
from pathlib import Path
from datetime import datetime, date
from openpyxl import load_workbook

from utils.experience import to_date, total_days
import re
from pathlib import Path


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


def parse_eoi_excel_old(path: Path) -> dict:
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

# parsers/eoi_excel.py
# -*- coding: utf-8 -*-




def norm(s):
    return re.sub(r"\s+", " ", (s or "").strip())


def cell(ws, r, c):
    v = ws.cell(row=r, column=c).value
    return norm(str(v)) if v is not None else ""


def find_sheet(wb):
    # normalmente es la primera hoja
    return wb[wb.sheetnames[0]]


def parse_formacion_table(ws, start_row=47, end_row=56):
    """
    Tabla de formación académica:
    Col A: Nivel (Colegiatura, Maestría, etc)
    Col B: Especialidad
    Col C: Fecha de extensión
    Col D: Centro de estudios
    Col E: Ciudad / País
    """
    items = []

    for r in range(start_row, end_row + 1):
        nivel = cell(ws, r, 1)
        if not nivel:
            continue

        especialidad = cell(ws, r, 2)
        fecha = cell(ws, r, 3)
        centro = cell(ws, r, 4)
        ciudad_pais = cell(ws, r, 5)

        # Si toda la fila está vacía excepto el "nivel", igual lo guardamos (puede servir para evaluación)
        items.append({
            "nivel": nivel,
            "especialidad": especialidad,
            "fecha": fecha,
            "centro": centro,
            "ciudad_pais": ciudad_pais,
        })

    return items


def infer_grado_principal(formacion_items):
    """
    Deriva campos para llenar formato de revisión preliminar.
    Regla simple: si hay "Título Profesional" lleno -> titulo
                 si hay "Bachiller" lleno -> bachiller
                 si hay "Egresado Universitario" lleno -> egresado
    """
    titulo = ""
    bachiller = ""
    egresado = ""

    for it in formacion_items:
        nivel = (it.get("nivel") or "").upper()
        esp = it.get("especialidad") or ""
        centro = it.get("centro") or ""
        fecha = it.get("fecha") or ""
        cp = it.get("ciudad_pais") or ""

        detalle = " | ".join([x for x in [esp, centro, fecha, cp] if x])

        if "TITULO" in nivel and "PROF" in nivel and detalle:
            titulo = f"{it.get('nivel')} - {detalle}"
        if "BACHILLER" in nivel and detalle:
            bachiller = f"{it.get('nivel')} - {detalle}"
        if "EGRES" in nivel and "UNIV" in nivel and detalle:
            egresado = f"{it.get('nivel')} - {detalle}"

    return titulo, bachiller, egresado


def parse_eoi_excel(xlsx_path: Path) -> dict:
    wb = load_workbook(xlsx_path, data_only=True)
    ws = find_sheet(wb)

    data = {
        "source_file": str(xlsx_path),
        "dni": "",
        "nombre_full": "",
        "email": "",
        "celular": "",
        "titulo": "",
        "bachiller": "",
        "egresado": "",
        "formacion_academica": {
            "requisito_texto": "",
            "items": [],
        },
        "cursos": [],
        "experiencias": [],
        "exp_general_dias": 0,
        "exp_especifica_dias": 0,
        "java_ok": False,
        "oracle_ok": False,
    }

    # ==========================
    # I) DATOS PERSONALES (10-23)
    # ==========================
    # Aquí hay varias formas de extraer DNI, nombre, email, celular.
    # En tu formato, suelen estar en filas específicas, pero varía.
    # Vamos por estrategia híbrida:
    # - Leer bloque 10-23 como texto y buscar patrones.
    block_personal = " ".join(cell(ws, r, 1) + " " + cell(ws, r, 2) + " " + cell(ws, r, 3) for r in range(10, 24))
    block_personal = norm(block_personal)

    # DNI (8 dígitos)
    m_dni = re.search(r"\b(\d{8})\b", block_personal)
    if m_dni:
        data["dni"] = m_dni.group(1)

    # Email
    m_email = re.search(r"([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})", block_personal)
    if m_email:
        data["email"] = m_email.group(1)

    # Celular (buscamos 9 dígitos)
    m_cel = re.search(r"\b(9\d{8})\b", block_personal)
    if m_cel:
        data["celular"] = m_cel.group(1)

    # Nombre completo:
    # En tu formato suele ser: Apellido paterno, apellido materno, nombres
    # pero el Excel puede tenerlo repartido.
    # Para este avance, tomamos un fallback: nombre desde carpeta o lo deja vacío.
    # (si quieres exacto, lo ubicamos por celdas exactas cuando me confirmes filas/columnas)
    # data["nombre_full"] = ...

    # ==========================
    # II) CONDICIÓN LABORAL (25-37)
    # ==========================
    # Por ahora lo dejamos guardado como texto (útil luego para evaluación)
    block_laboral = " ".join(cell(ws, r, 1) + " " + cell(ws, r, 2) + " " + cell(ws, r, 3) for r in range(25, 38))
    data["condicion_laboral_texto"] = norm(block_laboral)

    # ==========================
    # III) FORMACIÓN ACADÉMICA (40+)
    # ==========================
    # Requisito en fila 45 (según tu especificación)
    # Puede estar en col A o B; lo leemos de A..E por seguridad
    req = " ".join([cell(ws, 45, c) for c in range(1, 6)])
    data["formacion_academica"]["requisito_texto"] = norm(req)

    # Tabla 47-56
    items = parse_formacion_table(ws, start_row=47, end_row=56)
    data["formacion_academica"]["items"] = items

    # Derivar título/bachiller/egresado para el formato de salida
    titulo, bachiller, egresado = infer_grado_principal(items)
    data["titulo"] = titulo
    data["bachiller"] = bachiller
    data["egresado"] = egresado

    return data
