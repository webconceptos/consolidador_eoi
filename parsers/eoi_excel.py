# parsers/eoi_excel.py
# -*- coding: utf-8 -*-

import re
from pathlib import Path
from typing import Dict, Any, Optional, Tuple, List
from datetime import datetime, date
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

# -------------------------
# Utils
# -------------------------
def norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())

def _norm(s):
    return " ".join((str(s) if s is not None else "").strip().split())

def _cell(ws, r: int, c: int) -> str:
    v = ws.cell(row=r, column=c).value
    return norm(str(v)) if v is not None else ""


def _row_text(ws, r: int, c1: int = 1, c2: int = 10) -> str:
    parts = []
    for c in range(c1, c2 + 1):
        v = _cell(ws, r, c)
        if v:
            parts.append(v)
    return norm(" ".join(parts))


def _first_value_right(ws, r: int, c_start: int = 2, c_end: int = 12) -> str:
    """
    Devuelve el primer valor no vacío hacia la derecha en una fila (B..L por defecto).
    Útil cuando col A es etiqueta y el valor está en alguna col posterior.
    """
    for c in range(c_start, c_end + 1):
        v = _cell(ws, r, c)
        if v:
            return v
    return ""


def _find_in_row_regex(ws, r: int, pattern: str, c1: int = 1, c2: int = 12) -> str:
    """
    Busca un regex en el texto completo de la fila (A..L). Retorna primer match o "".
    """
    t = _row_text(ws, r, c1, c2)
    m = re.search(pattern, t)
    return m.group(1) if m else ""


def _match_label(label: str, *keys: str) -> bool:
    """
    True si todas las keys aparecen (en cualquier orden) dentro del label normalizado.
    """
    lb = label.lower()
    return all(k.lower() in lb for k in keys)

def _as_date_str(v):
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.date().strftime("%d/%m/%Y")
    if isinstance(v, date):
        return v.strftime("%d/%m/%Y")
    return _norm(v)

def _looks_like_yes(s: str) -> bool:
    s = _norm(s).upper()
    if not s:
        return False
    # heurísticas típicas en estas EDI
    return s in ("SI", "SÍ", "X", "OK", "1", "TRUE") or "SI" == s

# -------------------------
# Parsing: Datos Personales (tabla filas 12-23)
# -------------------------
def parse_datos_personales(ws, start_row=12, end_row=23, max_cols=12, debug=False):
    """
    Datos personales en estructura de 2 filas:
      - fila impar: VALORES
      - fila par: ENCABEZADOS
    Ej:
      12 headers: Apellido Paterno | Apellido Materno
      13 values : García Monterroso | Ramírez
      14 headers: Nombres | Lugar... | Día | Mes | Año
      15 values : Enrique Arturo | Piura | 7 | 3 | 1983
      ...
      22 headers: Teléfono | Celular | email
      23 values : 902... | 9........ | correo@...
    """

    def get_row_cells(r):
        return [_cell(ws, r, c) for c in range(1, max_cols + 1)]

    out = {
        "dni": "",
        "apellido_paterno": "",
        "apellido_materno": "",
        "nombres": "",
        "nombre_full": "",
        "email": "",
        "celular": "",
    }

    # Recorremos de 2 en 2: header_row, value_row
    r = start_row
    while r + 1 <= end_row:
        header_row = r
        value_row = r + 1

        headers = get_row_cells(header_row)
        values = get_row_cells(value_row)

        if debug:
            print(f"[DP2] header_row={header_row} => {headers}")
            print(f"[DP2] value_row ={value_row} => {values}")

        # construye mapa header->value por columna, ignorando vacíos
        for h, v in zip(headers, values):
            h_low = (h or "").lower()
            v = v or ""

            if not h_low:
                continue

            # Apellidos
            if "apellido paterno" in h_low:
                out["apellido_paterno"] = v
            elif "apellido materno" in h_low:
                out["apellido_materno"] = v

            # Nombres
            elif re.search(r"\bnombres\b", h_low):
                # evita que agarre "Lugar de nacimiento" como nombres (en tu ejemplo estaba en headers)
                if "lugar" not in v.lower():
                    out["nombres"] = v

            # DNI
            elif ("documento" in h_low and "identidad" in h_low) or "dni" in h_low:
                m = re.search(r"\b(\d{8})\b", v)
                if not m:
                    # a veces viene pegado en fila completa
                    m = re.search(r"\b(\d{8})\b", _row_text(ws, value_row, 1, max_cols))
                if m:
                    out["dni"] = m.group(1)

            # Celular
            elif "celular" in h_low:
                vv = re.sub(r"\D", "", v)
                if re.fullmatch(r"9\d{8}", vv):
                    out["celular"] = vv
                else:
                    # fallback: buscar en toda la fila de valores
                    m = re.search(r"\b(9\d{8})\b", _row_text(ws, value_row, 1, max_cols))
                    if m:
                        out["celular"] = m.group(1)

            # Email
            elif "email" in h_low or "correo" in h_low:
                m = re.search(r"([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})", v)
                if not m:
                    m = re.search(r"([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})", _row_text(ws, value_row, 1, max_cols))
                if m:
                    out["email"] = m.group(1)

        r += 2

    # nombre_full
    ap = " ".join([out["apellido_paterno"], out["apellido_materno"]]).strip()
    nm = out["nombres"].strip()
    out["nombre_full"] = norm(" ".join([ap, nm]))
    
    if debug:
        print("[DP2] RESULT =>", out["dni"], out["nombre_full"], out["celular"], out["email"])

    return out

def parse_formacion_academica_excel_old(ws, debug=False):
    """
    Formación Académica según tu plantilla real:
      - Encabezado: fila 49
      - Datos: filas 51..56
      - Columnas:
          B: Título*
          D: Especialidad
          F: Fecha Extensión
          H: Centro de Estudios
          J: Ciudad/País
    """
    start_row = 51
    end_row = 56
    
    print("iniciando FA")

    items = []
    for r in range(start_row, end_row + 1):
        titulo = _norm(ws.cell(row=r, column=2).value)  # B
        especialidad = _norm(ws.cell(row=r, column=4).value)  # D
        fecha = _as_date_str(ws.cell(row=r, column=6).value)  # F
        centro = _norm(ws.cell(row=r, column=8).value)  # H
        ciudad = _norm(ws.cell(row=r, column=10).value)  # J

        # Si no hay título en B, esta fila no es válida
        if not titulo:
            continue

        has_data = any([especialidad, fecha, centro, ciudad])

        item = {
            "row": r,
            "titulo_item": titulo,        # COLEGIATURA / TITULO / BACHILLER...
            "especialidad": especialidad,
            "fecha": fecha,
            "centro": centro,
            "ciudad": ciudad,
            "has_data": has_data
        }
        items.append(item)

        print(item)
        
        if 1==1:
            print(f"[FA] r={r} titulo='{titulo}' esp='{especialidad}' fecha='{fecha}' centro='{centro}' ciudad='{ciudad}' has={has_data}")

    print("Acabando el FOR")

    # Construimos un resumen compacto solo con filas llenas
    picked = [x for x in items if x["has_data"]]

    parts = []
    for it in picked:
        p = f"{it['titulo_item']}: {it['especialidad']}".strip()
        extras = []
        if it["fecha"]: extras.append(it["fecha"])
        if it["centro"]: extras.append(it["centro"])
        if it["ciudad"]: extras.append(it["ciudad"])
        if extras:
            p += " (" + " | ".join(extras) + ")"
        parts.append(p)

    resumen = " ; ".join(parts) if parts else ""
    
    print(resumen)

    return {
        "items": items,
        "resumen": resumen
    }

def parse_formacion_academica_excel(ws, debug=False):
    """
    Busca dinámicamente la tabla de Formación Académica por el encabezado:
      - 'Fecha de Extensión del Título' o 'Centro de Estudios' o 'Ciudad/ País'
    Luego lee filas de items: COLEGIATURA, MAESTRIA, EGRESADO..., TITULO, BACHILLER, EGRESADO UNIVERSITARIO
    """

    # 1) localizar la fila del header (donde aparece "Fecha de Extensión..." etc.)
    header_row = None
    # buscamos en un rango razonable
    for r in range(40, 70):
        row_text = " ".join([_norm(ws.cell(row=r, column=c).value) for c in range(1, 15)]).upper()
        if ("FECHA DE EXTENS" in row_text and "TITULO" in row_text) or ("CENTRO DE ESTUDIOS" in row_text) or ("CIUDAD/ PA" in row_text or "CIUDAD/PA" in row_text):
            header_row = r
            break

    if header_row is None:
        print("[FA] No se encontró header de tabla (40-70).")
        return {"items": [], "resumen": ""}

    # 2) las filas de datos empiezan típicamente 2 filas abajo del header
    # (en tu captura: header 49, datos 51)
    start_row = header_row + 2
    end_row = start_row + 10  # margen por si agregan filas

    # columnas esperadas según tu formato (B,D,F,H,J) (CFGHJ)
    col_titulo = 3
    col_esp = 6
    col_fecha = 7
    col_centro = 8
    col_ciudad = 10

    # keywords de filas esperadas (si el postulante dejó vacío, igual deben estar)
    expected_titles = ("COLEGIATURA", "MAESTRIA", "EGRESADO", "TITULO", "BACHILLER", "UNIVERSITARIO")

    items = []
    for r in range(start_row, end_row + 1):
        titulo = _norm(ws.cell(row=r, column=col_titulo).value)

        # si no hay en B, probamos si está en A (por merges o desplazamientos)
        if not titulo:
            titulo = _norm(ws.cell(row=r, column=1).value)

        titulo_up = titulo.upper()

        # solo consideramos filas que parezcan parte de la tabla
        if not any(k in titulo_up for k in expected_titles):
            continue

        especialidad = _norm(ws.cell(row=r, column=col_esp).value)
        fecha = _as_date_str(ws.cell(row=r, column=col_fecha).value)
        centro = _norm(ws.cell(row=r, column=col_centro).value)
        ciudad = _norm(ws.cell(row=r, column=col_ciudad).value)

        has_data = any([especialidad, fecha, centro, ciudad])

        item = {
            "row": r,
            "titulo_item": titulo,
            "especialidad": especialidad,
            "fecha": fecha,
            "centro": centro,
            "ciudad": ciudad,
            "has_data": has_data
        }
        items.append(item)

        if debug:
            print(f"[FA] r={r} titulo='{titulo}' esp='{especialidad}' fecha='{fecha}' centro='{centro}' ciudad='{ciudad}' has={has_data}")

    picked = [x for x in items if x["has_data"]]
    parts = []
    for it in picked:
        p = f"{it['titulo_item']}: {it['especialidad']}".strip()
        extras = []
        if it["fecha"]: extras.append(it["fecha"])
        if it["centro"]: extras.append(it["centro"])
        if it["ciudad"]: extras.append(it["ciudad"])
        if extras:
            p += " (" + " | ".join(extras) + ")"
        parts.append(p)

    resumen = " ; ".join(parts) if parts else ""

    return {"items": items, "resumen": resumen}

def _norm(x):
    return re.sub(r"\s+", " ", str(x).strip()) if x is not None else ""

def _as_date_str(v):
    if v is None or str(v).strip() == "":
        return ""
    if isinstance(v, (datetime, date)):
        return v.strftime("%d/%m/%Y")
    s = _norm(v)
    # deja tal cual si ya viene como texto dd/mm/aaaa u otro
    return s

def _as_int(v):
    if v is None or str(v).strip() == "":
        return 0
    try:
        # puede venir 12.0
        return int(float(str(v).replace(",", ".").strip()))
    except Exception:
        return 0

def is_start_experiencia_row(ws, r: int) -> bool:
    # Revisamos varias columnas por seguridad (A..K)
    txt = " ".join(str(ws.cell(r, c).value or "") for c in range(1, 12)).upper()

    # Caso 1: título de sección
    if "IV" in txt and "EXPERIENCIA" in txt:
        return True

    # Caso 2: cabecera típica de experiencia
    if ("NOMBRE DE LA ENTIDAD" in txt or "NOMBRE DE LA EMPRESA" in txt) and "FECHA DE INICIO" in txt:
        return True
    if "NOMBRE DEL PROYECTO" in txt and "FECHA DE CULMINACIÓN" in txt:
        return True

    return False

def is_bad_course_row(item: dict) -> bool:
    nro = (item.get("nro") or "").strip().upper()
    centro = (item.get("centro") or "").strip().upper()
    cap = (item.get("capacitacion") or "").strip().upper()
    fi = (item.get("fecha_inicio") or "").strip().upper()
    ff = (item.get("fecha_fin") or "").strip().upper()

    # encabezados típicos (cursos o experiencia)
    if nro in ("NO.", "NRO", "N°"):
        return True
    if "NOMBRE DE LA ENTIDAD" in centro or "NOMBRE DEL PROYECTO" in cap:
        return True
    if fi == "FECHA DE INICIO" or ff == "FECHA DE CULMINACIÓN":
        return True

    # si no tiene centro/cap y tampoco horas/fechas, no sirve
    if not (centro or cap):
        return True

    return False

def _is_header_like(nro, centro, cap, fi, ff):
    s1 = (nro or "").strip().lower()
    s2 = (centro or "").strip().lower()
    s3 = (cap or "").strip().lower()
    if s1 in ("no.", "n°", "nº"): 
        return True
    if "centro de estudios" in s2 and ("capacit" in s3 or "capacitación" in s3):
        return True
    if (fi or "").strip().lower().startswith("fecha"):
        return True
    if (ff or "").strip().lower().startswith("fecha"):
        return True
    return False

def parse_estudios_complementarios_excel_old(ws, debug=False):
    """
    Detecta bloques dinámicos b.1, b.2, b.3... y extrae items con columnas:
      N° -> C(3), Centro -> D(4), Capacitación -> F(6),
      Fecha Inicio -> H(8), Fecha Fin -> I(9), Horas -> J(10)
    Regla: data comienza 4 filas debajo de la cabecera del bloque.
    """
    # --- 1) localizar filas de cabecera de bloque (b.1, b.2, b.3, ...)
    block_rows = []
    max_scan = ws.max_row or 200

    # patrón flexible: "b.1", "B.1", "b1", etc.
    #pat = re.compile(r"^\s*b\s*\.?\s*(\d+)\s*$", re.IGNORECASE)
    pat = re.compile(r"\bb\s*\.?\s*(\d+)\b", re.IGNORECASE)

    #print("Patron a buscar" + pat)

    for r in range(1, max_scan + 1):
        a = _norm(ws.cell(row=r, column=3).value)  # columna CA
        if not a:
            continue
        #m = pat.match(a)
        m = pat.search(a)
        #print ("Valor de m: " + m) 
        if m:
            idx = int(m.group(1))
            # título suele estar en B/C... tomamos la fila completa para debug
            title = _norm(ws.cell(row=r, column=2).value) or _norm(ws.cell(row=r, column=3).value)
            #print("Titulo" + title)
            block_rows.append((idx, r, title))

    block_rows.sort(key=lambda x: x[1])
    if debug:
        print("[EC] bloques detectados:", block_rows)

    blocks = []
    total_horas = 0

    # --- 2) para cada bloque, define rango desde cabecera hasta antes del siguiente bloque
    for i, (idx, header_row, title) in enumerate(block_rows):
        next_header_row = block_rows[i + 1][1] if i + 1 < len(block_rows) else (max_scan + 1)

        # data comienza 4 filas debajo de cabecera
        data_start = header_row + 4
        data_end = next_header_row - 1

        # si no hay siguiente bloque, no uses max_scan directo:
        data_end = max_scan

        # corta si aparece IV. EXPERIENCIA
        for rr in range(data_start, max_scan + 1):
            if is_start_experiencia_row(ws, rr):
                data_end = rr - 1
                break


        items = []
        horas_block = 0

        if debug:
            print(f"[EC] b.{idx} header_row={header_row} data_start={data_start} data_end={data_end}")

        # --- 3) iterar filas de data
        for r in range(data_start, data_end + 1):
            nro = _norm(ws.cell(row=r, column=3).value)      # C
            centro = _norm(ws.cell(row=r, column=4).value)   # D
            cap = _norm(ws.cell(row=r, column=6).value)      # F
            fi = _as_date_str(ws.cell(row=r, column=8).value) # H
            ff = _as_date_str(ws.cell(row=r, column=9).value) # I
            horas = _as_int(ws.cell(row=r, column=10).value)  # J

            # condición de fila válida: al menos centro o capacitación
            if not (centro or cap):
                # ojo: no hacemos "break" porque puede haber filas vacías intermedias
                continue

            item = {
                "row": r,
                "nro": nro,
                "centro": centro,
                "capacitacion": cap,
                "fecha_inicio": fi,
                "fecha_fin": ff,
                "horas": horas,
            }
            
            items.append(item)
            if is_bad_course_row(item):
                continue
            items.append(item)

            horas_block += horas

            if debug:
                print(f"[EC]  r={r} nro='{nro}' centro='{centro}' cap='{cap}' fi='{fi}' ff='{ff}' horas={horas}")

        total_horas += horas_block

        blocks.append({
            "id": f"b.{idx}",
            "row": header_row,
            "title": title,
            "items": items,
            "total_horas": horas_block,
            "resumen": ""  # lo llenamos abajo
        })

    # --- 4) resumen por bloque (multilínea)
    def format_course_line(x):
        left = " - ".join([p for p in [x.get("centro",""), x.get("capacitacion","")] if _norm(p)])
        extras = " | ".join([p for p in [x.get("fecha_inicio",""), x.get("fecha_fin","")] if _norm(p)])
        h = x.get("horas", 0) or 0
        if extras and h:
            return f"{left} ({extras} | {h}h)"
        if extras:
            return f"{left} ({extras})"
        if h:
            return f"{left} ({h}h)"
        return left

    for b in blocks:
        lines = [format_course_line(x) for x in (b.get("items", []) or []) if (_norm(x.get("centro","")) or _norm(x.get("capacitacion","")))]
        b["resumen"] = "\n".join(lines).strip()

    resumen_global = "\n\n".join([f"{b['id'].upper()}:\n{b['resumen']}" for b in blocks if _norm(b.get("resumen",""))]).strip()
    
    print("RESUMEN GLOBAL")
    print(resumen_global)

    return {
        "blocks": blocks,
        "total_horas": total_horas,
        "resumen": resumen_global
    }

def parse_estudios_complementarios_excel(ws: Worksheet, debug: bool = False):
    """
    Estudios Complementarios (Sección II en tu formato EDI):
      - Bloques dinámicos: b.1, b.2, b.3, b.4, ... (hasta antes de IV. EXPERIENCIA)
      - Cada bloque: los datos comienzan 4 filas después de la fila del título del bloque.
      - Columnas (según tu plantilla real):
          C: N°
          D: Centro de estudios
          F: Capacitación
          H: Fecha de Inicio
          I: Fecha de Culminación
          J: Horas
    """
    def norm(x):
        if x is None:
            return ""
        s = str(x).replace("\u00a0", " ").strip()
        return re.sub(r"\s+", " ", s)

    def as_date_str(v):
        # deja pasar strings ya formateados; si es datetime/date, conviértelo (si ya tienes helper, úsalo)
        try:
            import datetime
            if isinstance(v, (datetime.datetime, datetime.date)):
                return v.strftime("%d/%m/%Y")
        except Exception:
            pass
        return norm(v)

    def is_section_iv(row_vals):
        txt = " ".join([norm(x) for x in row_vals if norm(x)])
        up = txt.upper()
        # Corta cuando empieza sección IV EXPERIENCIA
        return ("IV." in up and "EXPERI" in up) or up.startswith("IV") and "EXPERI" in up

    # Detecta "b.1", "b.1)", "b.1) texto..."
    pat_block = re.compile(r"^\s*b\s*\.?\s*(\d+)\s*\)?\s*", re.IGNORECASE)

    max_row = ws.max_row or 0
    blocks = []
    # 1) localizar cabeceras b.x
    for r in range(1, max_row + 1):
        # revisa columnas A..J (1..10) para encontrar "b.1)"
        row_vals = [ws.cell(row=r, column=c).value for c in range(1, 11)]
        if is_section_iv(row_vals):
            if debug:
                print(f"[EC] corte por sección IV en fila {r}")
            break

        for c in range(1, 11):
            v = norm(ws.cell(row=r, column=c).value)
            if not v:
                continue
            m = pat_block.match(v)
            if m:
                idx = int(m.group(1))
                block_id = f"b.{idx}"
                title = v  # o guarda toda la fila si quieres
                blocks.append({"id": block_id, "row": r, "title": title, "items": [], "resumen": "", "total_horas": 0})
                if debug:
                    print(f"[EC] detectado {block_id} en fila {r} col {c}: {title}")
                break  # no sigas revisando otras cols en la misma fila

    if not blocks:
        return {"blocks": [], "total_horas": 0, "resumen": ""}

    # 2) por cada bloque, leer items hasta antes del siguiente bloque o sección IV
    total_horas_global = 0
    for i, b in enumerate(blocks):
        header_row = b["row"]
        next_header_row = blocks[i + 1]["row"] if i + 1 < len(blocks) else None

        data_start = header_row + 4
        data_end = (next_header_row - 1) if next_header_row else max_row

        if debug:
            print(f"[EC] {b['id']} header_row={header_row} data_start={data_start} data_end={data_end}")

        seen = set()  # dedupe por (row,nro,centro,cap,fi,ff,horas)

        for r in range(data_start, data_end + 1):
            row_vals = [ws.cell(row=r, column=c).value for c in range(1, 11)]
            if is_section_iv(row_vals):
                if debug:
                    print(f"[EC] corte por sección IV dentro de {b['id']} en fila {r}")
                break

            nro = norm(ws.cell(row=r, column=3).value)     # C
            centro = norm(ws.cell(row=r, column=4).value)  # D
            cap = norm(ws.cell(row=r, column=6).value)     # F
            fi = as_date_str(ws.cell(row=r, column=8).value)  # H
            ff = as_date_str(ws.cell(row=r, column=9).value)  # I
            horas_raw = ws.cell(row=r, column=10).value       # J

            # limpia "Horas"
            horas = 0
            if horas_raw is not None and str(horas_raw).strip() != "":
                try:
                    horas = int(float(str(horas_raw).strip()))
                except Exception:
                    horas = 0

            # Ignorar filas de encabezado metidas en medio
            if nro.upper() in ("NO.", "N°", "NRO", "NRO.") or centro.upper().startswith("CENTRO DE ESTUDIOS"):
                continue
            if cap.upper() in ("CAPACITACIÓN", "CAPACITACION") and fi.upper().startswith("FECHA"):
                continue

            # fila vacía real
            if not (centro or cap):
                # si también está vacío nro/fechas/horas, pasa
                if not (nro or fi or ff or horas):
                    continue

            key = (r, nro, centro, cap, fi, ff, horas)
            if key in seen:
                continue
            seen.add(key)

            item = {
                "row": r,
                "nro": nro,
                "centro": centro,
                "capacitacion": cap,
                "fecha_inicio": fi,
                "fecha_fin": ff,
                "horas": horas,
            }
            b["items"].append(item)
            b["total_horas"] += horas

            if debug:
                print(f"[EC]  r={r} nro='{nro}' centro='{centro}' cap='{cap}' fi='{fi}' ff='{ff}' horas={horas}")

        total_horas_global += b["total_horas"]

        # resumen por bloque (multilínea)
        lines = []
        for it in b["items"]:
            if not (it["centro"] or it["capacitacion"]):
                continue
            line = " - ".join([x for x in [it["centro"], it["capacitacion"]] if x])
            extras = " | ".join([x for x in [it["fecha_inicio"], it["fecha_fin"]] if x])
            if it["horas"]:
                extras = (extras + " | " if extras else "") + f"{it['horas']}h"
            if extras:
                line = f"{line} ({extras})"
            lines.append(line)

        b["resumen"] = "\n".join(lines).strip()

    # resumen general
    resumen_parts = []
    for b in blocks:
        etiqueta = b["id"].upper()
        if b["resumen"]:
            resumen_parts.append(f"{etiqueta}:\n{b['resumen']}")
        else:
            resumen_parts.append(f"{etiqueta}:\n(sin cursos declarados)")

    return {
        "blocks": blocks,
        "total_horas": total_horas_global,
        "resumen": "\n\n".join(resumen_parts).strip(),
    }



# Reutiliza tu norm()
def norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())

def _parse_date_any(s: str) -> Optional[datetime]:
    """
    Acepta dd/mm/yyyy o dd-mm-yyyy. Devuelve datetime o None.
    """
    s = norm(s).replace("-", "/")
    if not s:
        return None
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", s)
    if not m:
        return None
    d, mth, y = map(int, m.groups())
    try:
        return datetime(y, mth, d)
    except Exception:
        return None

def _days_between(fi: Optional[datetime], ff: Optional[datetime]) -> int:
    if not fi or not ff:
        return 0
    delta = (ff - fi).days
    return max(delta, 0)

def _looks_like_header_row(values: List[str]) -> bool:
    """
    Heurística: si la fila tiene 'ENTIDAD', 'CARGO', 'FECHA' etc, es cabecera.
    """
    t = " ".join(v.upper() for v in values if v).strip()
    if not t:
        return False
    patterns = [
        r"\bENTIDAD\b", r"\bINSTITU", r"\bEMPRESA\b",
        r"\bCARGO\b", r"\bPUESTO\b",
        r"\bFECHA\b", r"\bINICIO\b", r"\bFIN\b",
        r"\bDESCRIP", r"\bFUNCION"
    ]
    hits = sum(1 for p in patterns if re.search(p, t))
    return hits >= 2

import re
from datetime import datetime
from typing import Dict, Any, List, Optional

def norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())

def _parse_date_any(s: str) -> Optional[datetime]:
    s = norm(s).replace("-", "/")
    if not s:
        return None
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", s)
    if not m:
        return None
    d, mo, y = map(int, m.groups())
    try:
        return datetime(y, mo, d)
    except Exception:
        return None

def _days_between(fi: Optional[datetime], ff: Optional[datetime]) -> int:
    if not fi or not ff:
        return 0
    return max((ff - fi).days, 0)

def _row_text(ws, row: int, col_from: int, col_to: int) -> str:
    parts = []
    for c in range(col_from, col_to + 1):
        v = ws.cell(row=row, column=c).value
        if v is None:
            continue
        parts.append(norm(str(v)))
    return " ".join([p for p in parts if p]).strip()

def _is_desc_label_row(ws, row: int) -> bool:
    # buscamos el rótulo aunque esté en una celda combinada
    t = _row_text(ws, row, 1, 12).upper()
    return ("DESCRIP" in t) and ("TRABAJO" in t) and ("REALIZ" in t)

def _get_desc_detail(ws, row: int) -> str:
    """
    La descripción detalle suele estar en una celda combinada.
    Estrategia robusta:
    - tomar el texto más "largo" de columnas A..J (o A..L) en esa fila
    """
    best = ""
    for c in range(1, 13):
        v = ws.cell(row=row, column=c).value
        s = norm(str(v)) if v is not None else ""
        if len(s) > len(best):
            best = s
    return best

def parse_experiencia_general_excel(
    ws,
    start_row: int = 101,
    end_row: int = 145,
    debug: bool = False
) -> Dict[str, Any]:
    """
    Experiencia General (EDI) - FORMATO REAL (bloques con descripción):
    - Fila datos: C..J
      C: nro
      D/E: entidad
      F: proyecto
      G: cargo/servicio
      H: fecha_inicio
      I: fecha_fin
      J: tiempo_en_cargo (texto)
    - Luego:
      + fila con texto "Descripción del trabajo Realizado:"
      + fila siguiente con el detalle (celda combinada)
    (El encabezado se repite por cada experiencia)
    """

    COL_C = 3  # C
    items: List[Dict[str, Any]] = []
    lines: List[str] = []
    seen = set()

    r = start_row
    while r <= end_row:
        # --- 1) leer fila base C..J
        nro = norm(str(ws.cell(r, COL_C + 0).value or ""))

        entidad = " ".join([
            norm(str(ws.cell(r, COL_C + 1).value or "")),  # D
            norm(str(ws.cell(r, COL_C + 2).value or "")),  # E
        ]).strip()

        proyecto = norm(str(ws.cell(r, COL_C + 3).value or ""))  # F
        cargo = norm(str(ws.cell(r, COL_C + 4).value or ""))     # G
        fi = norm(str(ws.cell(r, COL_C + 5).value or ""))        # H
        ff = norm(str(ws.cell(r, COL_C + 6).value or ""))        # I
        tiempo = norm(str(ws.cell(r, COL_C + 7).value or ""))    # J

        # Heurística: si es una fila “vacía” o parte del bloque de descripción, saltar.
        base_has_data = any([entidad, proyecto, cargo, fi, ff, tiempo, nro])

        # Si esta fila es el rótulo de descripción (por combinadas), NO es fila base
        if _is_desc_label_row(ws, r):
            r += 1
            continue

        if not base_has_data:
            r += 1
            continue

        # A veces hay fila de encabezados repetida (p.ej. "Nro | Entidad | ...")
        base_text = " ".join([nro, entidad, proyecto, cargo, fi, ff, tiempo]).upper()
        if ("NRO" in base_text and "ENTIDAD" in base_text and "PROYECTO" in base_text) or ("CARGO" in base_text and "FECHA" in base_text):
            r += 1
            continue

        # --- 2) capturar descripción asociada
        descripcion = ""
        if r + 1 <= end_row and _is_desc_label_row(ws, r + 1):
            # detalle normalmente en r+2
            if r + 2 <= end_row:
                descripcion = _get_desc_detail(ws, r + 2)
            # saltamos las 2 filas extra (label + detalle)
            next_r = r + 3
        else:
            next_r = r + 1

        # --- 3) normalización fechas y días
        ff_up = ff.upper()
        if ff_up in ("ACTUALIDAD", "ACTUAL", "A LA FECHA", "HASTA LA FECHA"):
            ff = "ACTUALIDAD"

        d_fi = _parse_date_any(fi)
        d_ff = datetime.now() if ff == "ACTUALIDAD" else _parse_date_any(ff)
        dias = _days_between(d_fi, d_ff)

        # --- 4) dedup por llave estable (incluye descripción)
        key = (
            entidad.lower(), proyecto.lower(), cargo.lower(),
            fi, ff, tiempo.lower(),
            descripcion.lower()
        )
        if key in seen:
            r = next_r
            continue
        seen.add(key)

        it = {
            "row": r,
            "nro": nro,
            "entidad": entidad,
            "proyecto": proyecto,
            "cargo": cargo,
            "fecha_inicio": fi,
            "fecha_fin": ff,
            "tiempo_en_cargo": tiempo,
            "dias_calc": dias,
            "descripcion": descripcion,
        }
        items.append(it)

        # --- 5) resumen humano (para log / escribir si quisieras)
        head = " | ".join([p for p in [entidad, proyecto, cargo] if p]).strip()
        fechas = " - ".join([p for p in [fi, ff] if p]).strip(" -")
        tail = " | ".join([p for p in [fechas, tiempo] if p]).strip()
        line = " | ".join([p for p in [head, tail] if p]).strip()
        if descripcion:
            line += f"\n  Desc: {descripcion}"
        lines.append(line)

        r = next_r

    total_dias = sum(int(x.get("dias_calc") or 0) for x in items)
    total_anios = round(total_dias / 365.0, 2) if total_dias else 0.0
    resumen = "\n\n".join([x for x in lines if x]).strip()

    out: Dict[str, Any] = {
        "items": items,
        "total_dias_calc": total_dias,
        "total_anios_calc": total_anios,
        "resumen": resumen,
    }

    if debug:
        print("[DEBUG EG] range:", start_row, "->", end_row, "| items:", len(items))
        print("  total_dias_calc:", total_dias, "| total_anios_calc:", total_anios)
        if items:
            sample = items[0]
            print("  sample:")
            print("   - entidad:", sample.get("entidad"))
            print("   - proyecto:", sample.get("proyecto"))
            print("   - cargo:", sample.get("cargo"))
            print("   - fi/ff:", sample.get("fecha_inicio"), "/", sample.get("fecha_fin"))
            print("   - tiempo:", sample.get("tiempo_en_cargo"))
            print("   - desc_len:", len(sample.get("descripcion") or ""))

    return out


# -------------------------
# API principal
# -------------------------
def parse_eoi_excel_old(xlsx_path: Path, debug: bool = False) -> Dict[str, Any]:
    wb = load_workbook(xlsx_path, data_only=True)
    ws = wb[wb.sheetnames[0]]

    dp = parse_datos_personales(ws, start_row=12, end_row=23, max_cols=12, debug=debug)
    fa = parse_formacion_academica_excel(ws, debug=debug)
    ec = parse_estudios_complementarios_excel(ws, debug=debug)

    # Compatibilidad: "cursos" (lista simple) para write_postulante viejo
    # Si ec trae bloques/items, aplanamos a líneas resumidas
    cursos_flat = []
    try:
        blocks = ec.get("blocks", []) or []
        for b in blocks:
            items = b.get("items", []) or []
            for it in items:
                centro = (it.get("centro") or "").strip()
                cap = (it.get("capacitacion") or "").strip()
                fi = (it.get("fecha_inicio") or it.get("fi") or "").strip()
                ff = (it.get("fecha_fin") or it.get("ff") or "").strip()
                horas = it.get("horas") or it.get("n_horas") or ""
                # arma una línea humana
                if centro or cap:
                    line = " - ".join([x for x in [centro, cap] if x])
                    extras = " | ".join([x for x in [fi, ff, str(horas).strip()] if x and str(x).strip()])
                    if extras:
                        line = f"{line} ({extras})"
                    cursos_flat.append(line)
    except Exception:
        # si cambia estructura, no reventar etapa 1
        cursos_flat = []

    data: Dict[str, Any] = {
        "source_file": str(xlsx_path),

        # datos personales
        "dni": dp.get("dni", ""),
        "apellido_paterno": dp.get("apellido_paterno", ""),
        "apellido_materno": dp.get("apellido_materno", ""),
        "nombres": dp.get("nombres", ""),
        "nombre_full": dp.get("nombre_full", ""),
        "email": dp.get("email", ""),
        "celular": dp.get("celular", ""),

        # formación académica
        "formacion_items": fa.get("items", []),
        "formacion_resumen": fa.get("resumen", ""),

        # ✅ estudios complementarios (nuevo)
        # Guardamos TODO (estructura completa) para que Task 30 pueda escribir b.1/b.2/b.3 dinámico
        "estudios_complementarios": ec,                 # dict completo: blocks, totals, etc.
        "ec_blocks": ec.get("blocks", []) or [],        # acceso directo
        "ec_total_horas": ec.get("total_horas", 0) or 0, # si lo calculas
        "ec_resumen": ec.get("resumen", "") or "",      # resumen general si existe

        # Compatibilidad con tu etapa vieja (si todavía se usa "cursos")
        "cursos": cursos_flat,

        # placeholders (para no romper etapas posteriores)
        "titulo": "",
        "bachiller": "",
        "egresado": "",
        "formacion_academica": {"requisito_texto": "", "items": []},

        "experiencias": [],
        "exp_general_dias": 0,
        "exp_especifica_dias": 0,
        "java_ok": False,
        "oracle_ok": False,
    }

    if debug:
        print("[EC] blocks =", len(data["ec_blocks"]), "| total_horas =", data["ec_total_horas"])
        if data["cursos"]:
            print("[EC] cursos_flat[0:3] =>", data["cursos"][:3])

    return data

def _flat_cursos_from_ec(ec: dict) -> list[str]:
    """Compatibilidad: aplana ec.blocks/items en líneas humanas, dedup y sin headers."""
    cursos = []
    seen = set()

    blocks = (ec or {}).get("blocks") or []
    for b in blocks:
        items = b.get("items") or []
        for it in items:
            centro = (it.get("centro") or "").strip()
            cap = (it.get("capacitacion") or "").strip()
            fi = (it.get("fecha_inicio") or it.get("fi") or "").strip()
            ff = (it.get("fecha_fin") or it.get("ff") or "").strip()
            horas = it.get("horas") or it.get("n_horas") or ""

            # filtro defensivo de headers/ruido
            if (it.get("nro") or "").strip().lower() in ("no.", "n°", "nº"):
                continue
            if centro.lower() == "centro de estudios" and cap.lower() in ("capacitacion", "capacitación"):
                continue

            if not (centro or cap):
                continue

            base = " - ".join([x for x in [centro, cap] if x]).strip(" -")
            extras_list = []
            if fi: extras_list.append(fi)
            if ff: extras_list.append(ff)
            # horas solo si es algo “real”
            h = str(horas).strip()
            if h and h not in ("0", "0.0"):
                extras_list.append(f"{h}h" if not h.lower().endswith("h") else h)

            line = base
            if extras_list:
                line = f"{base} ({' | '.join(extras_list)})"

            key = line.lower()
            if key in seen:
                continue
            seen.add(key)
            cursos.append(line)

    return cursos

def parse_eoi_excel(xlsx_path: Path, debug: bool = False) -> Dict[str, Any]:
    wb = load_workbook(xlsx_path, data_only=True)
    ws = wb[wb.sheetnames[0]]

    # =========================
    # 1) Parseadores base (SSOT)
    # =========================
    dp = parse_datos_personales(ws, start_row=12, end_row=23, max_cols=12, debug=debug)
    fa = parse_formacion_academica_excel(ws, debug=debug)
    ec = parse_estudios_complementarios_excel(ws, debug=debug)

    # Experiencia general (rango fijo por ahora; luego lo haremos dinámico con task_00)
    eg = parse_experiencia_general_excel(ws, start_row=101, end_row=145, debug=debug)

    exp_general_items = eg.get("items", []) or []
    exp_general_resumen = eg.get("resumen", "") or ""
    exp_general_dias = eg.get("total_dias_calc", eg.get("total_dias", 0)) or 0

    # =========================
    # 2) Construcción de DATA (sin duplicar keys)
    # =========================
    data: Dict[str, Any] = {
        "source_file": str(xlsx_path),

        # ---- Datos personales ----
        "dni": dp.get("dni", ""),
        "apellido_paterno": dp.get("apellido_paterno", ""),
        "apellido_materno": dp.get("apellido_materno", ""),
        "nombres": dp.get("nombres", ""),
        "nombre_full": dp.get("nombre_full", ""),
        "email": dp.get("email", ""),
        "celular": dp.get("celular", ""),

        # ---- Formación académica ----
        "formacion_items": fa.get("items", []) or [],
        "formacion_resumen": fa.get("resumen", "") or "",

        # ---- Estudios complementarios (SSOT) ----
        "estudios_complementarios": ec,

        # ---- Experiencia general (SSOT) ----
        "exp_general": eg,
        "exp_general_items": exp_general_items,
        "exp_general_resumen": exp_general_resumen,
        "exp_general_dias": exp_general_dias,

        # ---- Compatibilidad legacy (solo si aún lo necesitas) ----
        "cursos": _flat_cursos_from_ec(ec),

        # ---- Placeholders (solo los que realmente se usan después) ----
        "titulo": "",
        "bachiller": "",
        "egresado": "",
        "formacion_academica": {"requisito_texto": "", "items": []},

        "experiencias": [],
        "exp_especifica_dias": 0,
        "java_ok": False,
        "oracle_ok": False,
    }

    # =========================
    # 3) DEBUG estructurado
    # =========================
    if debug:
        print("\n[DEBUG PARSE_EOI] ===============================")
        print("[DP] dni=", data["dni"], "| nombre=", data["nombre_full"], "| email=", data["email"])
        print("[FA] resumen_len=", len(data["formacion_resumen"] or ""), "| items=", len(data["formacion_items"]))

        blocks_n = len((ec or {}).get("blocks") or [])
        total_h = (ec or {}).get("total_horas", 0)
        print(f"[EC] blocks={blocks_n} total_horas={total_h}")

        print(f"[EG] items={len(exp_general_items)} total_dias={exp_general_dias} resumen_len={len(exp_general_resumen)}")
        if exp_general_items:
            s0 = exp_general_items[0]
            print("  [EG sample] entidad=", s0.get("entidad"), "| proyecto=", s0.get("proyecto"))
            print("              cargo=", s0.get("cargo"), "| fi/ff=", s0.get("fecha_inicio"), "/", s0.get("fecha_fin"))
            print("              desc_len=", len(s0.get("descripcion") or ""))

        if data["cursos"]:
            print("[EC legacy cursos] first3 =>", data["cursos"][:3])
        print("[DEBUG PARSE_EOI] ===============================\n")

    return data
