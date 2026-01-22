# parsers/eoi_excel.py
# -*- coding: utf-8 -*-
"""
Parser EOI (Excel) - CONSOLIDADOR EOI

✅ Corregido “completamente”:
- Eliminadas funciones duplicadas (norm, _norm, _as_date_str, _parse_date_any, etc.)
- Eliminados parsers _old y prints sueltos.
- Formación Académica: detección dinámica de fila header + mapeo robusto de columnas.
- Estudios Complementarios: detección dinámica de bloques b.1, b.2... + corte por sección IV Experiencia.
- Experiencia General: mantiene tu parser por bloques con descripción (C..J + descripción) pero
  ahora permite rangos dinámicos via layout (si luego lo enchufas con Task_00).
- parse_eoi_excel(): construye un dict limpio (SSOT), sin keys duplicadas.

Requisitos:
  pip install openpyxl
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, Any, Optional, List, Tuple
from datetime import datetime, date

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet


# ============================================================
# Utils comunes
# ============================================================
def norm(x: Any) -> str:
    """Normaliza espacios y convierte None a ''."""
    if x is None:
        return ""
    s = str(x).replace("\u00a0", " ").strip()
    return re.sub(r"\s+", " ", s)


def cell_str(ws: Worksheet, r: int, c: int) -> str:
    return norm(ws.cell(row=r, column=c).value)


def row_text(ws: Worksheet, r: int, c1: int = 1, c2: int = 12) -> str:
    parts = []
    for c in range(c1, c2 + 1):
        v = cell_str(ws, r, c)
        if v:
            parts.append(v)
    return norm(" ".join(parts))


def as_date_str(v: Any) -> str:
    """Convierte date/datetime a dd/mm/yyyy; si es string, devuelve normalizado."""
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.date().strftime("%d/%m/%Y")
    if isinstance(v, date):
        return v.strftime("%d/%m/%Y")
    return norm(v)


def as_int(v: Any, default: int = 0) -> int:
    if v is None:
        return default
    s = norm(v)
    if not s:
        return default
    try:
        return int(float(s.replace(",", ".")))
    except Exception:
        return default


def parse_date_any(s: str) -> Optional[datetime]:
    """Acepta dd/mm/yyyy o dd-mm-yyyy."""
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


def days_between(fi: Optional[datetime], ff: Optional[datetime]) -> int:
    if not fi or not ff:
        return 0
    return max((ff - fi).days, 0)


def normalize_dni(s: str) -> str:
    s = norm(s)
    m = re.search(r"\b(\d{8})\b", s)
    return m.group(1) if m else s


def normalize_email(s: str) -> str:
    s = norm(s)
    m = re.search(r"([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})", s)
    return m.group(1) if m else s


def normalize_phone(s: str) -> str:
    s = norm(s)
    return re.sub(r"\D+", "", s)


# ============================================================
# 1) Datos Personales (filas 12-23; estructura header/value)
# ============================================================
def parse_datos_personales(ws: Worksheet, start_row: int = 12, end_row: int = 23, max_cols: int = 12, debug: bool = False) -> Dict[str, Any]:
    """
    Datos personales en estructura de 2 filas:
      - fila header
      - fila values
    (en tu formato: 12/13, 14/15, ..., 22/23)
    """
    out = {
        "dni": "",
        "apellido_paterno": "",
        "apellido_materno": "",
        "nombres": "",
        "nombre_full": "",
        "email": "",
        "celular": "",
    }

    def get_cells(r: int) -> List[str]:
        return [cell_str(ws, r, c) for c in range(1, max_cols + 1)]

    r = start_row
    while r + 1 <= end_row:
        header_row = r
        value_row = r + 1
        headers = get_cells(header_row)
        values = get_cells(value_row)

        if debug:
            print(f"[DP] header_row={header_row} -> {headers}")
            print(f"[DP] value_row ={value_row} -> {values}")

        for h, v in zip(headers, values):
            h_low = (h or "").lower()
            v = v or ""
            if not h_low:
                continue

            if "apellido paterno" in h_low:
                out["apellido_paterno"] = v
                continue
            if "apellido materno" in h_low:
                out["apellido_materno"] = v
                continue

            # nombres
            if re.search(r"\bnombres\b", h_low):
                out["nombres"] = v
                continue

            # DNI
            if ("documento" in h_low and "identidad" in h_low) or "dni" in h_low:
                out["dni"] = normalize_dni(v or row_text(ws, value_row, 1, max_cols))
                continue

            # celular
            if "celular" in h_low:
                vv = normalize_phone(v or row_text(ws, value_row, 1, max_cols))
                if vv:
                    out["celular"] = vv
                continue

            # email
            if "email" in h_low or "correo" in h_low:
                out["email"] = normalize_email(v or row_text(ws, value_row, 1, max_cols))
                continue

        r += 2

    ap = " ".join([out["apellido_paterno"], out["apellido_materno"]]).strip()
    nm = out["nombres"].strip()
    out["nombre_full"] = norm(" ".join([ap, nm]))

    # normalizaciones finales
    out["dni"] = normalize_dni(out["dni"])
    out["email"] = normalize_email(out["email"])
    out["celular"] = normalize_phone(out["celular"])

    if debug:
        print("[DP] RESULT =>", out)

    return out


# ============================================================
# 2) Formación Académica (detección dinámica)
# ============================================================
def _find_header_row_formacion(ws: Worksheet, scan_from: int = 35, scan_to: int = 80) -> Optional[int]:
    """
    Encuentra la fila donde está el header de la tabla de Formación Académica.
    Busca frases típicas: FECHA DE EXTENSIÓN, CENTRO DE ESTUDIOS, CIUDAD/PAÍS, ESPECIALIDAD.
    """
    for r in range(scan_from, min(scan_to, ws.max_row or scan_to) + 1):
        t = row_text(ws, r, 1, 15).upper()
        if ("FECHA" in t and ("EXTENS" in t or "EXTENSIÓN" in t) and ("TITULO" in t or "TÍTULO" in t)) \
           or ("CENTRO DE ESTUDIOS" in t) \
           or ("CIUDAD" in t and ("PAIS" in t or "PAÍS" in t)) \
           or ("ESPECIALIDAD" in t and ("TITULO" in t or "TÍTULO" in t)):
            return r
    return None


def _map_formacion_columns(ws: Worksheet, header_row: int, max_col: int = 20) -> Dict[str, int]:
    """
    Mapea columnas por contenido del header:
      - titulo_item
      - especialidad
      - fecha
      - centro
      - ciudad
    """
    mapping: Dict[str, int] = {}

    for c in range(1, max_col + 1):
        h = cell_str(ws, header_row, c).upper()
        if not h:
            continue

        # OJO: el "título" de fila suele ser una lista (COLEGIATURA/MAESTRIA/etc.)
        # y muchas plantillas lo ponen como "Título*" o "Título"
        if ("TITULO" in h or "TÍTULO" in h) and "FECHA" not in h:
            mapping.setdefault("titulo_item", c)

        if "ESPECIALIDAD" in h:
            mapping.setdefault("especialidad", c)

        if "FECHA" in h and ("EXTENS" in h or "EXTENSIÓN" in h or "EXTENSION" in h):
            mapping.setdefault("fecha", c)

        if "CENTRO" in h and "ESTUDIO" in h:
            mapping.setdefault("centro", c)

        if "CIUDAD" in h and ("PAIS" in h or "PAÍS" in h):
            mapping.setdefault("ciudad", c)

    return mapping


def parse_formacion_academica_excel(ws: Worksheet, debug: bool = False) -> Dict[str, Any]:
    """
    Lee la tabla de Formación Académica en tu EDI:
    - Detecta header de manera dinámica.
    - Mapea columnas por texto del header (para soportar corrimientos).
    - Lee filas bajo el header con "títulos" típicos: COLEGIATURA, MAESTRIA, TITULO, BACHILLER, EGRESADO, UNIVERSITARIO, etc.
    """
    header_row = _find_header_row_formacion(ws)
    if header_row is None:
        if debug:
            print("[FA] No se encontró header de tabla.")
        return {"items": [], "resumen": ""}

    colmap = _map_formacion_columns(ws, header_row)
    # fallback razonable si alguna columna no se detecta
    # (si tu plantilla es fija, esto te salva cuando no detecta el texto por merges)
    col_tit = colmap.get("titulo_item", 2)       # comúnmente B
    col_esp = colmap.get("especialidad", 4)      # comúnmente D
    col_fec = colmap.get("fecha", 6)             # comúnmente F
    col_cen = colmap.get("centro", 8)            # comúnmente H
    col_ciu = colmap.get("ciudad", 10)           # comúnmente J

    start_row = header_row + 2  # en tu formato: header 49, data 51
    end_row = min(start_row + 25, ws.max_row or (start_row + 25))

    expected_keys = ("COLEGIATURA", "MAESTR", "EGRESAD", "TITUL", "BACHILL", "UNIVERSIT")

    items: List[Dict[str, Any]] = []
    for r in range(start_row, end_row + 1):
        titulo = cell_str(ws, r, col_tit)
        # fallback: a veces está en col A por merges
        if not titulo:
            titulo = cell_str(ws, r, 1)

        if not titulo:
            # no hacemos break: hay casos con filas vacías intermedias
            continue

        titulo_up = titulo.upper()
        if not any(k in titulo_up for k in expected_keys):
            # probablemente ya salimos de la tabla
            # pero para no cortar mal, exigimos que además la fila esté vacía en el resto
            maybe_other = any(cell_str(ws, r, c) for c in (col_esp, col_fec, col_cen, col_ciu))
            if not maybe_other:
                continue
            # si hay data sin "titulo", igual lo tomamos
            # (pero normalmente no pasa)
        especialidad = cell_str(ws, r, col_esp)
        fecha = as_date_str(ws.cell(row=r, column=col_fec).value)
        centro = cell_str(ws, r, col_cen)
        ciudad = cell_str(ws, r, col_ciu)

        has_data = any([especialidad, fecha, centro, ciudad])

        items.append({
            "row": r,
            "titulo_item": titulo,
            "especialidad": especialidad,
            "fecha": fecha,
            "centro": centro,
            "ciudad": ciudad,
            "has_data": has_data
        })

        if debug:
            print(f"[FA] r={r} titulo='{titulo}' esp='{especialidad}' fecha='{fecha}' centro='{centro}' ciudad='{ciudad}' has={has_data}")

    picked = [x for x in items if x.get("has_data")]
    parts = []
    for it in picked:
        p = f"{it['titulo_item']}: {it['especialidad']}".strip()
        extras = [x for x in [it.get("fecha", ""), it.get("centro", ""), it.get("ciudad", "")] if norm(x)]
        if extras:
            p += " (" + " | ".join(extras) + ")"
        parts.append(p)

    resumen = " ; ".join(parts) if parts else ""
    return {"items": items, "resumen": resumen, "_meta": {"header_row": header_row, "colmap": colmap}}


# ============================================================
# 3) Estudios Complementarios (bloques b.1, b.2, ...)
# ============================================================
def _is_section_iv_experiencia(ws: Worksheet, r: int) -> bool:
    t = row_text(ws, r, 1, 12).upper()
    return ("IV" in t and "EXPERI" in t) or t.startswith("IV") and "EXPERI" in t


def _is_course_header_like(nro: str, centro: str, cap: str, fi: str, ff: str) -> bool:
    s1 = norm(nro).lower()
    s2 = norm(centro).lower()
    s3 = norm(cap).lower()
    s4 = norm(fi).lower()
    s5 = norm(ff).lower()
    if s1 in ("no.", "n°", "nº", "nro", "nro."):
        return True
    if "centro de estudios" in s2 and ("capacit" in s3):
        return True
    if s4.startswith("fecha") or s5.startswith("fecha"):
        return True
    return False


def parse_estudios_complementarios_excel(ws: Worksheet, debug: bool = False) -> Dict[str, Any]:
    """
    Estudios Complementarios:
      - Detecta bloques: b.1, b.2, b.3... (en cualquier columna A..J).
      - Data inicia 4 filas debajo del título del bloque.
      - Columnas (plantilla real):
          C: N°
          D: Centro
          F: Capacitación
          H: Fecha Inicio
          I: Fecha Fin
          J: Horas
      - Corta al entrar a sección IV. EXPERIENCIA
    """
    pat_block = re.compile(r"^\s*b\s*\.?\s*(\d+)\s*\)?\s*", re.IGNORECASE)
    max_row = ws.max_row or 0

    blocks: List[Dict[str, Any]] = []

    # 1) localizar cabeceras b.x
    for r in range(1, max_row + 1):
        if _is_section_iv_experiencia(ws, r):
            break

        for c in range(1, 11):
            v = cell_str(ws, r, c)
            if not v:
                continue
            m = pat_block.match(v)
            if m:
                idx = int(m.group(1))
                blocks.append({
                    "id": f"b.{idx}",
                    "row": r,
                    "title": v,
                    "items": [],
                    "total_horas": 0,
                    "resumen": ""
                })
                if debug:
                    print(f"[EC] detectado b.{idx} fila {r} col {c}: {v}")
                break

    if not blocks:
        return {"blocks": [], "total_horas": 0, "resumen": ""}

    # 2) por cada bloque, leer items
    total_horas_global = 0
    for i, b in enumerate(blocks):
        header_row = b["row"]
        next_header_row = blocks[i + 1]["row"] if i + 1 < len(blocks) else None

        data_start = header_row + 4
        data_end = (next_header_row - 1) if next_header_row else max_row

        if debug:
            print(f"[EC] {b['id']} data_start={data_start} data_end={data_end}")

        seen = set()

        for r in range(data_start, data_end + 1):
            if _is_section_iv_experiencia(ws, r):
                if debug:
                    print(f"[EC] corte por sección IV dentro de {b['id']} en fila {r}")
                break

            nro = cell_str(ws, r, 3)
            centro = cell_str(ws, r, 4)
            cap = cell_str(ws, r, 6)
            fi = as_date_str(ws.cell(row=r, column=8).value)
            ff = as_date_str(ws.cell(row=r, column=9).value)
            horas = as_int(ws.cell(row=r, column=10).value, default=0)

            # ignora headers incrustados
            if _is_course_header_like(nro, centro, cap, fi, ff):
                continue

            # fila vacía real
            if not (centro or cap or nro or fi or ff or horas):
                continue

            key = (nro, centro, cap, fi, ff, horas)
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

        # resumen por bloque
        lines = []
        for it in b["items"]:
            if not (it["centro"] or it["capacitacion"]):
                continue
            left = " - ".join([x for x in [it["centro"], it["capacitacion"]] if x])
            extras = " | ".join([x for x in [it["fecha_inicio"], it["fecha_fin"]] if x])
            if it["horas"]:
                extras = (extras + " | " if extras else "") + f"{it['horas']}h"
            lines.append(f"{left} ({extras})" if extras else left)
        b["resumen"] = "\n".join(lines).strip()

    # resumen global
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


def flat_cursos_from_ec(ec: Dict[str, Any]) -> List[str]:
    """Compatibilidad legacy: aplana ec.blocks/items a líneas humanas, dedup."""
    cursos: List[str] = []
    seen = set()

    blocks = (ec or {}).get("blocks") or []
    for b in blocks:
        for it in (b.get("items") or []):
            centro = norm(it.get("centro", ""))
            cap = norm(it.get("capacitacion", ""))
            fi = norm(it.get("fecha_inicio", ""))
            ff = norm(it.get("fecha_fin", ""))
            horas = it.get("horas", 0) or 0

            if not (centro or cap):
                continue

            base = " - ".join([x for x in [centro, cap] if x]).strip(" -")
            extras = [x for x in [fi, ff] if x]
            if horas:
                extras.append(f"{horas}h")
            line = f"{base} ({' | '.join(extras)})" if extras else base

            key = line.lower()
            if key in seen:
                continue
            seen.add(key)
            cursos.append(line)

    return cursos


# ============================================================
# 4) Experiencia General (bloques con descripción)
# ============================================================
def _is_desc_label_row(ws: Worksheet, row: int) -> bool:
    t = row_text(ws, row, 1, 12).upper()
    return ("DESCRIP" in t) and ("TRABAJO" in t) and ("REALIZ" in t)


def _get_desc_detail(ws: Worksheet, row: int) -> str:
    """Toma el texto más largo de A..L en esa fila (típico de celda combinada)."""
    best = ""
    for c in range(1, 13):
        s = norm(ws.cell(row=row, column=c).value)
        if len(s) > len(best):
            best = s
    return best


def parse_experiencia_general_excel(
    ws: Worksheet,
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
    """
    COL_C = 3
    items: List[Dict[str, Any]] = []
    lines: List[str] = []
    seen = set()

    r = start_row
    while r <= end_row:
        if _is_desc_label_row(ws, r):
            r += 1
            continue

        nro = norm(ws.cell(row=r, column=COL_C + 0).value)
        entidad = " ".join([
            norm(ws.cell(row=r, column=COL_C + 1).value),
            norm(ws.cell(row=r, column=COL_C + 2).value),
        ]).strip()

        proyecto = norm(ws.cell(row=r, column=COL_C + 3).value)
        cargo = norm(ws.cell(row=r, column=COL_C + 4).value)
        fi = norm(ws.cell(row=r, column=COL_C + 5).value)
        ff = norm(ws.cell(row=r, column=COL_C + 6).value)
        tiempo = norm(ws.cell(row=r, column=COL_C + 7).value)

        base_has_data = any([nro, entidad, proyecto, cargo, fi, ff, tiempo])
        if not base_has_data:
            r += 1
            continue

        # cabeceras repetidas
        base_text = " ".join([nro, entidad, proyecto, cargo, fi, ff, tiempo]).upper()
        if ("NRO" in base_text and "ENTIDAD" in base_text and "PROYECTO" in base_text) or ("CARGO" in base_text and "FECHA" in base_text):
            r += 1
            continue

        descripcion = ""
        if r + 1 <= end_row and _is_desc_label_row(ws, r + 1):
            if r + 2 <= end_row:
                descripcion = _get_desc_detail(ws, r + 2)
            next_r = r + 3
        else:
            next_r = r + 1

        ff_up = ff.upper()
        if ff_up in ("ACTUALIDAD", "ACTUAL", "A LA FECHA", "HASTA LA FECHA"):
            ff = "ACTUALIDAD"

        d_fi = parse_date_any(fi)
        d_ff = datetime.now() if ff == "ACTUALIDAD" else parse_date_any(ff)
        dias = days_between(d_fi, d_ff)

        key = (entidad.lower(), proyecto.lower(), cargo.lower(), fi, ff, tiempo.lower(), descripcion.lower())
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

    out = {
        "items": items,
        "total_dias_calc": total_dias,
        "total_anios_calc": total_anios,
        "resumen": resumen,
        "_meta": {"start_row": start_row, "end_row": end_row}
    }

    if debug:
        print(f"[EG] items={len(items)} total_dias={total_dias} total_anios={total_anios}")

    return out


# ============================================================
# API principal
# ============================================================
def parse_eoi_excel(
    xlsx_path: Path,
    debug: bool = False,
    layout: Optional[Dict[str, Any]] = None
) -> Dict[str, Any]:
    """
    Parser principal de Excel.

    layout (opcional) te permite rangos dinámicos (para cuando conectes Task_00):
      layout = {
        "datos_personales": {"start_row": 12, "end_row": 23},
        "experiencia_general": {"start_row": 101, "end_row": 145}
      }
    """
    wb = load_workbook(xlsx_path, data_only=True)
    ws = wb[wb.sheetnames[0]]

    layout = layout or {}
    dp_cfg = layout.get("datos_personales", {}) or {}
    eg_cfg = layout.get("experiencia_general", {}) or {}

    dp = parse_datos_personales(
        ws,
        start_row=int(dp_cfg.get("start_row", 12)),
        end_row=int(dp_cfg.get("end_row", 23)),
        max_cols=int(dp_cfg.get("max_cols", 12)),
        debug=debug
    )

    fa = parse_formacion_academica_excel(ws, debug=debug)
    ec = parse_estudios_complementarios_excel(ws, debug=debug)

    eg = parse_experiencia_general_excel(
        ws,
        start_row=int(eg_cfg.get("start_row", 101)),
        end_row=int(eg_cfg.get("end_row", 145)),
        debug=debug
    )

    # Campos SSOT
    data: Dict[str, Any] = {
        "source_file": str(xlsx_path),

        # Datos personales (normalizados)
        "dni": normalize_dni(dp.get("dni", "")),
        "apellido_paterno": norm(dp.get("apellido_paterno", "")),
        "apellido_materno": norm(dp.get("apellido_materno", "")),
        "nombres": norm(dp.get("nombres", "")),
        "nombre_full": norm(dp.get("nombre_full", "")),
        "email": normalize_email(dp.get("email", "")),
        "celular": normalize_phone(dp.get("celular", "")),

        # Formación académica
        "formacion_items": fa.get("items", []) or [],
        "formacion_resumen": fa.get("resumen", "") or "",
        "formacion_meta": fa.get("_meta", {}) or {},

        # Estudios complementarios (bloques)
        "estudios_complementarios": ec,

        # Experiencia general (bloques + descripción)
        "exp_general": eg,
        "exp_general_items": eg.get("items", []) or [],
        "exp_general_resumen": eg.get("resumen", "") or "",
        "exp_general_dias": int(eg.get("total_dias_calc", 0) or 0),

        # Compatibilidad legacy (si algo todavía lo consume)
        "cursos": flat_cursos_from_ec(ec),

        # Placeholders (solo si otras partes esperan estas keys)
        "experiencias": [],           # si luego agregas experiencia específica en otro parser
        "exp_especifica_dias": 0,
        "java_ok": False,
        "oracle_ok": False,
    }

    if debug:
        print("\n[DEBUG parse_eoi_excel]")
        print(" file:", xlsx_path)
        print(" dni:", data["dni"], "| nombre:", data["nombre_full"])
        print(" FA items:", len(data["formacion_items"]), "| EC blocks:", len((ec or {}).get("blocks") or []))
        print(" EG items:", len(data["exp_general_items"]), "| EG dias:", data["exp_general_dias"])
        print("[/DEBUG]\n")

    return data
