# tasks/task_41_eval_procesos_openai.py
# -*- coding: utf-8 -*-

import argparse
import json
from pathlib import Path
from datetime import datetime, date
from typing import Dict, Any, List, Optional, Tuple
from openpyxl.cell.cell import MergedCell

#from core.llm_client import evaluar_formacion, evaluar_estudios_complementarios
from core.openai_client import (
    evaluar_formacion,
    evaluar_estudios_complementarios,
    evaluar_experiencia_general,
    evaluar_experiencia_especifica,
)


OUT_FOLDER_NAME = "011. INSTALACIÓN DE COMITÉ"
PROCESADOS_SUBFOLDER = "procesados"
IN_CONSOLIDADO = "parsed_postulantes.jsonl"
IN_CRITERIA = "criteria_evaluacion.json"
OUT_EVAL = "evaluacion_postulantes.jsonl"
OUT_RESUMEN = "evaluacion_resumen.json"

def ts():
    return datetime.now().isoformat(timespec="seconds")

def norm(s: str) -> str:
    return " ".join((s or "").strip().split())

def read_jsonl(path: Path) -> List[Dict[str, Any]]:
    rows = []
    if not path.exists():
        return rows
    with path.open("r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                rows.append(json.loads(line))
    return rows

def write_jsonl(path: Path, rows: List[Dict[str, Any]]):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        for r in rows:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")

def write_json(path: Path, obj: Dict[str, Any]):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")

def ensure_postulante_dict(p: Any) -> Dict[str, Any]:
    """
    Asegura que 'p' sea un dict válido.
    - Si ya es dict => OK
    - Si es str con JSON => json.loads
    - Si no se puede => lanza ValueError
    """
    if isinstance(p, dict):
        return p

    if isinstance(p, str):
        s = p.strip()
        if not s:
            raise ValueError("Postulante vacío (string vacío)")
        try:
            obj = json.loads(s)
        except Exception as e:
            raise ValueError(f"Postulante no es JSON válido (string): {e}")
        if not isinstance(obj, dict):
            raise ValueError(f"Postulante JSON no es objeto dict, es {type(obj).__name__}")
        return obj

    raise ValueError(f"Postulante tiene tipo inválido: {type(p).__name__}")

def is_evaluable_postulante(p: Dict[str, Any]) -> Tuple[bool, str]:
    # si no tiene DNI, ni lo intentes
    dni = (p.get("dni") or "").strip()
    if not dni:
        return False, "sin DNI"

    # experiencia: si no hay items y todo está en 0, no hay con qué evaluar
    eg = p.get("exp_general") or {}
    ee = p.get("exp_especifica") or {}

    eg_items = eg.get("items") or []
    ee_items = ee.get("items") or []

    eg_dias = p.get("exp_general_dias", 0) or 0
    ee_dias = p.get("exp_especifica_dias", 0) or 0

    # si viene de PDF y no se extrajo nada, casi seguro falló parse/OCR
    meta = p.get("_meta") or {}
    tipo = (meta.get("tipo") or "").upper()

    if (not eg_items and not ee_items) and (eg_dias == 0 and ee_dias == 0):
        if tipo == "PDF":
            return False, "PDF sin experiencia extraída (probable parse fallido / requiere OCR)"
        return False, "sin experiencia extraída"

    return True, "OK"

def get_formacion_text(p: Dict[str, Any]) -> str:
    # 1) resumen directo
    v = p.get("formacion_resumen")
    if isinstance(v, str) and v.strip():
        return v.strip()

    # 2) dict formacion_obligatoria
    fo = p.get("formacion_obligatoria")
    if isinstance(fo, dict):
        r = fo.get("resumen")
        if isinstance(r, str) and r.strip():
            return r.strip()

    # 3) items
    items = p.get("formacion_items") or []
    if isinstance(items, list) and items:
        lines = []
        for it in items:
            if not isinstance(it, dict):
                continue
            grado = norm(str(it.get("grado", "") or ""))
            carrera = norm(str(it.get("carrera", "") or ""))
            entidad = norm(str(it.get("entidad", "") or ""))
            line = " | ".join([x for x in [grado, carrera, entidad] if x])
            if line:
                lines.append(line)
        if lines:
            return "\n".join(lines)

    return ""

def _parse_fecha_any(x: str) -> date | None:
    """
    Intenta parsear fechas comunes:
    YYYY-MM-DD
    DD/MM/YYYY
    """
    x = (x or "").strip()
    if not x:
        return None

    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(x, fmt).date()
        except ValueError:
            continue
    return None

def get_formacion_fecha_minima(p: Dict[str, Any]) -> str:
    """
    Retorna la fecha mínima (más antigua) de Formación Académica.
    Devuelve string ISO (YYYY-MM-DD) o "" si no hay datos válidos.
    """
    fo = p.get("formacion_obligatoria") or {}
    items = fo.get("items") or []
    fechas: list[date] = []

    if isinstance(items, list):
        for it in items:
            if not isinstance(it, dict):
                continue

            f_raw = it.get("fecha")
            f = _parse_fecha_any(f_raw)
            if f:
                fechas.append(f)

    if not fechas:
        return ""

    return min(fechas).isoformat()

def get_ec_blocks(p: Dict[str, Any]) -> List[Dict[str, Any]]:
    ec = p.get("estudios_complementarios") or {}
    blocks = ec.get("blocks")
    out = []
    if isinstance(blocks, list):
        for b in blocks:
            if isinstance(b, dict):
                out.append({
                    "id": norm(str(b.get("id", ""))),
                    "title": norm(str(b.get("title", ""))),
                    "resumen": (b.get("resumen") or "").strip() if isinstance(b.get("resumen"), str) else ""
                })
    return out

def get_ec_fallback_text(p: Dict[str, Any]) -> str:
    for k in ("estudios_complementarios_resumen", "ec_resumen", "cursos_resumen"):
        v = p.get(k)
        if isinstance(v, str) and v.strip():
            return v.strip()
    return ""

def get_experiencia_general_text(p: Dict[str, Any]) -> str:
    """
    Arma evidencia EG SOLO con empresa | cargo | fechas,
    y total años/días ya calculados (sin solapamiento) si existen.
    """

    # 1) Obtiene experiencia general desde _fill_payload (si existe)
    _fill_payload=p.get("_fill_payload", {} or {})                  
    
    exp_general_total_text=_fill_payload.get("exp_general_total_text")
    exp_general_resumen_text=_fill_payload.get("exp_general_resumen_text")  
    exp_general_detalle_text=_fill_payload.get("exp_general_detalle_text")

    experiencia_general=  (exp_general_total_text if exp_general_total_text else "") + "\n" + \
                          "Con el siguiente resumen :\n" + (exp_general_resumen_text if exp_general_resumen_text else "")  
    
    #experiencia_general_detalle= (exp_general_total_text if exp_general_total_text else "") + "\n" + "Con el siguiente detalle :\n" + (exp_general_detalle_text if exp_general_detalle_text else "")  
    experiencia_general_detalle=  None

    head = []
    if experiencia_general is not None:
        head.append(f"Experiencia General, calculada sin solapamientos: {experiencia_general}")
    if experiencia_general_detalle is not None:
        head.append(f"Experiencia General con descripción, calculada sin solapamientos:: {experiencia_general_detalle}")

    out = []
    if head:
        out.append("\n".join(head))
    else:
        out.append("experiencias: (sin registros)")

    return "\n\n".join(out).strip()

def get_experiencia_especifica_text(p: Dict[str, Any]) -> str:
    """
    Arma evidencia EE SOLO con empresa | cargo | fechas,
    y total años/días ya calculados (sin solapamiento) si existen.
    """
    # 1) Obtiene experiencia especifica desde _fill_payload (si existe)
    _fill_payload=p.get("_fill_payload", {} or {})                  
    
    exp_especifica_total_text=_fill_payload.get("exp_especifica_total_text")
    exp_especifica_resumen_text=_fill_payload.get("exp_especifica_resumen_text")  
    exp_especifica_detalle_text=_fill_payload.get("exp_especifica_detalle_text")

    experiencia_especifica=  (exp_especifica_total_text if exp_especifica_total_text else "") + "\n" + \
                          "Con el siguiente resumen :\n" + (exp_especifica_resumen_text if exp_especifica_resumen_text else "")  
    
    #experiencia_especifica_detalle= (exp_especifica_total_text if exp_especifica_total_text else "") + "\n" + "Con el siguiente detalle :\n" + (exp_especifica_detalle_text if exp_especifica_detalle_text else "")  
    experiencia_especifica_detalle=  None

    # 2) lista de items (empresa/cargo/fechas). Ajusta keys si tus nombres difieren.
#    exp_especifica= p.get("exp_especifica") or {}

    #    items = (
    #        exp_especifica.get("items")  or
    #        []
    #    )

    #    lines = []
    #    for it in items:
    #        emp = (it.get("empresa") or it.get("entidad") or it.get("institucion") or "").strip()
    #        cargo = (it.get("cargo") or it.get("puesto") or "").strip()
    #        fi = (it.get("fecha_inicio") or it.get("fi") or "").strip()
    #        ff = (it.get("fecha_fin") or it.get("ff") or "").strip()
    #        desc = (it.get("descripcion") or it.get("desc") or "").strip()

    #        if not (emp or cargo or fi or ff):
    #            continue

        # SOLO empresa | cargo | fechas
    #        seg = " | ".join([x for x in [emp, cargo, f"{fi} - {ff}".strip()] if x])
    #        if seg:
    #            lines.append(seg)

    head = []
    if experiencia_especifica is not None:
        head.append(f"Experiencia Específica, calculada sin solapamientos: {experiencia_especifica}")
    if experiencia_especifica_detalle is not None:
        head.append(f"Experiencia Específica con descripción, calculada sin solapamientos:: {experiencia_especifica_detalle}")

    out = []
    if head:
        out.append("\n".join(head))
    #if lines: ##Se usara solo si hay items
    #    out.append("experiencias:\n- " + "\n- ".join(lines))
    else:
        out.append("experiencias: (sin registros)")
    
    return "\n\n".join(out).strip()

def eval_one_postulante_old(p: Dict[str, Any], criteria: Dict[str, Any], debug: bool = False) -> Dict[str, Any]:
    nombre = p.get("nombre_full", "(sin nombre)")
    dni = p.get("dni", "")

    # --- FA ---
    criterio_fa = criteria["criterios"]["FA"]["criterio_item"]["text"]
    criterio_fa_row = criteria["criterios"]["FA"]["criterio_item"].get("row")

    formacion = get_formacion_text(p)

    fa = evaluar_formacion(
        criterio_text=criterio_fa,
        formacion_postulante=formacion,
        debug=debug,
    )

    # --- EC ---
    criteria_ec_blocks = criteria["criterios"]["EC"]["blocks"]
    ec_blocks_post = get_ec_blocks(p)
    ec_fallback = get_ec_fallback_text(p)

    ec_results = []
    ec_puntaje_total = 0
    ec_eliminatorio_no_cumple = False

    for i, cb in enumerate(criteria_ec_blocks):
        criterio_ec = cb["criterio_item"]["text"]
        criterio_ec_row = cb["criterio_item"].get("row")
        modo = cb.get("modo_evaluacion")
        valor = cb.get("valor")

        if i < len(ec_blocks_post) and ec_blocks_post[i].get("resumen"):
            evidencia = ec_blocks_post[i]["resumen"]
            ev_source = f"blocks[{i}] id={ec_blocks_post[i].get('id')}"
        else:
            evidencia = ec_fallback
            ev_source = "fallback_text"

        r_ec = evaluar_estudios_complementarios(
            criterio_text=criterio_ec,
            evidencia_postulante=evidencia,
            debug=debug
        )

        puntaje = 0
        if str(modo).lower() == "puntaje":
            try:
                vnum = int(valor) if str(valor).isdigit() else 0
            except Exception:
                vnum = 0
            puntaje = vnum if r_ec.get("estado") == "CUMPLE" else 0
            ec_puntaje_total += puntaje
        else:
            # eliminatorio cumpla/no cumpla
            # (si tu JSON usa otra convención, aquí se ajusta)
            if r_ec.get("estado") == "NO_CUMPLE":
                ec_eliminatorio_no_cumple = True

        ec_results.append({
            "id": cb.get("id", f"EC.{i+1}"),
            "estado": r_ec.get("estado"),
            "evidencia": r_ec.get("evidencia"),
            "justificacion": r_ec.get("justificacion"),
            "confianza": r_ec.get("confianza"),
            "modo_evaluacion": modo,
            "valor": valor,
            "puntaje": puntaje,
            "_meta": {
                "criterio_row": criterio_ec_row,
                "evidencia_source": ev_source,
                "modelo": r_ec["_llm_meta"]["model"],
                "timestamp": r_ec["_llm_meta"]["timestamp"],
            }
        })


    # --- ExpGeneral ---
    criteria_eg_lines = criteria["criterios"]["EG"]["lines"]
    eg_lines_post = get_ec_blocks(p)
    eg_fallback = get_ec_fallback_text(p)

    ec_results = []
    ec_puntaje_total = 0
    ec_eliminatorio_no_cumple = False

    for i, cb in enumerate(criteria_ec_blocks):
        criterio_ec = cb["criterio_item"]["text"]
        criterio_ec_row = cb["criterio_item"].get("row")
        modo = cb.get("modo_evaluacion")
        valor = cb.get("valor")

        if i < len(ec_blocks_post) and ec_blocks_post[i].get("resumen"):
            evidencia = ec_blocks_post[i]["resumen"]
            ev_source = f"blocks[{i}] id={ec_blocks_post[i].get('id')}"
        else:
            evidencia = ec_fallback
            ev_source = "fallback_text"

        r_ec = evaluar_estudios_complementarios(
            criterio_text=criterio_ec,
            evidencia_postulante=evidencia,
            debug=debug
        )

        puntaje = 0
        if str(modo).lower() == "puntaje":
            try:
                vnum = int(valor) if str(valor).isdigit() else 0
            except Exception:
                vnum = 0
            puntaje = vnum if r_ec.get("estado") == "CUMPLE" else 0
            ec_puntaje_total += puntaje
        else:
            # eliminatorio cumpla/no cumpla
            # (si tu JSON usa otra convención, aquí se ajusta)
            if r_ec.get("estado") == "NO_CUMPLE":
                ec_eliminatorio_no_cumple = True

        ec_results.append({
            "id": cb.get("id", f"EC.{i+1}"),
            "estado": r_ec.get("estado"),
            "evidencia": r_ec.get("evidencia"),
            "justificacion": r_ec.get("justificacion"),
            "confianza": r_ec.get("confianza"),
            "modo_evaluacion": modo,
            "valor": valor,
            "puntaje": puntaje,
            "_meta": {
                "criterio_row": criterio_ec_row,
                "evidencia_source": ev_source,
                "modelo": r_ec["_llm_meta"]["model"],
                "timestamp": r_ec["_llm_meta"]["timestamp"],
            }
        })


    return {
        "dni": dni,
        "nombre_full": nombre,
        "FA": {
            "estado": fa.get("estado"),
            "evidencia": fa.get("evidencia"),
            "justificacion": fa.get("justificacion"),
            "confianza": fa.get("confianza"),
            "eliminatorio": True,
        },
        "EC": {
            "blocks": ec_results,
            "puntaje_total": ec_puntaje_total,
            "eliminatorio_no_cumple": ec_eliminatorio_no_cumple
        },
        "_meta": {
            "evaluated_at": ts(),
            "fa_criterio_row": criterio_fa_row,
            "modelo": fa["_llm_meta"]["model"],
            "timestamp": fa["_llm_meta"]["timestamp"],
        }
    }

def eval_one_postulante(p: Dict[str, Any], criteria: Dict[str, Any], debug: bool = False) -> Dict[str, Any]:
    nombre = p.get("nombre_full", "(sin nombre)")
    dni = p.get("dni", "")

    # --- FA ---
    criterio_fa = criteria["criterios"]["FA"]["criterio_item"]["text"]
    criterio_fa_row = criteria["criterios"]["FA"]["criterio_item"].get("row")

    formacion = get_formacion_text(p)
    fecha_formacion_minima = get_formacion_fecha_minima(p)

    fa = evaluar_formacion(
        criterio_text=criterio_fa,
        formacion_postulante=formacion,
        debug=debug,
    )

    # --- EC ---
    criteria_ec_blocks = criteria["criterios"]["EC"]["blocks"]
    ec_blocks_post = get_ec_blocks(p)
    ec_fallback = get_ec_fallback_text(p)

    ec_results = []
    ec_puntaje_total = 0
    ec_eliminatorio_no_cumple = False

    for i, cb in enumerate(criteria_ec_blocks):
        criterio_ec = cb["criterio_item"]["text"]
        criterio_ec_row = cb["criterio_item"].get("row")
        modo = cb.get("modo_evaluacion")
        valor = cb.get("valor")

        if i < len(ec_blocks_post) and ec_blocks_post[i].get("resumen"):
            evidencia = ec_blocks_post[i]["resumen"]
            ev_source = f"blocks[{i}] id={ec_blocks_post[i].get('id')}"
        else:
            evidencia = ec_fallback
            ev_source = "fallback_text"

        r_ec = evaluar_estudios_complementarios(
            criterio_text=criterio_ec,
            evidencia_postulante=evidencia,
            debug=debug
        )

        puntaje = 0
        if str(modo).lower() == "puntaje":
            try:
                vnum = int(valor) if str(valor).isdigit() else 0
            except Exception:
                vnum = 0
            puntaje = vnum if r_ec.get("estado") == "CUMPLE" else 0
            ec_puntaje_total += puntaje
        else:
            if r_ec.get("estado") == "NO_CUMPLE":
                ec_eliminatorio_no_cumple = True

        ec_results.append({
            "id": cb.get("id", f"EC.{i+1}"),
            "estado": r_ec.get("estado"),
            "evidencia": r_ec.get("evidencia"),
            "justificacion": r_ec.get("justificacion"),
            "confianza": r_ec.get("confianza"),
            "modo_evaluacion": modo,
            "valor": valor,
            "puntaje": puntaje,
            "_meta": {
                "criterio_row": criterio_ec_row,
                "evidencia_source": ev_source,
                "modelo": r_ec["_llm_meta"]["model"],
                "timestamp": r_ec["_llm_meta"]["timestamp"],
            }
        })

    # =====================================================================
    # NUEVO: --- EG (Experiencia General) ---
    # =====================================================================
    criteria_eg_lines = criteria["criterios"]["EG"]["lines"]

    # evidencia EG desde parsed_postulante.jsonl (ya calculada sin solapamiento)
    # Ideal: empresa | cargo | fechas + total años/días
    eg_evidencia = get_experiencia_general_text(p)  # <-- IMPORTANTE
    
    eg_results = []
    eg_puntaje_total = 0
    eg_eliminatorio_no_cumple = False

    for i, ln in enumerate(criteria_eg_lines):
        criterio_eg = ln["criterio_item"]["text"]
        criterio_eg_row = ln["criterio_item"].get("row")
        modo = ln.get("modo_evaluacion")
        valor = ln.get("valor")

        r_eg = evaluar_experiencia_general(
            criterio_text=criterio_eg,
            evidencia_postulante=eg_evidencia,
            fecha_formacion_minima=fecha_formacion_minima,
            debug=debug
        )

        puntaje = 0
        if str(modo).lower() == "puntaje":
            # muchos criterios traen "XX" -> no es número, queda 0
            try:
                vnum = int(valor) if str(valor).isdigit() else 0
            except Exception:
                vnum = 0
            puntaje = vnum if r_eg.get("estado") == "CUMPLE" else 0
            eg_puntaje_total += puntaje
        else:
            # eliminatorio Cumple/NoCumple (primera línea usualmente)
            if r_eg.get("estado") == "NO_CUMPLE":
                eg_eliminatorio_no_cumple = True

        eg_results.append({
            "id": ln.get("id", f"EG.{i+1}"),
            "estado": r_eg.get("estado"),
            "anios_detectados": r_eg.get("anios_detectados"),
            "evidencia": r_eg.get("evidencia"),
            "justificacion": r_eg.get("justificacion"),
            "confianza": r_eg.get("confianza"),
            "modo_evaluacion": modo,
            "valor": valor,
            "puntaje": puntaje,
            "_meta": {
                "criterio_row": criterio_eg_row,
                "evidencia_source": "parsed_postulante.jsonl",
                "modelo": r_eg["_llm_meta"]["model"],
                "timestamp": r_eg["_llm_meta"]["timestamp"],
            }
        })

    # ---------------------------------------------------------------------

    return {
        "dni": dni,
        "nombre_full": nombre,
        "FA": {
            "estado": fa.get("estado"),
            "evidencia": fa.get("evidencia"),
            "justificacion": fa.get("justificacion"),
            "confianza": fa.get("confianza"),
            "eliminatorio": True,
        },
        "EC": {
            "blocks": ec_results,
            "puntaje_total": ec_puntaje_total,
            "eliminatorio_no_cumple": ec_eliminatorio_no_cumple
        },
        # NUEVO: EG agregado (no toca FA/EC)
        "EG": {
            "lines": eg_results,
            "puntaje_total": eg_puntaje_total,
            "eliminatorio_no_cumple": eg_eliminatorio_no_cumple
        },
        "_meta": {
            "evaluated_at": ts(),
            "fa_criterio_row": criterio_fa_row,
            "modelo": fa["_llm_meta"]["model"],
            "timestamp": fa["_llm_meta"]["timestamp"],
        }
    }

def resumen_proceso(resultados: List[Dict[str, Any]]) -> Dict[str, Any]:
    total = len(resultados)
    fa_cumple = sum(1 for r in resultados if (r.get("FA", {}) or {}).get("estado") == "CUMPLE")
    fa_no = sum(1 for r in resultados if (r.get("FA", {}) or {}).get("estado") == "NO_CUMPLE")
    fa_inf = sum(1 for r in resultados if (r.get("FA", {}) or {}).get("estado") == "INFO_INSUFICIENTE")

    ec_elim = sum(1 for r in resultados if (r.get("EC", {}) or {}).get("eliminatorio_no_cumple") is True)

    # top puntajes EC
    top = sorted(
        [{"dni": r.get("dni",""), "nombre_full": r.get("nombre_full",""), "ec_puntaje": (r.get("EC",{}) or {}).get("puntaje_total", 0)}
         for r in resultados],
        key=lambda x: x["ec_puntaje"],
        reverse=True
    )[:10]

    return {
        "total_postulantes": total,
        "FA": {"CUMPLE": fa_cumple, "NO_CUMPLE": fa_no, "INFO_INSUFICIENTE": fa_inf},
        "EC": {"eliminatorio_no_cumple": ec_elim},
        "top_10_ec_puntaje": top,
        "generated_at": ts()
    }

def write_value_safe(ws, row: int, col: int, value):
    """
    Escribe en (row,col) aunque haya merges:
    si cae en un MergedCell, escribe en el anchor (top-left) del merge.
    """
    cell = ws.cell(row=row, column=col)
    if not isinstance(cell, MergedCell):
        cell.value = value
        return

    coord = cell.coordinate
    for r in ws.merged_cells.ranges:
        if coord in r:
            ws.cell(row=r.min_row, column=r.min_col).value = value
            return

    # si no detecta el merge (raro), no escribe
    return

def _estado_to_excel(v: str) -> str:
    v = (v or "").strip().upper()
    if v in ("CUMPLE", "NO_CUMPLE", "INFO_INSUFICIENTE"):
        return v.replace("_", " ")
    return v

def _safe_int(x, default=0) -> int:
    try:
        if x is None:
            return default
        s = str(x).strip()
        return int(s) if s.isdigit() else default
    except Exception:
        return default

def write_eval_postulante_to_excel(
    ws,
    score_col: int,
    criteria: Dict[str, Any],
    result: Dict[str, Any],
):
    """
    ws: hoja 'Evaluación CV'
    score_col: columna de puntaje del slot (G, I, K...) en número (7,9,11...)
    criteria: tu criteria_evaluacion.json (dict)
    result: salida de eval_one_postulante (dict con FA/EC/EG/EE)
    """

    total = 0

    # -------------------------
    # FA
    # -------------------------
    fa_row = criteria["criterios"]["FA"]["criterio_item"].get("row")
    if fa_row:
        fa_estado = _estado_to_excel(result.get("FA", {}).get("estado"))
        write_value_safe(ws, fa_row, score_col, fa_estado)

    # -------------------------
    # EC
    # -------------------------
    ec_blocks = criteria["criterios"]["EC"]["blocks"]
    ec_res = (result.get("EC", {}) or {}).get("blocks", []) or []

    for i, cb in enumerate(ec_blocks):
        rrow = cb["criterio_item"].get("row")
        modo = (cb.get("modo_evaluacion") or "").strip()
        valor = cb.get("valor")

        estado = None
        puntaje = 0

        if i < len(ec_res):
            estado = (ec_res[i].get("estado") or "").strip().upper()

        if rrow:
            if modo.lower() == "puntaje":
                # Si cumple => suma puntaje del criterio, si no => 0
                puntaje = _safe_int(valor, 0) if estado == "CUMPLE" else 0
                write_value_safe(ws, rrow, score_col, puntaje)
                total += puntaje
            else:
                write_value_safe(ws, rrow, score_col, _estado_to_excel(estado))

    # -------------------------
    # EG (Experiencia General)
    # -------------------------
    eg_lines = criteria["criterios"]["EG"]["lines"]
    eg_res = (result.get("EG", {}) or {}).get("lines", []) or []

    for i, line in enumerate(eg_lines):
        rrow = line["criterio_item"].get("row")
        modo = (line.get("modo_evaluacion") or "").strip()
        valor = line.get("valor")

        estado = None
        if i < len(eg_res):
            estado = (eg_res[i].get("estado") or "").strip().upper()

        if rrow:
            if modo.lower() == "puntaje":
                puntaje = _safe_int(valor, 0) if estado == "CUMPLE" else 0
                write_value_safe(ws, rrow, score_col, puntaje)
                total += puntaje
            else:
                write_value_safe(ws, rrow, score_col, _estado_to_excel(estado))

    # -------------------------
    # EE (Experiencia Específica)
    # -------------------------
    ee_lines = criteria["criterios"]["EE"]["lines"]
    ee_res = (result.get("EE", {}) or {}).get("lines", []) or []

    for i, line in enumerate(ee_lines):
        rrow = line["criterio_item"].get("row")
        modo = (line.get("modo_evaluacion") or "").strip()
        valor = line.get("valor")

        estado = None
        if i < len(ee_res):
            estado = (ee_res[i].get("estado") or "").strip().upper()

        if rrow:
            if modo.lower() == "puntaje":
                puntaje = _safe_int(valor, 0) if estado == "CUMPLE" else 0
                write_value_safe(ws, rrow, score_col, puntaje)
                total += puntaje
            else:
                write_value_safe(ws, rrow, score_col, _estado_to_excel(estado))

    # -------------------------
    # TOTAL (en tu JSON: stop_row = 22)
    # -------------------------
    stop_row = (criteria.get("_meta") or {}).get("stop_row")
    if stop_row:
        write_value_safe(ws, int(stop_row), score_col, total)

    return total

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True, help="Ruta raíz con carpetas de procesos")
    ap.add_argument("--only-proc", default="", help="Nombre que contenga el proceso (filtro)")
    ap.add_argument("--limit", type=int, default=0, help="Limitar postulantes por proceso (0=sin límite)")
    ap.add_argument("--debug", action="store_true")
    ap.add_argument("--write-resumen", action="store_true", help="Genera evaluacion_resumen.json")
    args = ap.parse_args()

    root = Path(args.root)
    if not root.exists():
        raise SystemExit(f"No existe root: {root}")

    only = norm(args.only_proc).lower()

    procesos = [p for p in root.iterdir() if p.is_dir()]
    procesos.sort(key=lambda p: p.name.lower())

    print(f"[task_41] root={root} procesos={len(procesos)}")

    ok = 0
    skip = 0
    fail = 0

    for proc_dir in procesos:
        proceso = proc_dir.name
        if only and only not in proceso.lower():
            continue

        out_dir = proc_dir / OUT_FOLDER_NAME
        criteria_path = out_dir / PROCESADOS_SUBFOLDER / IN_CRITERIA
        consolidado_path = out_dir / IN_CONSOLIDADO

        print(f"\n[task_41] PROCESO: {proceso}")
        print(f"[task_41]   out_dir={out_dir}")

        if not out_dir.exists():
            print(f"[task_41]   SKIP: no existe 011")
            skip += 1
            continue
        if not criteria_path.exists():
            print(f"[task_41]   SKIP: falta {IN_CRITERIA}")
            skip += 1
            continue
        if not consolidado_path.exists():
            print(f"[task_41]   SKIP: falta {IN_CONSOLIDADO}")
            skip += 1
            continue

        try:
            criteria = json.loads(criteria_path.read_text(encoding="utf-8"))
            postulantes = read_jsonl(consolidado_path)

            if args.limit and args.limit > 0:
                postulantes = postulantes[: args.limit]

            print(f"[task_41]   postulantes={len(postulantes)}")

            resultados = []
            #for idx, p in enumerate(postulantes, start=1):
            #    nombre = p.get("nombre_full", "(sin nombre)")
            #    print(f"[task_41]     ({idx}/{len(postulantes)}) evaluando => {nombre}")
            #    resultados.append(eval_one_postulante(p, criteria, debug=args.debug))

#########################################

            for idx, p in enumerate(postulantes, start=1):
                p = ensure_postulante_dict(p)  # tu normalizador

                nombre = p.get("nombre_full", "(sin nombre)")
                okok, motivo = is_evaluable_postulante(p)

                print(f"[task_41]     ({idx}/{len(postulantes)}) evaluando => {nombre}")

                if not okok:
                    print(f"[task_41]       SKIP => {motivo}")
                    resultados.append({
                        "dni": p.get("dni", ""),
                        "nombre_full": nombre,
                        "estado": "NO_EVALUABLE",
                        "motivo": motivo,
                        "source_file": p.get("source_file") or (p.get("_meta") or {}).get("ruta", ""),
                    })
                    continue

                resultados.append(eval_one_postulante(p, criteria, debug=args.debug))

##########################################

            out_eval = out_dir / OUT_EVAL
            write_jsonl(out_eval, resultados)
            print(f"[task_41]  generado: {out_eval.name}")

            if args.write_resumen:
                out_res = out_dir / OUT_RESUMEN
                write_json(out_res, resumen_proceso(resultados))
                print(f"[task_41]  generado: {out_res.name}")

            ok += 1

        except Exception as e:
            print(f"[task_41]   FAIL: {repr(e)}")
            fail += 1

    print("")
    print(f"[task_41] resumen: OK={ok} SKIP={skip} FAIL={fail}")


if __name__ == "__main__":
    main()
