# core/llm_client.py
import os
import json
from datetime import datetime

import re
from typing import Any, Dict

import certifi
import httpx

from openai import OpenAI
from dotenv import load_dotenv

load_dotenv()

MODEL = os.getenv("OPENAI_MODEL", "gpt-4.1-mini")
API_KEY = os.getenv("OPENAI_API_KEY", "").strip()

if not API_KEY:
    raise RuntimeError(
        "OPENAI_API_KEY no está configurado. "
        "Define OPENAI_API_KEY en tu .env o en variables de entorno."
    )

#_client = OpenAI(api_key=API_KEY)

_http = httpx.Client(verify=certifi.where(), timeout=60.0)
_client = OpenAI(api_key=API_KEY, http_client=_http)


def _json_only_system_prompt() -> str:
    return (
        "Eres un evaluador técnico de procesos de selección pública.\n"
        "Responde únicamente en JSON válido.\n"
        "No inventes información.\n"
        "Si no hay evidencia suficiente, responde INFO_INSUFICIENTE.\n"
        "No agregues texto fuera del JSON."
    )


def _parse_json_or_fail(content: str) -> dict:
    try:
        return json.loads(content)
    except Exception as e:
        raise ValueError(f"Respuesta LLM no es JSON válido:\n{content}") from e



def parse_llm_json(text: str) -> Dict[str, Any]:
    if text is None:
        raise ValueError("Respuesta LLM vacía")

    s = text.strip()

    # Quitar fences tipo ```json ... ``` o ``` ... ```
    if s.startswith("```"):
        s = re.sub(r"^```(?:json)?\s*", "", s, flags=re.IGNORECASE).strip()
        s = re.sub(r"\s*```$", "", s).strip()

    # Intento directo
    try:
        obj = json.loads(s)
        if not isinstance(obj, dict):
            raise ValueError(f"JSON no es objeto dict, es {type(obj).__name__}")
        return obj
    except Exception:
        pass

    # Buscar primer objeto { ... } dentro del texto
    m = re.search(r"\{[\s\S]*\}", s)
    if m:
        obj = json.loads(m.group(0))
        if not isinstance(obj, dict):
            raise ValueError(f"JSON embebido no es dict, es {type(obj).__name__}")
        return obj

    raise ValueError(f"Respuesta LLM no es JSON válido:\n{s}")



def evaluar_formacion(criterio_text: str, formacion_postulante: str, debug: bool = False) -> dict:
    """
    Evalúa Formación Académica (CUMPLE / NO_CUMPLE / INFO_INSUFICIENTE)
    """
    system_prompt = _json_only_system_prompt()

    user_prompt = f"""
                CRITERIO OFICIAL (FORMACIÓN ACADÉMICA):
                \"\"\"{criterio_text}\"\"\"

                INFORMACIÓN DEL POSTULANTE:
                \"\"\"{formacion_postulante}\"\"\"

                INSTRUCCIONES:
                - Determina si el postulante CUMPLE, NO_CUMPLE o INFO_INSUFICIENTE.
                - Cita literalmente la evidencia si existe.
                - Justifica brevemente.

                Devuelve SOLO este JSON:
                {{
                "estado": "CUMPLE | NO_CUMPLE | INFO_INSUFICIENTE",
                "evidencia": "...",
                "justificacion": "...",
                "confianza": 0.0
                }}
                """.strip()

    resp = _client.chat.completions.create(
        model=MODEL,
        temperature=0.0,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
    )

    content = resp.choices[0].message.content.strip()
    #data = _parse_json_or_fail(content)
    data = parse_llm_json(content)

    data["_llm_meta"] = {
        "model": MODEL,
        "timestamp": datetime.now().isoformat(timespec="seconds"),
    }

    if debug:
        data["_debug"] = {
            "criterio": criterio_text,
            "input": formacion_postulante,
            "raw": content,
        }

    return data

def evaluar_estudios_complementarios(criterio_text: str, evidencia_postulante: str, debug: bool = False) -> dict:
    """
    Evalúa Estudios Complementarios para 1 bloque.
    """
    system_prompt = _json_only_system_prompt()

    user_prompt = f"""
        CRITERIO OFICIAL (ESTUDIOS COMPLEMENTARIOS):
        \"\"\"{criterio_text}\"\"\"

        EVIDENCIA DEL POSTULANTE:
        \"\"\"{evidencia_postulante}\"\"\"

        INSTRUCCIONES:
        - Determina si el postulante CUMPLE, NO_CUMPLE o INFO_INSUFICIENTE.
        - Cita evidencia literal si existe.
        - Justifica brevemente.

        Devuelve SOLO este JSON:
        {{
        "estado": "CUMPLE | NO_CUMPLE | INFO_INSUFICIENTE",
        "evidencia": "...",
        "justificacion": "...",
        "confianza": 0.0
        }}
        """.strip()

    resp = _client.chat.completions.create(
        model=MODEL,
        temperature=0.0,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
    )

    content = resp.choices[0].message.content.strip()
    #data = _parse_json_or_fail(content)
    data = parse_llm_json(content)

    data["_llm_meta"] = {
        "model": MODEL,
        "timestamp": datetime.now().isoformat(timespec="seconds"),
    }

    if debug:
        data["_debug"] = {
            "criterio": criterio_text,
            "input": evidencia_postulante,
            "raw": content,
        }

    return data

def evaluar_experiencia_general(
    criterio_text: str,
    evidencia_postulante: str,
    debug: bool = False
) -> dict:
    """
    Evalúa Experiencia General (CUMPLE / NO_CUMPLE / INFO_INSUFICIENTE)

    evidencia_postulante debe venir ya preprocesada por tu pipeline (task_20/task_40),
    por ejemplo:
    - lista de experiencias (empresa | cargo | fechas)
    - total años calculados SIN solapamiento
    """

    system_prompt = _json_only_system_prompt()

    user_prompt = f"""
                CRITERIO OFICIAL (EXPERIENCIA GENERAL):
                \"\"\"{criterio_text}\"\"\"

                EVIDENCIA DEL POSTULANTE (ya consolidada):
                \"\"\"{evidencia_postulante}\"\"\"

                INSTRUCCIONES:
                - Determina si el postulante CUMPLE, NO_CUMPLE o INFO_INSUFICIENTE.
                - NO inventes fechas, cargos ni años.
                - Si la evidencia contiene "total_anios_calc" o "total_dias_calc", úsalo como referencia principal.
                - Si hay conflicto entre evidencia textual y total_anios_calc, prioriza total_anios_calc.
                - Cita evidencia literal.
                - Justifica brevemente.

                Devuelve SOLO este JSON:
                {{
                "estado": "CUMPLE | NO_CUMPLE | INFO_INSUFICIENTE",
                "anios_detectados": 0.0,
                "evidencia": "...",
                "justificacion": "...",
                "confianza": 0.0
                }}
                """.strip()

    resp = _client.chat.completions.create(
        model=MODEL,
        temperature=0.0,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
    )

    content = resp.choices[0].message.content.strip()
    #data = _parse_json_or_fail(content)
    data = parse_llm_json(content)

    data["_llm_meta"] = {
        "model": MODEL,
        "timestamp": datetime.now().isoformat(timespec="seconds"),
    }

    if debug:
        data["_debug"] = {
            "criterio": criterio_text,
            "input": evidencia_postulante,
            "raw": content,
        }

    return data

def evaluar_experiencia_especifica(
    criterio_text: str,
    evidencia_postulante: str,
    debug: bool = False
) -> dict:
    """
    Evalúa Experiencia Específica (CUMPLE / NO_CUMPLE / INFO_INSUFICIENTE)

    evidencia_postulante debe venir ya preprocesada por tu pipeline (task_20/task_40),
    idealmente con:
    - experiencias filtradas como "específicas"
    - total años calculados SIN solapamiento
    """

    system_prompt = _json_only_system_prompt()

    user_prompt = f"""
                CRITERIO OFICIAL (EXPERIENCIA ESPECÍFICA):
                \"\"\"{criterio_text}\"\"\"

                EVIDENCIA DEL POSTULANTE (ya consolidada):
                \"\"\"{evidencia_postulante}\"\"\"

                INSTRUCCIONES:
                - Determina si el postulante CUMPLE, NO_CUMPLE o INFO_INSUFICIENTE.
                - NO inventes información.
                - Usa "total_anios_calc" o "total_dias_calc" si están presentes.
                - Si no hay evidencia clara de funciones relacionadas (analista funcional / analista de sistemas / similares),
                responde INFO_INSUFICIENTE aunque haya años.
                - Cita evidencia literal.
                - Justifica brevemente.

                Devuelve SOLO este JSON:
                {{
                "estado": "CUMPLE | NO_CUMPLE | INFO_INSUFICIENTE",
                "anios_detectados": 0.0,
                "evidencia": "...",
                "justificacion": "...",
                "confianza": 0.0
                }}
                """.strip()

    resp = _client.chat.completions.create(
        model=MODEL,
        temperature=0.0,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
    )

    content = resp.choices[0].message.content.strip()
    #data = _parse_json_or_fail(content)
    data = parse_llm_json(content)

    data["_llm_meta"] = {
        "model": MODEL,
        "timestamp": datetime.now().isoformat(timespec="seconds"),
    }

    if debug:
        data["_debug"] = {
            "criterio": criterio_text,
            "input": evidencia_postulante,
            "raw": content,
        }

    return data
