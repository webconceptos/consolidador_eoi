# core/llm_client.py
import os
import json
from datetime import datetime

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

_client = OpenAI(api_key=API_KEY)


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
    data = _parse_json_or_fail(content)

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
    data = _parse_json_or_fail(content)

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
