# tasks/task_41_eval_postulantes_openai.py
# -*- coding: utf-8 -*-

import argparse
import json
import re
from pathlib import Path
from datetime import datetime

from core.llm_client import evaluar_formacion, evaluar_estudios_complementarios


OUT_NAME = "evaluacion_postulantes.jsonl"

def read_jsonl(path: Path) -> list[dict]:
    rows = []
    with path.open("r", encoding="utf-8") as f:
        for line in f:
            if line.strip():
                rows.append(json.loads(line))
    return rows

def write_jsonl(path: Path, rows: list[dict]):
    with path.open("w", encoding="utf-8") as f:
        for r in rows:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")

def split_b_blocks(text: str) -> dict[str, str]:
    """
    Retorna {"B.1": "...", "B.2": "..."} si detecta.
    Si no detecta, retorna {}.
    Acepta B.1, B1, B.1:, etc.
    """
    if not text:
        return {}
    t = text.replace("\r\n", "\n").strip()

    # Detecta encabezados: B.1 o B1, con o sin ":".
    rx = re.compile(r"^\s*(B\.?\s*\d+)\s*:?\s*$", re.IGNORECASE)

    blocks: dict[str, list[str]] = {}
    current = None

    for line in t.split("\n"):
        m = rx.match(line.strip())
        if m:
            key = m.group(1).upper().replace(" ", "")
            if key.startswith("B") and not key.startswith("B."):
                key = "B." + key[1:]
            current = key
            blocks.setdefault(current, [])
            continue

        if current:
            blocks[current].append(line)

    # limpia
    out = {}
    for k, lines in blocks.items():
        body = "\n".join(lines).strip()
        # colapsa triples saltos, pero conserva doble salto
        body = re.sub(r"\n{3,}", "\n\n", body).strip()
        out[k] = body
    return out

def get_ec_blocks_from_postulante(p: dict) -> list[dict]:
    ec = p.get("estudios_complementarios") or {}
    blocks = ec.get("blocks")
    if isinstance(blocks, list):
        out = []
        for b in blocks:
            if isinstance(b, dict):
                out.append({
                    "id": str(b.get("id","")).strip(),
                    "title": str(b.get("title","")).strip(),
                    "resumen": str(b.get("resumen","")).strip(),
                })
        return out
    return []

def get_ec_fallback_text(p: dict) -> str:
    v = p.get("estudios_complementarios_resumen")
    if isinstance(v, str) and v.strip():
        return v.strip()
    return ""

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--criteria", required=True, help="Ruta a criteria_evaluacion.json")
    ap.add_argument("--postulantes", required=True, help="Ruta a parsed_postulantes.jsonl")
    ap.add_argument("--out", default="", help="Ruta de salida (jsonl)")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--debug", action="store_true")
    args = ap.parse_args()

    criteria = json.loads(Path(args.criteria).read_text(encoding="utf-8"))
    postulantes = read_jsonl(Path(args.postulantes))


    if args.limit > 0:
        postulantes = postulantes[: args.limit]

    criterio_fa = criteria["criterios"]["FA"]["criterio_item"]["text"]
    criterio_row = criteria["criterios"]["FA"]["criterio_item"]["row"]

    criteria_ec_blocks = criteria["criterios"]["EC"]["blocks"]

    resultados = []

    for p in postulantes:

        # ---------------- FA ----------------        
        print(f"[task_41] Evaluando FA => {p.get('nombre_full','(sin nombre)')}")

        formacion_obligatoria = p.get("formacion_obligatoria", "") 
        formacion=formacion_obligatoria.get("resumen","")

        r_llm = evaluar_formacion(
            criterio_text=criterio_fa,
            formacion_postulante=formacion,
            debug=args.debug,
        )


        # ---------------- EC ----------------
        ec_blocks_post = get_ec_blocks_from_postulante(p)
        ec_fallback = get_ec_fallback_text(p)

        ec_results = []
        ec_eliminatorio_no_cumple = False
        ec_puntaje_total = 0

        for i, cb in enumerate(criteria_ec_blocks):
            criterio_ec = cb["criterio_item"]["text"]
            criterio_ec_row = cb["criterio_item"]["row"]
            modo = cb.get("modo_evaluacion")
            valor = cb.get("valor")

            # evidencia: prioridad blocks por índice, sino fallback
            evidencia = ""
            ev_source = ""
            if i < len(ec_blocks_post) and ec_blocks_post[i].get("resumen"):
                evidencia = ec_blocks_post[i]["resumen"]
                ev_source = f"blocks[{i}] id={ec_blocks_post[i].get('id')}"
            else:
                evidencia = ec_fallback
                ev_source = "fallback_text"

            r_ec = evaluar_estudios_complementarios(
                criterio_text=criterio_ec,
                evidencia_postulante=evidencia,
                debug=args.debug
            )

            # regla de puntaje (simple y controlada)
            puntaje = 0
            if modo == "Puntaje":
                try:
                    vnum = int(valor) if str(valor).isdigit() else 0
                except:
                    vnum = 0
                puntaje = vnum if r_ec.get("estado") == "CUMPLE" else 0
                ec_puntaje_total += puntaje
            else:
                # Cumple / No Cumple es eliminatorio en EC.1 (y en general si así lo defines)
                if r_ec.get("estado") == "NO_CUMPLE" and str(valor).upper() == "CUMPLE_NOCUMPLE":
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
                    "modelo": r_ec["_llm_meta"]["model"],
                    "timestamp": r_ec["_llm_meta"]["timestamp"],
                    "evidencia_source": ev_source,
                }
            })



        result = {
            "dni": p.get("dni"),
            "nombre_full": p.get("nombre_full"),
            "FA": {
                "estado": r_llm.get("estado"),
                "evidencia": r_llm.get("evidencia"),
                "justificacion": r_llm.get("justificacion"),
                "confianza": r_llm.get("confianza"),
                "eliminatorio": True,
            },
            "EC":{
                "blocks": ec_results,
                "puntaje_total": ec_puntaje_total,
                "eliminatorio_no_cumple": ec_eliminatorio_no_cumple
            },
            "_meta": {
                "criterio_row": criterio_row,
                "criterio_text": criterio_fa,
                "modelo": r_llm["_llm_meta"]["model"],
                "timestamp": r_llm["_llm_meta"]["timestamp"],
            },
        }

        resultados.append(result)

    out_path = Path(args.out) if args.out else Path(args.postulantes).parent / OUT_NAME
    write_jsonl(out_path, resultados)

    print(f"[task_41] ✅ Evaluación FA completada")
    print(f"[task_41] archivo generado: {out_path}")


if __name__ == "__main__":
    main()
