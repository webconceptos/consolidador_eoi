
# exporters/excel_exporter.py
# -*- coding: utf-8 -*-
from __future__ import annotations

from pathlib import Path
from typing import Dict, Any, List

from openpyxl import Workbook


def export_consolidado_to_excel(items: List[Dict[str, Any]], out_path: Path):
    """
    Export simple: 1 fila por postulante.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "consolidado"

    header = [
        "proceso", "carpeta_postulante", "archivo", "tipo",
        "dni", "nombre_full", "email", "celular",
        "exp_general_dias", "exp_especifica_dias",
        "warnings"
    ]
    ws.append(header)

    for it in items:
        meta = (it.get("_meta") or {})
        ws.append([
            meta.get("proceso", ""),
            meta.get("carpeta_postulante", ""),
            meta.get("archivo", ""),
            meta.get("tipo", ""),
            it.get("dni", ""),
            it.get("nombre_full", ""),
            it.get("email", ""),
            it.get("celular", ""),
            it.get("exp_general_dias", 0),
            it.get("exp_especifica_dias", 0),
            it.get("_parse_warnings", ""),
        ])

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
