
# parsers/eoi_pdf.py
# -*- coding: utf-8 -*-
"""
Parser EOI PDF (placeholder)

Aquí solo dejamos una estructura mínima para no romper Task 20.
Si ya tienes un parser real, reemplaza este archivo.
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, Any

def parse_eoi_pdf(path: Path, use_ocr: bool = False) -> Dict[str, Any]:
    # Placeholder: no inventamos datos
    return {
        "dni": "",
        "nombre_full": "",
        "email": "",
        "celular": "",
        "exp_general_dias": 0,
        "exp_especifica_dias": 0,
        "experiencias": [],
        "cursos": [],
        "_parse_warnings": "PDF_PARSER_PLACEHOLDER",
    }
