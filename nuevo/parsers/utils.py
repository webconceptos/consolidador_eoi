
# parsers/utils.py
# -*- coding: utf-8 -*-
from __future__ import annotations

import re
from datetime import datetime, date
from typing import Any, Optional

def norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())

def safe_int(x: Any, default: int = 0) -> int:
    try:
        if x is None:
            return default
        if isinstance(x, bool):
            return int(x)
        if isinstance(x, (int, float)):
            return int(x)
        s = str(x).strip()
        if not s:
            return default
        s = re.sub(r"[^\d\-]", "", s)
        return int(s) if s else default
    except Exception:
        return default

def _parse_date_any(x: Any) -> Optional[datetime]:
    if x is None:
        return None
    if isinstance(x, datetime):
        return x
    if isinstance(x, date):
        return datetime(x.year, x.month, x.day)
    s = norm(str(x))
    if not s:
        return None

    # normaliza separadores
    s2 = s.replace(".", "/").replace("-", "/")
    # formatos comunes: dd/mm/yyyy o yyyy/mm/dd
    m = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{4})", s2)
    if m:
        d, mo, y = map(int, m.groups())
        return datetime(y, mo, d)

    m = re.fullmatch(r"(\d{4})/(\d{1,2})/(\d{1,2})", s2)
    if m:
        y, mo, d = map(int, m.groups())
        return datetime(y, mo, d)

    # ISO parcial
    try:
        return datetime.fromisoformat(s.replace("Z", ""))
    except Exception:
        return None

def _days_between(fi: Optional[datetime], ff: Optional[datetime]) -> int:
    if not fi or not ff:
        return 0
    try:
        delta = ff - fi
        return max(0, int(delta.days))
    except Exception:
        return 0
