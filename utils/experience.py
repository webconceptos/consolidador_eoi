
from __future__ import annotations
from dataclasses import dataclass
from datetime import date, datetime
from typing import Iterable, List, Optional, Tuple


def to_date(x) -> Optional[date]:
    if x is None:
        return None
    if isinstance(x, date) and not isinstance(x, datetime):
        return x
    if isinstance(x, datetime):
        return x.date()
    if isinstance(x, str):
        s = x.strip()
        if not s:
            return None
        for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(s, fmt).date()
            except Exception:
                pass
    return None


def normalize_interval(start: date, end: date) -> Tuple[date, date]:
    # Asegura start <= end
    return (start, end) if start <= end else (end, start)


def merge_intervals(intervals: Iterable[Tuple[date, date]]) -> List[Tuple[date, date]]:
    ints = []
    for a, b in intervals:
        if a is None or b is None:
            continue
        a, b = normalize_interval(a, b)
        ints.append((a, b))
    if not ints:
        return []
    ints.sort(key=lambda t: t[0])
    merged = [ints[0]]
    for s, e in ints[1:]:
        ps, pe = merged[-1]
        # Overlap or contiguous? La macro suele eliminar superposición;
        # para experiencia en días, consideramos contiguo como unión también.
        if s <= pe:
            if e > pe:
                merged[-1] = (ps, e)
        else:
            merged.append((s, e))
    return merged


def total_days(intervals: Iterable[Tuple[date, date]], inclusive: bool = True) -> int:
    """
    Suma días de la unión de intervalos evitando duplicar superposiciones.
    inclusive=True cuenta ambos extremos (end-start+1), típico en conteos por días calendario.
    """
    merged = merge_intervals(intervals)
    days = 0
    for s, e in merged:
        d = (e - s).days + (1 if inclusive else 0)
        if d > 0:
            days += d
    return days


def ymd_from_days(days: int) -> Tuple[int, int, int]:
    """
    Replica la lógica típica de estas hojas: años=INT(dias/365),
    meses=INT((dias - años*365)/30), dias_restantes=...
    """
    if days < 0:
        days = 0
    years = days // 365
    rem = days - years * 365
    months = rem // 30
    rem2 = rem - months * 30
    return int(years), int(months), int(rem2)
