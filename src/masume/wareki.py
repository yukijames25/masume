"""Japanese era (和暦) helpers."""
from __future__ import annotations

import datetime as _dt
import re

# (era name, first day)
ERAS = [
    ("令和", _dt.date(2019, 5, 1)),
    ("平成", _dt.date(1989, 1, 8)),
    ("昭和", _dt.date(1926, 12, 25)),
    ("大正", _dt.date(1912, 7, 30)),
    ("明治", _dt.date(1868, 1, 25)),
]

_ISO_RE = re.compile(r"^\s*(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})日?\s*$")


def parse_date(value) -> _dt.date | None:
    if isinstance(value, _dt.datetime):
        return value.date()
    if isinstance(value, _dt.date):
        return value
    if not isinstance(value, str):
        return None
    m = _ISO_RE.match(value)
    if not m:
        return None
    try:
        return _dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None


def to_wareki(d: _dt.date) -> tuple[str, int]:
    for name, start in ERAS:
        if d >= start:
            return name, d.year - start.year + 1
    raise ValueError(f"date before Meiji era: {d}")


def wareki_year_str(year: int) -> str:
    return "元" if year == 1 else str(year)
