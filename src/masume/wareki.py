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
_ERA_START = {name: start.year for name, start in ERAS}
_ERA_ALIAS = {"令": "令和", "R": "令和", "平": "平成", "H": "平成", "昭": "昭和", "S": "昭和",
              "大": "大正", "T": "大正", "明": "明治", "M": "明治"}
# 「昭和60年6月15日」「平成元年4月2日」「S60.6.15」「H2/4/2」「R8.11.19」
_WAREKI_RE = re.compile(
    r"^\s*(令和|平成|昭和|大正|明治|[令平昭大明]|[RHSTMrhstm])\s*(元|\d{1,2})\s*[年./\-]\s*(\d{1,2})\s*[月./\-]\s*(\d{1,2})\s*日?\s*$"
)


def _z2h(s: str) -> str:
    return s.translate(str.maketrans("０１２３４５６７８９／－．", "0123456789/-."))


def parse_date(value) -> _dt.date | None:
    if isinstance(value, _dt.datetime):
        return value.date()
    if isinstance(value, _dt.date):
        return value
    if not isinstance(value, str):
        return None
    value = _z2h(value)
    m = _ISO_RE.match(value)
    try:
        if m:
            return _dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        m = _WAREKI_RE.match(value)
        if m:
            era = _ERA_ALIAS.get(m.group(1).upper() if len(m.group(1)) == 1 else m.group(1), m.group(1))
            y = 1 if m.group(2) == "元" else int(m.group(2))
            return _dt.date(_ERA_START[era] + y - 1, int(m.group(3)), int(m.group(4)))
    except (ValueError, KeyError):
        return None
    return None


_YM_RE = re.compile(r"^\s*(\d{4})\s*[-/.年]\s*(\d{1,2})\s*月?\s*$")
_WAREKI_YM_RE = re.compile(
    r"^\s*(令和|平成|昭和|大正|明治|[令平昭大明]|[RHSTMrhstm])\s*(元|\d{1,2})\s*[年./\-]\s*(\d{1,2})\s*月?\s*$")


def parse_year_month(value) -> _dt.date | None:
    """「2026-12」「2026年12月」「令和8年12月」「R8.12」 -> the 1st of that month."""
    if not isinstance(value, str):
        return None
    value = _z2h(value)
    try:
        m = _YM_RE.match(value)
        if m:
            return _dt.date(int(m.group(1)), int(m.group(2)), 1)
        m = _WAREKI_YM_RE.match(value)
        if m:
            era = _ERA_ALIAS.get(m.group(1).upper() if len(m.group(1)) == 1 else m.group(1), m.group(1))
            y = 1 if m.group(2) == "元" else int(m.group(2))
            return _dt.date(_ERA_START[era] + y - 1, int(m.group(3)), 1)
    except (ValueError, KeyError):
        return None
    return None


def to_wareki(d: _dt.date) -> tuple[str, int]:
    for name, start in ERAS:
        if d >= start:
            return name, d.year - start.year + 1
    raise ValueError(f"date before Meiji era: {d}")


def wareki_year_str(year: int) -> str:
    return "元" if year == 1 else str(year)
