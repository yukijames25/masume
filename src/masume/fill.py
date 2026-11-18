"""Fill detected form fields with values and write a new .xlsx."""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from . import ooxml
from .detect import CHECK_CHARS, CHECKED_CHARS, Field, _DATE_PLACEHOLDER, _INLINE_BLANK, detect, normalize_label
from .ooxml import Workbook, parse_ref
from .wareki import parse_date, to_wareki, wareki_year_str


class FillError(Exception):
    pass


@dataclass
class FillReport:
    written: list[dict] = field(default_factory=list)
    unknown_keys: list[str] = field(default_factory=list)
    unfilled: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "written": self.written,
            "unknown_keys": self.unknown_keys,
            "unfilled": self.unfilled,
            "errors": self.errors,
            "ok": not self.errors and not self.unknown_keys,
        }


def _uses_wareki(wb: Workbook, sheet_name: str) -> bool:
    sheet = wb.sheet(sheet_name)
    for c in sheet.cells.values():
        if isinstance(c.value, str) and re.search(r"令和|平成|昭和", c.value):
            return True
    return False


def format_date(value, template: str | None, wareki: bool) -> str:
    d = parse_date(value)
    if d is None:
        return str(value)
    if template:
        m = _DATE_PLACEHOLDER.match(template)
        if m:
            era = m.group("era")
            pre = m.group("pre") or ""
            post = m.group("post") or ""
            if era == "西暦" or (era is None and not wareki):
                body = f"{d.year}年{d.month}月{d.day}日"
                if era == "西暦":
                    body = "西暦" + body
            else:
                name, y = to_wareki(d)
                body = f"{name}{wareki_year_str(y)}年{d.month}月{d.day}日"
            return f"{pre}{body}{post}"
    if wareki:
        name, y = to_wareki(d)
        return f"{name}{wareki_year_str(y)}年{d.month}月{d.day}日"
    return f"{d.year}年{d.month}月{d.day}日"


def format_choice(template: str, selected, mark: str = "■") -> str:
    if isinstance(selected, str):
        selected = [selected]
    selected = [normalize_label(s) for s in selected]
    out = []
    i = 0
    found: set[str] = set()
    boxes = CHECK_CHARS + CHECKED_CHARS
    while i < len(template):
        ch = template[i]
        if ch in boxes:
            m = re.match(r"[" + boxes + r"]\s*([^" + boxes + r"\s　（）()、,，]+)", template[i:])
            opt = normalize_label(m.group(1)) if m else ""
            if opt in selected:
                out.append(mark)
                found.add(opt)
            else:
                out.append("□")
            i += 1
            continue
        out.append(ch)
        i += 1
    missing = [s for s in selected if s not in found]
    if missing:
        raise FillError(f"option(s) not in form: {missing}")
    return "".join(out)


def format_inline(template: str, value) -> str:
    m = _INLINE_BLANK.match(template)
    if not m:
        return str(value)
    start, end = m.span("blank")
    return template[:start] + str(value) + template[end:]


def resolve(fields: list[Field], key: str) -> Field | None:
    for f in fields:
        if f.id == key:
            return f
    nk = normalize_label(key)
    cand = [f for f in fields if normalize_label(f.label) == nk]
    if len(cand) == 1:
        return cand[0]
    cand = [f for f in fields if f.id.split("/")[-1].split("#")[0] == nk]
    if len(cand) == 1:
        return cand[0]
    return None


_CELL_KEY = re.compile(r"^(?:(?P<sheet>[^!]+)!)?\$?(?P<ref>[A-Za-z]{1,3}\$?\d+)$")


def fill(
    src: str,
    dst: str,
    data: dict,
    sheet: str | None = None,
    mark: str = "■",
    strict: bool = False,
    fields: list[Field] | None = None,
) -> FillReport:
    """Fill *src* into *dst*. Pass *fields* (e.g. a corrected, saved field
    definition list) to skip automatic detection."""
    wb = ooxml.load(src)
    if fields is None:
        fields = detect(wb, sheet)
    report = FillReport()
    updates: dict[str, dict[tuple[int, int], object]] = {}
    sheet_path = {s.name: s.path for s in wb.sheets}
    default_sheet = wb.sheet(sheet).name
    touched: set[str] = set()

    def put(sheet_name: str, ref: str, value, key: str, fid: str | None):
        r, c = parse_ref(ref)
        cell = wb.sheet(sheet_name).cells.get((r, c))
        if cell is not None and cell.formula is not None:
            report.errors.append(f"{key}: {sheet_name}!{ref} is a formula cell")
            return
        updates.setdefault(sheet_path[sheet_name], {})[(r, c)] = value
        report.written.append({"key": key, "field": fid, "cell": f"{sheet_name}!{ref}", "value": value})

    for key, value in data.items():
        f = resolve(fields, key)
        if f is None:
            m = _CELL_KEY.match(key)
            if m and (m.group("sheet") is None or m.group("sheet") in sheet_path):
                put(m.group("sheet") or default_sheet, m.group("ref").replace("$", ""), value, key, None)
            else:
                report.unknown_keys.append(key)
            continue
        touched.add(f.id)
        try:
            if f.kind == "table":
                _fill_table(f, value, put, key)
                continue
            if value is None:
                v = None
            elif f.kind == "inline_date":
                v = format_date(value, f.template, _uses_wareki(wb, f.sheet))
            elif f.kind == "date":
                v = format_date(value, None, _uses_wareki(wb, f.sheet)) if parse_date(value) else value
            elif f.kind == "choice":
                v = format_choice(f.template or "", value, mark)
            elif f.kind == "inline_text":
                v = format_inline(f.template or "", value)
            else:
                v = value
            put(f.sheet, f.cell, v, key, f.id)
        except FillError as e:
            report.errors.append(f"{key}: {e}")

    report.unfilled = [f.id for f in fields if f.id not in touched]
    if strict and report.unknown_keys:
        raise FillError(f"unknown keys: {report.unknown_keys}")
    if not report.errors:
        try:
            ooxml.save_with_updates(src, dst, updates)
        except ooxml.WriteError as e:
            report.errors.append(str(e))
    return report


def _fill_table(f: Field, value, put, key: str) -> None:
    if not isinstance(value, list):
        raise FillError("table value must be a list of rows")
    if len(value) > len(f.rows):
        raise FillError(f"table has room for {len(f.rows)} rows, got {len(value)}")
    col_ids = [c["id"] for c in f.columns]
    for i, row in enumerate(value):
        r = f.rows[i]
        if isinstance(row, dict):
            items = []
            for k, v in row.items():
                nk = normalize_label(k)
                if nk not in col_ids:
                    raise FillError(f"unknown column {k!r} (have {col_ids})")
                items.append((col_ids.index(nk), v))
        elif isinstance(row, list):
            if len(row) > len(col_ids):
                raise FillError(f"row {i + 1} has {len(row)} values, table has {len(col_ids)} columns")
            items = list(enumerate(row))
        else:
            raise FillError("table rows must be objects or arrays")
        for ci, v in items:
            col = f.columns[ci]["col"]
            put(f.sheet, f"{col}{r}", v, f"{key}[{i}].{col_ids[ci]}", f.id)
