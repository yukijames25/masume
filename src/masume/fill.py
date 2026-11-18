"""Fill detected form fields with values and write a new .xlsx."""
from __future__ import annotations

import datetime as _dt
import os
import re
from dataclasses import dataclass, field

from . import ooxml
from .detect import (BOX_CHARS, Field, _INLINE_BLANK, detect, form_sheets, normalize_label,
                     placeholder_dates)
from .ooxml import Workbook, parse_ref
from .wareki import parse_date, to_wareki, wareki_year_str

ERA_ABBR = {"明治": "明", "大正": "大", "昭和": "昭", "平成": "平", "令和": "令"}
_SPLIT = re.compile(r"[-－ー‐−/／:：.．\s　()（）]+")


class FillError(Exception):
    pass


@dataclass
class FillReport:
    written: list[dict] = field(default_factory=list)
    unknown_keys: list[str] = field(default_factory=list)
    unfilled: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    verified: bool | None = None  # every written value read back from the saved file

    def to_dict(self) -> dict:
        saved = not self.errors and self.verified is not None
        return {
            "ok": not self.errors and not self.unknown_keys,
            "saved": saved,
            "verified": self.verified,
            ("written" if saved else "planned"): self.written,
            "unknown_keys": self.unknown_keys,
            "unfilled": self.unfilled,
            "warnings": self.warnings,
            "errors": self.errors,
        }


def _uses_wareki(wb: Workbook, sheet_name: str) -> bool:
    sheet = wb.sheet(sheet_name)
    for c in sheet.cells.values():
        if isinstance(c.value, str) and re.search(r"令和|平成|昭和", c.value):
            return True
    return False


# --------------------------------------------------------------------------
# value formatters
# --------------------------------------------------------------------------
def _era_text(era_tpl: str | None, d: _dt.date, wareki: bool) -> tuple[str, int | None]:
    """Return (era prefix, year to print) — year None means 西暦."""
    if era_tpl == "西暦" or (era_tpl is None and not wareki):
        return ("西暦" if era_tpl == "西暦" else ""), None
    name, y = to_wareki(d)
    return name, y


def _date_text(era: str, y: int | None, d: _dt.date) -> str:
    if y is None:
        return f"{era}{d.year}年{d.month}月{d.day}日"
    return f"{era}{wareki_year_str(y)}年{d.month}月{d.day}日"


def _age(born: _dt.date, today: _dt.date) -> int:
    return today.year - born.year - ((today.month, today.day) < (born.month, born.day))


_RANGE_SEP = re.compile(r"\s*(?:[～〜~]|\s/\s|\s+to\s+|(?<=\d)/(?=\d{4}[-/.年]))\s*")


def format_date(value, template: str | None, wareki: bool, today: _dt.date | None = None) -> str:
    """Render a date (or a from/to pair for ranges) into a placeholder template."""
    if isinstance(value, dict):
        values = [value.get("from") or value.get("start"), value.get("to") or value.get("end")]
    elif isinstance(value, (list, tuple)):
        values = list(value)
    elif isinstance(value, str) and _RANGE_SEP.search(value):
        values = _RANGE_SEP.split(value.strip())
    else:
        values = [value]
    dates = [parse_date(v) for v in values]
    if not dates or any(d is None for d in dates):
        if template and placeholder_dates(template):
            raise FillError(
                f"日付として読めません: {value!r}。'2026-11-19' や '昭和60年6月15日' の形で渡してください"
                "（そのまま書き込みたい場合はセル番地をキーにする）")
        return str(value)
    today = today or _dt.date.today()
    ms = placeholder_dates(template) if template else []
    if not ms:
        era, y = _era_text(None, dates[0], wareki)
        return _date_text(era, y, dates[0])
    if len(dates) > len(ms):
        raise FillError(f"form has {len(ms)} date slot(s), got {len(dates)} dates")
    out = []
    pos = 0
    for m, d in zip(ms, dates):
        era, y = _era_text(m.group("era"), d, wareki)
        out.append(template[pos:m.start()])
        out.append(_date_text(era, y, d))
        pos = m.end()
    out.append(template[pos:])
    text = "".join(out)
    # （満　　歳） -> age,  （　　日間） -> inclusive length of a range
    text = re.sub(r"満[\s　]+歳", lambda _: f"満{_age(dates[0], today)}歳", text, count=1)
    if len(dates) >= 2:
        days = (dates[1] - dates[0]).days + 1
        text = re.sub(r"[（(][\s　]+日間[）)]", f"（{days}日間）", text, count=1)
    return text


def format_choice(template: str, selected, mark: str = "■") -> str:
    extra: dict[str, str] = {}
    if isinstance(selected, dict):  # {"その他": "親戚"}
        extra = {normalize_label(k): str(v) for k, v in selected.items() if v not in (None, True, False, "")}
        selected = [k for k, v in selected.items() if v not in (None, False)]
    if isinstance(selected, str):
        selected = [selected]
    sel = []
    for s in selected:
        m = re.match(r"^(.+?)[（(](.+)[）)]$", str(s))
        if m:  # "その他（親戚）"
            extra[normalize_label(m.group(1))] = m.group(2)
            s = m.group(1)
        sel.append(normalize_label(s))
    out = []
    i = 0
    found: set[str] = set()
    while i < len(template):
        ch = template[i]
        if ch in BOX_CHARS:
            m = re.match(r"[" + BOX_CHARS + r"]\s*([^" + BOX_CHARS + r"\s　（）()、,，]+)", template[i:])
            opt = normalize_label(m.group(1)) if m else ""
            if opt in sel:
                out.append(mark)
                found.add(opt)
                if opt in extra and m:
                    j = i + m.end()
                    pm = re.match(r"[\s　]*[（(][\s　]*[）)]", template[j:])
                    if pm:
                        out.append(template[i + 1:j] + f"（{extra[opt]}）")
                        i = j + pm.end()
                        continue
            else:
                out.append("□")
            i += 1
            continue
        out.append(ch)
        i += 1
    missing = [s for s in sel if s not in found]
    if missing:
        raise FillError(f"option(s) not in form: {missing}")
    return "".join(out)


def format_circle(template: str, options: list[str], selected) -> str:
    """「男 ・ 女」 + "女" -> 「男 ・ （女）」 (cells cannot hold a drawn circle)."""
    sel = normalize_label(selected if isinstance(selected, str) else selected[0])
    norm = [normalize_label(o) for o in options]
    if sel not in norm:
        raise FillError(f"option not in form: {sel!r} (have {options})")
    o = options[norm.index(sel)]
    parts = re.split(r"([・･/／])", template)
    for i, p in enumerate(parts):
        if normalize_label(p) == sel:
            lead = p[: len(p) - len(p.lstrip(" 　"))]
            trail = p[len(p.rstrip(" 　")):]
            parts[i] = f"{lead}（{o}）{trail}"
            break
    return "".join(parts)


def format_inline(template: str, value) -> str:
    m = _INLINE_BLANK.match(template)
    if not m:
        return str(value)
    start, end = m.span("blank")
    return template[:start] + str(value) + template[end:]


def _split_value(value, n: int) -> list:
    if isinstance(value, (list, tuple)):
        return list(value)
    s = str(value).strip()
    parts = [p for p in _SPLIT.split(s) if p != ""]
    if len(parts) != n:
        nums = re.findall(r"\d+", s)
        if len(nums) == n:
            parts = nums
    return parts


def format_pattern(template: str, value) -> str:
    """「〒　　－」+ "123-4567" -> 「〒123－4567」;  「　時　分」+ "9:30" -> 「9時30分」"""
    slots = [(m.start(), m.end()) for m in re.finditer(r"[\s　]+", template)]
    body = template.rstrip(" 　")
    if body and body[-1] in "－-ー‐−" :
        end = len(body)
        slots = [s for s in slots if s[0] < end] + [(end, len(template))]
    parts = _split_value(value, len(slots))
    if "時" in template and parts and re.fullmatch(r"0\d", str(parts[0])):
        parts[0] = str(int(parts[0]))
    if len(parts) == len(slots):
        out, pos = [], 0
        for (a, b), p in zip(slots, parts):
            out.append(template[pos:a])
            out.append(str(p))
            pos = b
        out.append(template[pos:])
        return "".join(out)
    if slots and len(parts) == 1:
        a, b = slots[0] if len(slots) == 1 else max(slots, key=lambda s: s[1] - s[0])
        return template[:a] + str(parts[0]) + template[b:]
    raise FillError(
        f"「{template}」の空欄{len(slots)}か所に {value!r} を入れられません。"
        f"区切り付きの文字列（例 '100-0001'）か、{len(slots)}要素のリストで渡してください")


def _num(v):
    if isinstance(v, str) and re.fullmatch(r"[1-9]\d{0,14}|0", v):
        return int(v)
    return v


# --------------------------------------------------------------------------
def resolve(fields: list[Field], key: str) -> Field | None:
    for f in fields:
        if f.id == key:
            return f
    pool = fields
    if "!" in key:
        sname, rest = key.split("!", 1)
        sub = [f for f in fields if f.sheet == sname]
        if sub:
            pool, key = sub, rest
            for f in pool:
                if f.id in (key, f"{sname}!{key}"):
                    return f
    nk = normalize_label(key)
    for getter in (
        lambda f: f.id.split("!")[-1],
        lambda f: normalize_label(f.label),
        lambda f: f.id.split("!")[-1].split("/")[-1].split("#")[0],
    ):
        cand = [f for f in pool if getter(f) == nk]
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
    today: _dt.date | None = None,
) -> FillReport:
    """Fill *src* into *dst*. Pass *fields* (e.g. a corrected, saved field
    definition list) to skip automatic detection."""
    wb = ooxml.load(src)
    if fields is None:
        fields = detect(wb, sheet)
    report = FillReport()
    updates: dict[str, dict[tuple[int, int], object]] = {}
    controls: dict[str, tuple] = {}  # shape id -> (Control, checked)
    sheet_path = {s.name: s.path for s in wb.sheets}
    default_sheet = wb.sheet(sheet).name if sheet else form_sheets(wb)[0].name
    touched: set[str] = set()

    def put(sheet_name: str, ref: str, value, key: str, fid: str | None):
        try:
            r, c = parse_ref(ref)
        except ValueError as e:
            report.errors.append(f"{key}: {e}")
            return
        sh = wb.sheet(sheet_name)
        cell = sh.cells.get((r, c))
        if cell is not None and cell.formula is not None:
            report.errors.append(f"{key}: {sheet_name}!{ref} is a formula cell")
            return
        why = sh.lock_reason(r, c)
        if why:
            report.errors.append(f"{key}: {sheet_name}!{ref} is {why}; Excel would discard or reject it")
            return
        mg = sh.merge_of(r, c)
        if mg and (mg[0], mg[1]) != (r, c):
            report.errors.append(
                f"{key}: {sheet_name}!{ref} is hidden inside merged cells; write to "
                f"{ooxml.make_ref(mg[0], mg[1])} instead")
            return
        if isinstance(value, str):
            value = ooxml.clean_text(value)
        try:
            ooxml.check_value(value)
        except ooxml.WriteError as e:
            report.errors.append(f"{key}: {e}")
            return
        updates.setdefault(sheet_path[sheet_name], {})[(r, c)] = value
        report.written.append({"key": key, "field": fid, "cell": f"{sheet_name}!{ref}", "value": value})

    for key, value in data.items():
        if value is None or value == "" or value == [] or (
                isinstance(value, str) and value.strip().upper() in ("YYYY-MM-DD", "YYYY/MM/DD")):
            continue  # null / empty / template placeholder: leave the form untouched
        f = resolve(fields, key)
        if f is None:
            m = _CELL_KEY.match(key)
            if m and (m.group("sheet") is None or m.group("sheet") in sheet_path):
                sname = m.group("sheet") or default_sheet
                ref = m.group("ref").replace("$", "")
                if m.group("sheet") is None and _far_outside(wb.sheet(sname), ref):
                    # "TEL1" looks like a cell address but is almost certainly a field name
                    report.unknown_keys.append(key)
                    continue
                put(sname, ref, value, key, None)
            else:
                report.unknown_keys.append(key)
            continue
        touched.add(f.id)
        try:
            if f.kind == "table":
                _fill_table(f, value, put, key)
            elif f.kind in ("split", "digits"):
                _fill_parts(f, value, put, key)
            elif f.kind == "choice_cells":
                _fill_choice_cells(f, value, put, key, mark)
            elif f.kind == "checkbox":
                ctl = _control_of(wb, f)
                on = _truthy(value, f)
                controls[ctl.shape_id] = (ctl, on)
                report.written.append({"key": key, "field": f.id, "cell": f"{f.sheet}!{f.range}",
                                       "value": on, "control": ctl.shape_id})
            else:
                if (f.kind == "inline_date" and len(placeholder_dates(f.template or "")) > 1
                        and parse_date(value) is not None):
                    report.warnings.append(f"{key}: 期間の欄に日付が1つだけです（終了日は空欄のまま）")
                put(f.sheet, f.cell, _format_single(f, value, wb, mark, today), key, f.id)
        except FillError as e:
            report.errors.append(f"{key}: {e}")

    report.unfilled = [f.id for f in fields if f.id not in touched]
    if strict and report.unknown_keys:
        raise FillError(f"unknown keys: {report.unknown_keys}")
    if not report.errors:
        try:
            ooxml.save_with_updates(src, dst, updates, controls=list(controls.values()))
        except ooxml.WriteError as e:
            report.errors.append(str(e))
        else:
            report.verified = _verify(dst, report.written, report.errors)
            if not report.verified:
                try:
                    os.remove(dst)  # never leave a file we could not verify
                except OSError:
                    pass
    return report


def _far_outside(sheet, ref: str) -> bool:
    try:
        r, c = parse_ref(ref)
    except ValueError:
        return False  # put() reports the range error
    return c > max(sheet.max_col, 26) + 26 or r > max(sheet.max_row, 100) + 1000


def _format_single(f: Field, value, wb: Workbook, mark: str, today):
    if value is None:
        return None
    if f.kind == "inline_date":
        return format_date(value, f.template, _uses_wareki(wb, f.sheet), today)
    if f.kind == "date":
        return format_date(value, None, _uses_wareki(wb, f.sheet), today) if parse_date(value) else value
    if f.kind == "choice":
        return format_choice(f.template or "", value, mark)
    if f.kind == "circle":
        return format_circle(f.template or "", f.options, value)
    if f.kind == "inline_text":
        return format_inline(f.template or "", value)
    if f.kind == "pattern":
        return format_pattern(f.template or "", value)
    if f.kind == "text" and f.template and isinstance(value, (str, int)):
        v = str(value)
        return v if v.startswith(f.template) else f.template + v  # keep the printed 〒
    return _num(value) if isinstance(value, str) else value


_ERA_SPLIT = {"令和", "平成", "昭和", "西暦"}


def _fill_parts(f: Field, value, put, key: str) -> None:
    n = len(f.parts)
    if f.kind == "digits":
        chars = list(value) if isinstance(value, (list, tuple)) else list(re.sub(r"[-－ー‐−\s　]", "", str(value)))
        if len(chars) > n:
            raise FillError(f"{len(chars)} characters do not fit into {n} boxes")
        for p, ch in zip(f.parts, chars):
            put(f.sheet, p["cell"], _num(ch) if isinstance(ch, str) else ch, key, f.id)
        return
    if isinstance(value, dict):
        by_suffix = {normalize_label(p.get("suffix") or ""): p for p in f.parts}
        for k, v in value.items():
            p = by_suffix.get(normalize_label(k))
            if p is None:
                raise FillError(f"unknown part {k!r} (have {[p.get('suffix') for p in f.parts]})")
            put(f.sheet, p["cell"], _num(v) if isinstance(v, str) else v, key, f.id)
        return
    d = parse_date(value) if f.template in _ERA_SPLIT else None
    if d is not None:
        # 「令和[ ]年[ ]月[ ]日」 drawn over separate cells
        if f.template == "西暦":
            y = d.year
        else:
            era, y = to_wareki(d)
            if era != f.template:
                raise FillError(f"{value!r} は{era}の日付です（様式は「{f.template}」）")
        nums = {"年": y, "月": d.month, "日": d.day}
        for p in f.parts:
            if p.get("suffix") in nums:
                put(f.sheet, p["cell"], nums[p["suffix"]], key, f.id)
        return
    parts = _split_value(value, n)
    if len(parts) == 1 and n == 2 and re.fullmatch(r"\d{7}", str(parts[0])):
        parts = [parts[0][:3], parts[0][3:]]  # 郵便番号 1000001 -> 100 / 0001
    if len(parts) != n:
        raise FillError(f"expected {n} parts {[p.get('suffix') for p in f.parts]}, got {parts!r}")
    for p, v in zip(f.parts, parts):
        put(f.sheet, p["cell"], _num(v) if isinstance(v, str) else v, key, f.id)


def _fill_choice_cells(f: Field, value, put, key: str, mark: str) -> None:
    sel = [value] if isinstance(value, str) else list(value)
    sel_n = [normalize_label(s) for s in sel]
    opts = [normalize_label(p["option"]) for p in f.parts]
    missing = [s for s in sel_n if s not in opts]
    if missing:
        raise FillError(f"option(s) not in form: {missing}")
    for p, o in zip(f.parts, opts):
        put(f.sheet, p["cell"], mark if o in sel_n else "□", key, f.id)


def _control_of(wb: Workbook, f: Field):
    sid = f.parts[0].get("control") if f.parts else None
    for c in wb.sheet(f.sheet).controls:
        if c.shape_id == sid:
            return c
    raise FillError(f"check box {sid!r} not found in sheet {f.sheet!r}")


_YES = {"true", "yes", "on", "1", "はい", "有", "あり", "する", "✓", "✔", "レ", "■", "☑", "○", "〇", "済"}
_NO = {"false", "no", "off", "0", "いいえ", "無", "なし", "しない", "□", "☐"}


def _truthy(value, f: Field) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    t = normalize_label(str(value)).lower()
    if t in _YES or t == normalize_label(f.label).lower():
        return True
    if t in _NO:
        return False
    raise FillError(f"check box takes true / false (or 「{f.label}」), got {value!r}")


def _verify(dst: str, written: list[dict], errors: list[str]) -> bool:
    out = ooxml.load(dst)
    ok = True
    for w in written:
        sname, ref = w["cell"].rsplit("!", 1)
        if "control" in w:
            got = any(c.checked for c in out.sheet(sname).controls if c.shape_id == w["control"])
            if got != w["value"]:
                ok = False
                errors.append(f"verify failed at {w['cell']}: check box should be {w['value']}")
            continue
        got = out.sheet(sname).value(*parse_ref(ref))
        want = w["value"]
        if want == "":
            want = None
        if got != want and not (isinstance(want, float) and isinstance(got, (int, float)) and abs(got - want) < 1e-9):
            ok = False
            errors.append(f"verify failed at {w['cell']}: wrote {want!r}, read {got!r}")
    return ok


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
            if v is None or v == "":
                continue  # leave the cell as it is
            col = f.columns[ci]["col"]
            put(f.sheet, f"{col}{r}", _num(v) if isinstance(v, str) else v, f"{key}[{i}].{col_ids[ci]}", f.id)
