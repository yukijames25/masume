"""Render an Excel 方眼紙 sheet into compact, LLM-friendly Markdown."""
from __future__ import annotations

import re

from .detect import CHECKED_CHARS, Grid, detect_sheet
from .ooxml import Workbook

KIND_JA = {
    "text": "テキスト",
    "date": "日付",
    "inline_date": "日付(セル内)",
    "choice": "選択",
    "inline_text": "テキスト(セル内)",
    "table": "表",
    "circle": "選択(○で囲む)",
    "choice_cells": "選択(□別セル)",
    "pattern": "書式入り(〒・電話・時刻)",
    "split": "分割欄",
    "digits": "1文字1マス",
    "checkbox": "チェックボックス",
}


def _cell_md(text: str) -> str:
    return text.replace("|", "\\|").replace("\n", " ")


def current_value(sheet, f):
    """What is in the form now, for a (saved) field definition."""
    from .ooxml import parse_ref

    def at(ref):
        return sheet.value(*parse_ref(ref))

    if f.kind == "table":
        rows = []
        for r in f.rows:
            row = {c["id"]: at(f"{c['col']}{r}") for c in f.columns}
            if any(v not in (None, "") for v in row.values()):
                rows.append(row)
        return rows or None
    if f.kind == "digits":
        chars = [at(p["cell"]) for p in f.parts]
        return "".join("" if c is None else str(c) for c in chars) or None
    if f.kind == "split":
        vals = [at(p["cell"]) for p in f.parts]
        if all(v in (None, "") for v in vals):
            return None
        pre = f.template if f.template and f.template != "日付" else ""  # 令和8年12月, （082）251－1234
        return pre + "".join(f"{'' if v is None else v}{p.get('suffix') or ''}" for v, p in zip(vals, f.parts))
    if f.kind == "checkbox":
        sid = f.parts[0]["control"] if f.parts else None
        return True if any(c.shape_id == sid and c.checked for c in sheet.controls) else None
    if f.kind == "choice_cells":
        marks = CHECKED_CHARS + "レ○〇"
        if f.parts and f.parts[0].get("circle"):
            return [p["option"] for p in f.parts
                    if re.fullmatch(r"[（(○〇]\s*" + re.escape(p["option"]) + r"\s*[）)]?",
                                    str(at(p["cell"]) or "").strip())] or None
        sel = [p["option"] for p in f.parts if (str(at(p["cell"]) or "").strip() or "-") in marks]  # "" is in every string
        return sel or None
    v = at(f.cell)
    if f.kind == "choice" and isinstance(v, str) and f.options:
        # 「■年次有給休暇　□特別休暇…」 -> just what is ticked
        sel = [o for o in f.options if re.search("[" + CHECKED_CHARS + r"]\s*" + re.escape(o), v)]
        return sel or None
    if f.kind in ("inline_date", "pattern", "inline_text", "choice", "circle", "text") and f.template and v == f.template:
        return None  # still the printed placeholder
    return v


def _field_cells(fields) -> set:
    out = set()
    for f in fields:
        out.add(f.cell)
        out.update(p["cell"] for p in f.parts if "cell" in p)
    return out


def _context_after(g: Grid, f, skip: set = frozenset()) -> str:
    """The printed words right after the last box of a field (「月分（翌月10日納期限分）以降…」)."""
    from .ooxml import parse_ref
    r, c = parse_ref(f.parts[-1]["cell"])
    reg = g.at(r, c)
    c = (reg.c2 if reg is not None else c) + 1
    seen = 0
    while c <= g.max_col and seen < 4:
        reg = g.at(r, c)
        if reg is None:
            break
        if not reg.empty and isinstance(reg.value, str) and reg.ref not in skip:  # not another field's value
            t = _SPACE_RUN.sub("", reg.text)
            if len(t) > 1:
                return t[:30]
        seen += 0 if reg.empty else 1
        c = reg.c2 + 1
    return ""


_SPACE_RUN = re.compile(r"[\s\u3000]+")


def sheet_to_markdown(wb: Workbook, sheet_name: str | None = None, layout: bool = True,
                      fields: list | None = None) -> str:
    """With *fields* (saved from inspect --json before filling) ids stay stable and
    current values are shown, which is how a filled form should be checked."""
    sheet = wb.sheet(sheet_name)
    if fields is None:
        fields = detect_sheet(sheet, wb.styles)
    else:
        fields = [f for f in fields if f.sheet == sheet.name]
        for f in fields:
            f.value = current_value(sheet, f)
    lines = [f"# {sheet.name}", ""]

    lines += ["## 入力欄", "", "| id | 種類 | セル | 補足 |", "|---|---|---|---|"]
    grid = Grid(sheet, wb.styles)
    skip = _field_cells(fields)
    for f in fields:
        extra = ""
        if f.options:
            extra = "選択肢: " + " / ".join(f.options)
        elif f.columns:
            extra = "列: " + " / ".join(c["id"] for c in f.columns) + f"（最大{len(f.rows)}行）"
        elif f.parts:
            if any(p.get("suffix") for p in f.parts):
                # read like the form: 令和[E15]年[I15]月[M15]日, [M42]）[Y42]-[AQ42]
                pre = f.template if f.template and f.template != "日付" else ""
                extra = "書式: " + pre + "".join(f"[{p['cell']}]{p.get('suffix') or ''}" for p in f.parts)
            else:
                extra = "マス: " + " ".join(p["cell"] for p in f.parts)
        elif f.template:
            extra = "書式: " + f.template.replace("\n", " ")
        if f.suffix:
            extra = (extra + f" 単位: {f.suffix}").strip()
        if f.value not in (None, "", []):
            v = f.value
            if isinstance(v, list) and v and isinstance(v[0], dict):
                v = f"{len(v)}行記入済み"
            elif isinstance(v, list):
                v = "・".join(map(str, v))
            extra = (extra + f" 現在値: {v}").strip()
        if f.kind == "split" and f.template and f.parts and (f.template == "日付" or f.template in ("令和", "平成", "昭和", "西暦")):
            ctx = _context_after(grid, f, skip)
            if ctx:
                extra = (extra + f" 続き: {ctx}").strip()  # 「日付#2」 alone does not say which date this is
        if f.section:
            extra = (f"[{f.section}] " + extra).strip()
        kind = "選択(○で囲む)" if f.kind == "choice_cells" and f.parts and f.parts[0].get("circle") else KIND_JA.get(f.kind, f.kind)
        lines.append(f"| {_cell_md(f.id)} | {kind} | {f.range} | {_cell_md(extra)} |")
    if not fields:
        lines.append("| (検出なし) | | | |")
    lines.append("")

    if layout:
        g = Grid(sheet, wb.styles)
        box_refs = {f.cell for f in fields if f.kind in ("text", "date")}
        by_row: dict[int, list] = {}
        for reg in g.regions():
            by_row.setdefault(reg.r1, []).append(reg)
        lines += ["## レイアウト（論理行）", ""]
        prev = None
        repeat = 0
        for r in sorted(by_row):
            parts = []
            for reg in sorted(by_row[r], key=lambda x: x.c1):
                if reg.formula and not reg.text:
                    parts.append("（数式）")
                elif reg.empty:
                    parts.append(f"［{reg.ref}］" if reg.ref in box_refs else "［ ］")
                else:
                    parts.append(_cell_md(reg.text))
            line = " ｜ ".join(parts)
            sig = line.replace("［ ］", "").strip(" ｜")
            if prev is not None and sig == prev and sig == "":
                repeat += 1
                continue
            if repeat:
                lines.append(f"…（空行 ×{repeat}）")
                repeat = 0
            prev = sig
            lines.append(f"- R{r}: {line}")
        if repeat:
            lines.append(f"…（空行 ×{repeat}）")
        lines.append("")
    return "\n".join(lines)


def workbook_to_markdown(wb: Workbook, layout: bool = True, fields: list | None = None) -> str:
    if fields is not None:
        names = list(dict.fromkeys(f.sheet for f in fields))
        return "\n".join(sheet_to_markdown(wb, n, layout, fields) for n in names)
    return "\n".join(
        sheet_to_markdown(wb, s.name, layout) for s in wb.sheets if not s.hidden
    )
