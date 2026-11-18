"""Render an Excel 方眼紙 sheet into compact, LLM-friendly Markdown."""
from __future__ import annotations

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
        return "".join(f"{'' if v is None else v}{p.get('suffix') or ''}" for v, p in zip(vals, f.parts))
    if f.kind == "choice_cells":
        sel = [p["option"] for p in f.parts if str(at(p["cell"]) or "") in CHECKED_CHARS]
        return sel or None
    v = at(f.cell)
    if f.kind in ("inline_date", "pattern", "inline_text", "choice", "circle") and v == f.template:
        return None  # still the printed placeholder
    return v


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
    for f in fields:
        extra = ""
        if f.options:
            extra = "選択肢: " + " / ".join(f.options)
        elif f.columns:
            extra = "列: " + " / ".join(c["id"] for c in f.columns) + f"（最大{len(f.rows)}行）"
        elif f.parts:
            extra = "マス: " + " ".join(p["cell"] + (p.get("suffix") or "") for p in f.parts)
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
        if f.section:
            extra = (f"[{f.section}] " + extra).strip()
        lines.append(f"| {_cell_md(f.id)} | {KIND_JA.get(f.kind, f.kind)} | {f.range} | {_cell_md(extra)} |")
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
