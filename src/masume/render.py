"""Render an Excel 方眼紙 sheet into compact, LLM-friendly Markdown."""
from __future__ import annotations

from .detect import Grid, detect_sheet
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


def sheet_to_markdown(wb: Workbook, sheet_name: str | None = None, layout: bool = True) -> str:
    sheet = wb.sheet(sheet_name)
    fields = detect_sheet(sheet, wb.styles)
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
        if f.value not in (None, ""):
            extra = (extra + f" 現在値: {f.value}").strip()
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


def workbook_to_markdown(wb: Workbook, layout: bool = True) -> str:
    return "\n".join(
        sheet_to_markdown(wb, s.name, layout) for s in wb.sheets if not s.hidden
    )
