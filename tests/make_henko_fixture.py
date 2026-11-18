"""Build fixtures_hard/henko_todoke.xlsx (requires openpyxl).

A made-up 変更届 with the 「事項 | 変更前 | 変更後」 comparison grid seen in a
real municipal form: row labels on the left, two wide column headers, and in
each cell a box, a 〒 box over an address box, or a phone split 「（ ）－」.

    python tests/make_henko_fixture.py
"""
from __future__ import annotations

import os

from openpyxl import Workbook
from openpyxl.styles import Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries

HERE = os.path.join(os.path.dirname(__file__), "fixtures_hard")
T = Side(style="thin")


def put(ws, rng, value=None, box=True):
    if ":" in rng:
        ws.merge_cells(rng)
    c1, r1, c2, r2 = range_boundaries(rng if ":" in rng else f"{rng}:{rng}")
    if value is not None:
        ws.cell(r1, c1).value = value
    if box:
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.cell(r, c).border = Border(left=T if c == c1 else None, right=T if c == c2 else None,
                                              top=T if r == r1 else None, bottom=T if r == r2 else None)


def build(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "変更届"
    for c in range(1, 30):
        ws.column_dimensions[get_column_letter(c)].width = 2.6
    ws["B2"] = "名称・所在地等変更届出書"
    ws["B4"] = "２　変更事項"
    put(ws, "B5:G5", "事　　項")
    put(ws, "H5:Q5", "変　更　前")
    put(ws, "R5:AA5", "変　更　後")
    put(ws, "B6:G6", "フリガナ"); put(ws, "H6:Q6"); put(ws, "R6:AA6")
    put(ws, "B7:G8", "名称\n（氏　　名）"); put(ws, "H7:Q8"); put(ws, "R7:AA8")
    put(ws, "B9:G10", "所在地\n（住　　所）")
    for left in ("H", "R"):
        c = ws[f"{left}9"].column
        put(ws, f"{left}9", "〒", box=False)
        put(ws, f"{get_column_letter(c + 1)}9:{get_column_letter(c + 6)}9")
        put(ws, f"{left}10:{get_column_letter(c + 9)}10")
    put(ws, "B11:G11", "電話番号")
    for left in ("H", "R"):
        c = ws[f"{left}11"].column
        L = get_column_letter
        put(ws, f"{L(c)}11", "（", box=False)
        put(ws, f"{L(c + 1)}11:{L(c + 2)}11")
        put(ws, f"{L(c + 3)}11", "）", box=False)
        put(ws, f"{L(c + 4)}11:{L(c + 5)}11")
        put(ws, f"{L(c + 6)}11", "－", box=False)
        put(ws, f"{L(c + 7)}11:{L(c + 9)}11")
    wb.save(path)


if __name__ == "__main__":
    os.makedirs(HERE, exist_ok=True)
    out = os.path.join(HERE, "henko_todoke.xlsx")
    build(out)
    print("wrote", out)
