"""Build fixtures_hard/frame_todoke.xlsx (requires openpyxl).

A made-up 届出書 that copies the *layout habits* found in real municipal tax
forms (2026-11-19 test on public forms): labels sitting inside one big ruled
frame with the value written in the open space beside them, a two-line label
「所在地」+「（住所）」, a merged label 「担当者名」 with open space beside it,
「令和[ ]年[ ]月[ ]日提出」 drawn over separate cells, one merged box per digit
for 法人番号, an office-only 受付印 box and a printed addressee 「〇〇市長」.

    python tests/make_frame_fixture.py
"""
from __future__ import annotations

import os

from openpyxl import Workbook
from openpyxl.styles import Border, Side
from openpyxl.utils import get_column_letter

HERE = os.path.join(os.path.dirname(__file__), "fixtures_hard")
T = Side(style="thin")


def side(ws, r, c, **kw):
    cell = ws.cell(r, c)
    b = cell.border
    cell.border = Border(left=kw.get("left", b.left), right=kw.get("right", b.right),
                         top=kw.get("top", b.top), bottom=kw.get("bottom", b.bottom))


def build(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "届出書"
    for c in range(1, 32):
        ws.column_dimensions[get_column_letter(c)].width = 2.6
    ws["E2"] = "特別徴収に関する届出書"
    # office-only stamp box
    ws.merge_cells("B4:F4"); ws["B4"] = "受付印"
    ws.merge_cells("B5:F8")
    for c in range(2, 7):
        side(ws, 4, c, top=T); side(ws, 8, c, bottom=T)
    side(ws, 5, 2, left=T); side(ws, 5, 6, right=T)
    # 「指定番号」 with its one-character boxes in the row under it
    ws.merge_cells("Y2:AB2"); ws["Y2"] = "指定番号"
    for c in range(25, 31):
        side(ws, 3, c, left=T, right=T, top=T, bottom=T)
    # printed addressee and a date split over cells
    ws["B10"] = "〇〇市長"
    for rng in ("B15:C15", "D15:E15", "G15:H15", "J15:K15", "M15:O15"):
        ws.merge_cells(rng)
    ws["B15"] = "令和"; ws["F15"] = "年"; ws["I15"] = "月"; ws["M15"] = "日提出"
    # an address box with 〒 printed in its corner
    ws.merge_cells("B17:E17"); ws["B17"] = "送付先"
    ws.merge_cells("F17:P18"); ws["F17"] = "〒"
    for c in range(6, 17):
        side(ws, 17, c, top=T); side(ws, 18, c, bottom=T)
    # 「変更年月日」 with [ ]年[ ]月[ ]日 in the row under it
    ws.merge_cells("R20:AA20"); ws["R20"] = "変更年月日"
    for rng in ("R21:T21", "V21:W21", "Y21:Z21"):
        ws.merge_cells(rng)
    ws["U21"] = "年"; ws["X21"] = "月"; ws["AA21"] = "日"
    # big frame J4:AD13, labels inside, values in the open space
    for c in range(10, 31):
        side(ws, 4, c, top=T); side(ws, 13, c, bottom=T)
    for r in range(4, 14):
        side(ws, r, 10, left=T); side(ws, r, 30, right=T)
    ws.merge_cells("K5:N5"); ws["K5"] = "所　在　地"
    ws.merge_cells("K6:N6"); ws["K6"] = "（　住　所　）"
    for c in range(10, 31):
        side(ws, 7, c, bottom=T)
    ws.merge_cells("K8:N8"); ws["K8"] = "名　　称"
    ws.merge_cells("K9:N9"); ws["K9"] = "（　氏　名　）"
    for c in range(10, 31):
        side(ws, 10, c, bottom=T)
    ws.merge_cells("K11:N11"); ws["K11"] = "法人番号"
    for c in range(15, 28):  # 13 one-character boxes O11:AA11
        side(ws, 11, c, left=T, right=T, top=T, bottom=T)
    ws.merge_cells("K12:N12"); ws["K12"] = "担当者名"
    ws["U12"] = "担当課"  # unmerged caption under a rule: its value goes under it (U13)
    wb.save(path)


if __name__ == "__main__":
    os.makedirs(HERE, exist_ok=True)
    build(os.path.join(HERE, "frame_todoke.xlsx"))
    print("wrote", os.path.join(HERE, "frame_todoke.xlsx"))
