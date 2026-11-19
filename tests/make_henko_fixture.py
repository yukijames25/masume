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


def build_boxes(path):
    """「受取方法  [ ]電子データ [ ]書面 | [ ]電子データ [ ]書面」: check boxes drawn as boxed cells."""
    wb = Workbook()
    ws = wb.active
    ws.title = "申出書"
    for c in range(1, 40):
        ws.column_dimensions[get_column_letter(c)].width = 2.6
    put(ws, "B2:G2", "事　　項")
    put(ws, "H2:V2", "変更前（旧）")
    put(ws, "W2:AK2", "変更後（新）")
    for r, label in ((3, "受取方法"), (6, "通知方法")):
        put(ws, f"B{r}:G{r + 2}", label)
        for left in (8, 23):  # H, W
            L = get_column_letter
            put(ws, f"{L(left + 1)}{r + 1}:{L(left + 2)}{r + 1}")
            put(ws, f"{L(left + 3)}{r}:{L(left + 7)}{r + 2}", "電子データ", box=False)
            put(ws, f"{L(left + 8)}{r + 1}:{L(left + 9)}{r + 1}")
            put(ws, f"{L(left + 10)}{r}:{L(left + 13)}{r + 2}", "書面", box=False)
    wb.save(path)


def build_misc(path):
    """「有 （[ ]年[ ]月[ ]日） ・ 無」 in separate cells, and 「電話」 with 「（　）　－」 printed under it."""
    wb = Workbook()
    ws = wb.active
    ws.title = "申請書"
    for c in range(1, 30):
        ws.column_dimensions[get_column_letter(c)].width = 2.6
    ws["B1"] = "令和　年　月　日"
    put(ws, "B2:H3", "取り消されたことの有無")
    for ref, v in (("J2", "有"), ("K2", "（"), ("N2", "年"), ("Q2", "月"), ("T2", "日"), ("U2", "）"), ("V2", "・"), ("W2", "無")):
        ws[ref] = v
    for rng in ("L2:M2", "O2:P2", "R2:S2"):
        ws.merge_cells(rng)
    put(ws, "B6:D6", "電話", box=False)
    for ref, v in (("E7", "（"), ("G7", "）"), ("I7", "－")):
        ws[ref] = v
    wb.save(path)


def build_meisai(path):
    """6か月の明細表: 「年　月」 ([ ]・[ ]) + 人/円 twice, told apart by the captions above."""
    wb = Workbook()
    ws = wb.active
    ws.title = "明細"
    for c in range(1, 20):
        ws.column_dimensions[get_column_letter(c)].width = 2.6
    put(ws, "B2:D3", "月区分")
    put(ws, "F2:I2", "支給人員・支給額", box=False)
    put(ws, "F3:I3", "（臨時雇用者を除く。）", box=False)
    put(ws, "L2:O2", "臨時雇用者の", box=False)
    put(ws, "L3:O3", "支給人員・支給額", box=False)
    for c in range(5, 17):
        ws.cell(2, c).border = Border(top=T)
    put(ws, "B4:D4", "年　月")
    put(ws, "E4:F4", "人"); put(ws, "G4:J4", "円"); put(ws, "K4:L4", "人"); put(ws, "M4:P4", "円")
    for r in range(5, 9):
        put(ws, f"B{r}"); ws[f"C{r}"] = "・"; put(ws, f"D{r}")
        for rng in ("E{r}:F{r}", "G{r}:J{r}", "K{r}:L{r}", "M{r}:P{r}"):
            put(ws, rng.format(r=r))
    wb.save(path)


if __name__ == "__main__":
    os.makedirs(HERE, exist_ok=True)
    out = os.path.join(HERE, "henko_todoke.xlsx")
    build(out)
    print("wrote", out)
    out = os.path.join(HERE, "box_choice.xlsx")
    build_boxes(out)
    print("wrote", out)
    out = os.path.join(HERE, "uwamu_tel.xlsx")
    build_misc(out)
    print("wrote", out)
    out = os.path.join(HERE, "meisai.xlsx")
    build_meisai(out)
    print("wrote", out)
