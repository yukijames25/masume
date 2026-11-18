"""Generate realistic Excel 方眼紙 test forms. Requires openpyxl (dev only).

    python tests/make_fixtures.py
"""
from __future__ import annotations

import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

HERE = os.path.join(os.path.dirname(__file__), "fixtures")
thin = Side(style="thin")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)


def grid(ws, cols=34, width=2.6):
    for c in range(1, cols + 1):
        ws.column_dimensions[get_column_letter(c)].width = width


def put(ws, rng, value=None, border=True, bold=False, center=False):
    ws.merge_cells(rng) if ":" in rng else None
    first = rng.split(":")[0]
    if value is not None:
        ws[first] = value
    if bold:
        ws[first].font = Font(bold=True, size=14)
    if center:
        ws[first].alignment = Alignment(horizontal="center", vertical="center")
    if border:
        from openpyxl.utils.cell import range_boundaries

        c1, r1, c2, r2 = range_boundaries(rng if ":" in rng else f"{rng}:{rng}")
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.cell(r, c).border = BOX


def kyuka():
    wb = Workbook()
    ws = wb.active
    ws.title = "休暇届"
    grid(ws)
    put(ws, "B2:AG3", "休　暇　届", border=False, bold=True, center=True)
    put(ws, "T5:AG5", "令和　　年　　月　　日", border=False)
    put(ws, "B7:P7", "【申請者】", border=False)
    put(ws, "B8:G8", "所　属")
    put(ws, "H8:S8")
    put(ws, "T8:W8", "社員番号")
    put(ws, "X8:AG8")
    put(ws, "B9:G9", "氏　名")
    put(ws, "H9:AG9")
    put(ws, "B11:P11", "【休暇内容】", border=False)
    put(ws, "B12:G12", "休暇の種類")
    put(ws, "H12:AG12", "□年次有給休暇　□特別休暇　□代休　□その他")
    put(ws, "B13:G13", "開始日")
    put(ws, "H13:P13")
    put(ws, "Q13:T13", "終了日")
    put(ws, "U13:AG13")
    put(ws, "B14:G14", "日　数")
    put(ws, "H14:K14")
    put(ws, "L14:M14", "日", border=False)
    put(ws, "B15:AG15", "理由", border=True)
    put(ws, "B16:AG19")
    put(ws, "B21:AG21", "備考：＿＿＿＿＿＿＿＿＿＿", border=False)
    wb.save(os.path.join(HERE, "kyuka_todoke.xlsx"))


def keihi():
    wb = Workbook()
    ws = wb.active
    ws.title = "経費精算書"
    grid(ws)
    put(ws, "B2:AG3", "経費精算書", border=False, bold=True, center=True)
    put(ws, "B5:G5", "申請日")
    put(ws, "H5:Q5")
    put(ws, "B6:G6", "部署名")
    put(ws, "H6:Q6")
    put(ws, "R6:U6", "氏名")
    put(ws, "V6:AG6")
    put(ws, "B8:AG8", "明細", border=False)
    put(ws, "B9:C9", "No")
    put(ws, "D9:I9", "日付")
    put(ws, "J9:X9", "内容")
    put(ws, "Y9:AG9", "金額")
    for i in range(5):
        r = 10 + i
        put(ws, f"B{r}:C{r}", i + 1)
        put(ws, f"D{r}:I{r}")
        put(ws, f"J{r}:X{r}")
        put(ws, f"Y{r}:AG{r}")
    put(ws, "B15:X15", "合計")
    put(ws, "Y15:AG15", "=SUM(Y10:Y14)")
    put(ws, "B17:G17", "支払方法")
    put(ws, "H17:AG17", "□現金　□振込")
    wb.save(os.path.join(HERE, "keihi_seisan.xlsx"))


def plain():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "会員登録票"
    ws["A3"] = "氏名："
    ws["A4"] = "メール："
    ws["A5"] = "生年月日："
    ws["A6"] = "電話番号："
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 30
    wb.save(os.path.join(HERE, "plain_colon.xlsx"))


if __name__ == "__main__":
    os.makedirs(HERE, exist_ok=True)
    kyuka()
    keihi()
    plain()
    print("fixtures written to", HERE)
