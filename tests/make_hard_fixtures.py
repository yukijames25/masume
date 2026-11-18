"""Generate *tricky* but realistic Japanese Excel 方眼紙 forms. Requires openpyxl.

    python tests/make_hard_fixtures.py

Each builder also returns EXPECTED: what a human would call the field list
(label -> top-left cell of the input area). See fixtures_hard/REPORT.md.
"""
from __future__ import annotations

import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries

HERE = os.path.join(os.path.dirname(__file__), "fixtures_hard")
thin = Side(style="thin")
dotted = Side(style="dotted")


def grid(ws, cols=34, width=2.6):
    for c in range(1, cols + 1):
        ws.column_dimensions[get_column_letter(c)].width = width


def put(ws, rng, value=None, border="box", bold=False, center=False, vertical=False, size=None):
    """Write like Excel does: merged range, border drawn on the outline only.

    border: "box" (outline), "bottom" (underline only), None.
    """
    if ":" in rng:
        ws.merge_cells(rng)
    c1, r1, c2, r2 = range_boundaries(rng if ":" in rng else f"{rng}:{rng}")
    if value is not None:
        ws.cell(r1, c1).value = value
    top = ws.cell(r1, c1)
    if bold or size:
        top.font = Font(bold=bold, size=size or 11)
    if center or vertical:
        top.alignment = Alignment(horizontal="center", vertical="center",
                                  wrap_text=vertical, text_rotation=255 if vertical else 0)
    if border is None:
        return
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            if border == "bottom":
                if r == r2:
                    ws.cell(r, c).border = Border(bottom=thin)
                continue
            ws.cell(r, c).border = Border(
                left=thin if c == c1 else None,
                right=thin if c == c2 else None,
                top=thin if r == r1 else None,
                bottom=thin if r == r2 else None,
            )


def save(wb, name):
    wb.save(os.path.join(HERE, name))


# ---------------------------------------------------------------------------
def rirekisho():
    """JIS-style 履歴書."""
    wb = Workbook()
    ws = wb.active
    ws.title = "履歴書"
    grid(ws)
    put(ws, "B2:P3", "履　歴　書", border=None, bold=True, size=16)
    put(ws, "Q3:Y3", "令和　　年　　月　　日現在", border=None)
    put(ws, "B4:E4", "ふりがな")
    put(ws, "F4:Y4")
    put(ws, "B5:E6", "氏　　名")
    put(ws, "F5:Y6")
    put(ws, "Z4:AG10", "写真をはる位置", center=True)
    put(ws, "B7:E7", "生年月日")
    put(ws, "F7:T7", "昭和・平成　　年　　月　　日生（満　　歳）")
    put(ws, "U7:V7", "性別")
    put(ws, "W7:Y7", "男 ・ 女", center=True)
    put(ws, "B8:E8", "ふりがな")
    put(ws, "F8:Y8")
    put(ws, "B9:E10", "現住所")
    put(ws, "F9:Y9", "〒　　　－")
    put(ws, "F10:Y10")
    put(ws, "B11:E11", "電話")
    put(ws, "F11:N11")
    put(ws, "O11:Q11", "E-mail")
    put(ws, "R11:AG11")
    put(ws, "B13:D13", "年", center=True)
    put(ws, "E13:F13", "月", center=True)
    put(ws, "G13:AG13", "学歴・職歴", center=True)
    for r in range(14, 26):
        put(ws, f"B{r}:D{r}")
        put(ws, f"E{r}:F{r}")
        put(ws, f"G{r}:AG{r}")
    put(ws, "B27:D27", "年", center=True)
    put(ws, "E27:F27", "月", center=True)
    put(ws, "G27:AG27", "免許・資格", center=True)
    for r in range(28, 32):
        put(ws, f"B{r}:D{r}")
        put(ws, f"E{r}:F{r}")
        put(ws, f"G{r}:AG{r}")
    put(ws, "B33:AG33", "志望の動機、特技、好きな学科、アピールポイントなど")
    put(ws, "B34:AG38")
    put(ws, "B40:AG40", "本人希望記入欄（特に給料・職種・勤務時間・勤務地・その他についての希望などがあれば記入）")
    put(ws, "B41:AG44")
    save(wb, "rirekisho.xlsx")


def juminhyo():
    """住民票の写し等交付請求書: vertical group labels, ※ markers, 〒, era lists."""
    wb = Workbook()
    ws = wb.active
    ws.title = "請求書"
    grid(ws)
    put(ws, "B2:AG3", "住民票の写し等交付請求書", border=None, bold=True, center=True, size=14)
    put(ws, "T4:AG4", "令和　　年　　月　　日", border=None)
    put(ws, "B5:AG5", "（あて先）〇〇市長", border=None)
    # 請求者 block: vertical label spanning 6 rows
    put(ws, "B6:C11", "請求者", vertical=True)
    put(ws, "D6:H7", "※住　所")
    put(ws, "I6:AG6", "〒　　　－")
    put(ws, "I7:AG7")
    put(ws, "D8:H8", "フリガナ")
    put(ws, "I8:AG8")
    put(ws, "D9:H10", "※氏　名")
    put(ws, "I9:AG10")
    put(ws, "D11:H11", "電話番号")
    put(ws, "I11:R11")
    put(ws, "S11:V11", "生年月日")
    put(ws, "W11:AG11", "明・大・昭・平・令　　年　　月　　日")
    # 必要な方 block
    put(ws, "B13:C17", "必要な方", vertical=True)
    put(ws, "D13:H13", "※住　所")
    put(ws, "I13:AG13")
    put(ws, "D14:H14", "※氏　名")
    put(ws, "I14:AG14")
    put(ws, "D15:H15", "生年月日")
    put(ws, "I15:AG15", "昭和・平成・令和　　年　　月　　日")
    put(ws, "D16:H16", "請求者との関係")
    put(ws, "I16:AG16", "□本人　□同一世帯の方　□その他（　　　　　　　）")
    put(ws, "D17:H17", "※必要通数")
    put(ws, "I17:L17")
    put(ws, "M17:N17", "通", border=None)
    put(ws, "B19:H19", "※使用目的")
    put(ws, "I19:AG19", "□就職　□年金　□保険　□その他（　　　　　　　　）")
    put(ws, "B21:AG21", "※印の欄は必ず記入してください。", border=None)
    save(wb, "juminhyo.xlsx")


def shinsei_kikan():
    """休職・出張申請 with a period, times, split phone, per-digit 口座番号."""
    wb = Workbook()
    ws = wb.active
    ws.title = "申請書"
    grid(ws)
    put(ws, "B2:AG3", "出張申請書", border=None, bold=True, center=True, size=14)
    put(ws, "T4:AG4", "申請日：令和　　年　　月　　日", border=None)
    put(ws, "B6:G6", "所属")
    put(ws, "H6:S6")
    put(ws, "T6:W6", "氏名")
    put(ws, "X6:AE6")
    put(ws, "AF6:AG6", "印", center=True)
    put(ws, "B7:G7", "電話番号")
    put(ws, "H7:S7", "（　　　　）　　　　－")
    put(ws, "B8:G8", "出張期間")
    put(ws, "H8:AG8", "令和　　年　　月　　日　～　令和　　年　　月　　日（　　日間）")
    put(ws, "B9:G9", "出発時刻")
    put(ws, "H9:S9", "　　時　　分")
    put(ws, "T9:W9", "帰着時刻")
    put(ws, "X9:AG9", "　　時　　分")
    put(ws, "B10:G10", "出張先")
    put(ws, "H10:AG10")
    put(ws, "B11:G11", "概算金額")
    put(ws, "H11:Q11")
    put(ws, "R11", "円", border=None)
    put(ws, "B13:AG13", "【振込先口座】", border=None)
    put(ws, "B14:G14", "金融機関名")
    put(ws, "H14:S14")
    put(ws, "T14:AG14", "銀行・信用金庫　　　　　支店")
    put(ws, "B15:G15", "口座種別")
    put(ws, "H15:AG15", "□普通　□当座")
    put(ws, "B16:G16", "口座番号")
    for c in range(8, 15):  # H..N: one digit per cell
        put(ws, get_column_letter(c) + "16")
    put(ws, "B17:G17", "口座名義（カナ）")
    put(ws, "H17:AG17")
    save(wb, "shinsei_kikan.xlsx")


def ringi():
    """稟議書 with 決裁 stamp boxes."""
    wb = Workbook()
    ws = wb.active
    ws.title = "稟議書"
    grid(ws)
    put(ws, "B2:P3", "稟　議　書", border=None, bold=True, size=16)
    put(ws, "W2:W5", "決裁", vertical=True)
    put(ws, "X2:Z2", "部長", center=True)
    put(ws, "AA2:AC2", "課長", center=True)
    put(ws, "AD2:AF2", "担当", center=True)
    put(ws, "X3:Z5")
    put(ws, "AA3:AC5")
    put(ws, "AD3:AF5")
    put(ws, "B7:F7", "稟議番号")
    put(ws, "G7:N7")
    put(ws, "O7:S7", "起案日")
    put(ws, "T7:AF7", "令和　　年　　月　　日")
    put(ws, "B8:F8", "起案部署")
    put(ws, "G8:N8")
    put(ws, "O8:S8", "起案者")
    put(ws, "T8:AF8")
    put(ws, "B9:F9", "件　　名")
    put(ws, "G9:AF9")
    put(ws, "B10:F10", "金　　額")
    put(ws, "G10:R10")
    put(ws, "S10:T10", "円", border=None)
    put(ws, "B11:F11", "支払先")
    put(ws, "G11:AF11")
    put(ws, "B12:AF12", "内容・目的")
    put(ws, "B13:AF18")
    put(ws, "B19:F19", "添付書類")
    put(ws, "G19:AF19", "□見積書　□カタログ　□契約書案　□その他")
    save(wb, "ringi.xlsx")


def multisheet():
    """Hidden list sheet + 記入例 (filled example) + 本紙 (blank)."""
    wb = Workbook()
    lst = wb.active
    lst.title = "リスト"
    for i, v in enumerate(["営業部", "総務部", "開発部"], 1):
        lst.cell(i, 1).value = v
    lst.sheet_state = "hidden"

    def body(ws, example: bool):
        grid(ws)
        put(ws, "B2:AG3", "入館申請書" + ("（記入例）" if example else ""), border=None, bold=True, center=True)
        put(ws, "T4:AG4", "令和８年１１月１９日" if example else "令和　　年　　月　　日", border=None)
        put(ws, "B6:G6", "部署")
        put(ws, "H6:S6", "営業部" if example else None)
        put(ws, "T6:W6", "氏名")
        put(ws, "X6:AG6", "山田　太郎" if example else None)
        put(ws, "B7:G7", "入館日")
        put(ws, "H7:S7", "2026/12/1" if example else None)
        put(ws, "T7:W7", "人数")
        put(ws, "X7:AA7", 3 if example else None)
        put(ws, "AB7:AC7", "名", border=None)
        put(ws, "B8:G8", "来訪目的")
        put(ws, "H8:AG8", "■打合せ　□工事　□その他" if example else "□打合せ　□工事　□その他")
        put(ws, "B9:G9", "備考")
        put(ws, "H9:AG9")

    body(wb.create_sheet("記入例"), True)
    body(wb.create_sheet("本紙"), False)
    save(wb, "multisheet.xlsx")


def kintou():
    """均等割付 labels, units/labels *right* of the box, split boxes, 御中."""
    wb = Workbook()
    ws = wb.active
    ws.title = "申込書"
    grid(ws)
    put(ws, "B2:AG3", "研修参加申込書", border=None, bold=True, center=True)
    put(ws, "B4:N4")
    put(ws, "O4:P4", "御中", border=None)
    put(ws, "B6:G6", "氏　　　名")
    put(ws, "H6:AG6")
    put(ws, "B7:G7", "年　　齢")
    put(ws, "H7:J7")
    put(ws, "K7:L7", "歳", border=None)
    put(ws, "B8:G8", "勤続年数")
    put(ws, "H8:I8")
    put(ws, "J8", "年", border=None)
    put(ws, "K8:L8")
    put(ws, "M8:N8", "ヶ月", border=None)
    put(ws, "B9:G9", "電話番号")
    put(ws, "H9:K9")
    put(ws, "L9", "－", border=None)
    put(ws, "M9:P9")
    put(ws, "Q9", "－", border=None)
    put(ws, "R9:U9")
    put(ws, "B10:G10", "性　　別")
    put(ws, "H10", "□", border=None)
    put(ws, "I10:K10", "男性", border=None)
    put(ws, "L10", "□", border=None)
    put(ws, "M10:O10", "女性", border=None)
    put(ws, "B11:G11", "郵便番号")
    put(ws, "H11", "〒", border=None)
    put(ws, "I11:K11")
    put(ws, "L11", "-", border=None)
    put(ws, "M11:P11")
    save(wb, "kintou.xlsx")


def underline():
    """Unbordered 方眼紙: input areas are just underlined, headings underlined too."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    grid(ws)
    put(ws, "B2:AG2", "入会申込書", border="bottom", bold=True, center=True, size=14)
    put(ws, "B4:H4", "■ 申込者情報", border=None)
    for c in range(9, 34):  # decorative rule after the heading (unmerged cells)
        ws.cell(4, c).border = Border(bottom=thin)
    put(ws, "B6:F6", "フリガナ", border=None)
    put(ws, "G6:AG6", border="bottom")
    put(ws, "B7:F7", "氏名", border=None)
    put(ws, "G7:AG7", border="bottom")
    put(ws, "B8:F8", "生年月日", border=None)
    put(ws, "G8:R8", "　　　　年　　月　　日", border="bottom")
    put(ws, "B9:F9", "住所", border=None)
    # unmerged underline: each cell has its own bottom border
    for c in range(7, 34):
        ws.cell(9, c).border = Border(bottom=thin)
    put(ws, "B10:F10", "メール", border=None)
    put(ws, "G10:AG10", border="bottom")
    put(ws, "B12:AG12", "上記の通り申し込みます。", border=None)
    put(ws, "R14:U14", "署名", border=None)
    put(ws, "V14:AG14", border="bottom")
    save(wb, "underline.xlsx")


if __name__ == "__main__":
    os.makedirs(HERE, exist_ok=True)
    for fn in (rirekisho, juminhyo, shinsei_kikan, ringi, multisheet, kintou, underline):
        fn()
    print("hard fixtures written to", HERE)
