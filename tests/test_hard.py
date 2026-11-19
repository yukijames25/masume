"""Regression tests on tricky, realistic forms (see make_hard_fixtures.py)."""
import datetime as dt
import os
import shutil
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "src"))

from masume import detect, fill, load  # noqa: E402
from masume.fill import format_date, format_pattern  # noqa: E402

FX = os.path.join(HERE, "fixtures_hard")
TODAY = dt.date(2026, 11, 19)


def fx(name):
    return os.path.join(FX, name + ".xlsx")


def fmap(name):
    return {f.id: f for f in detect(load(fx(name)))}


class HardFixtureTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def run_fill(self, name, data):
        out = os.path.join(self.tmp, name + "_out.xlsx")
        rep = fill(fx(name), out, data, today=TODAY)
        self.assertTrue(rep.to_dict()["ok"], rep.to_dict())
        self.assertTrue(rep.verified)
        return load(out)

    # ------------------------------------------------------------------
    def test_rirekisho(self):
        f = fmap("rirekisho")
        for k in ("記入日", "氏名(ふりがな)", "氏名", "生年月日", "性別", "現住所(ふりがな)", "現住所(〒)",
                  "現住所", "電話", "E-mail", "学歴・職歴", "免許・資格"):
            self.assertIn(k, f)
        self.assertEqual(f["性別"].kind, "circle")
        self.assertEqual(f["学歴・職歴"].kind, "table")
        self.assertEqual([c["id"] for c in f["学歴・職歴"].columns], ["年", "月", "学歴・職歴"])
        self.assertEqual(len(f["学歴・職歴"].rows), 12)
        self.assertTrue(any(k.startswith("本人希望記入欄") for k in f))
        wb = self.run_fill("rirekisho", {
            "記入日": "2026-11-19", "氏名(ふりがな)": "やまだ たろう", "氏名": "山田 太郎",
            "生年月日": "1990-04-02", "性別": "男", "現住所(〒)": "123-4567", "現住所": "東京都千代田区1-1",
            "学歴・職歴": [{"年": 2009, "月": 4, "学歴・職歴": "〇〇大学 入学"}],
        })
        s = wb.sheet()
        self.assertEqual(s.value(3, 17), "令和8年11月19日現在")
        self.assertEqual(s.value(7, 6), "平成2年4月2日生（満36歳）")
        self.assertEqual(s.value(7, 23), "（男） ・ 女")
        self.assertEqual(s.value(9, 6), "〒123－4567")
        self.assertEqual(s.value(14, 7), "〇〇大学 入学")

    def test_juminhyo(self):
        f = fmap("juminhyo")
        self.assertEqual(f["住所(〒)"].kind, "pattern")
        self.assertEqual(f["請求者/住所"].cell, "I7")
        self.assertEqual(f["請求者/氏名"].section, "請求者")
        self.assertEqual(f["必要な方/氏名"].section, "必要な方")
        self.assertNotIn("あて先）〇〇市長", " ".join(f))
        wb = self.run_fill("juminhyo", {
            "請求者/生年月日": "1985-06-15", "必要な方/生年月日": "2020-01-10",
            "請求者との関係": {"その他": "祖父"}, "使用目的": "その他（パスポート申請）",
            "必要通数": 2,
        })
        s = wb.sheet()
        self.assertEqual(s.value(11, 23), "昭和60年6月15日")
        self.assertEqual(s.value(15, 9), "令和2年1月10日")
        self.assertEqual(s.value(16, 9), "□本人　□同一世帯の方　■その他（祖父）")
        self.assertEqual(s.value(19, 9), "□就職　□年金　□保険　■その他（パスポート申請）")

    def test_shinsei_kikan(self):
        f = fmap("shinsei_kikan")
        self.assertEqual(f["口座番号"].kind, "digits")
        self.assertEqual(len(f["口座番号"].parts), 7)
        self.assertEqual(f["出発時刻"].kind, "pattern")
        self.assertEqual(f["口座名義（カナ）"].kind, "text")
        wb = self.run_fill("shinsei_kikan", {
            "申請日": "2026-11-19", "電話番号": "03-1234-5678",
            "出張期間": "2026-12-01～2026-12-03", "出発時刻": "9:30", "帰着時刻": [18, "05"],
            "口座番号": "1234567", "支店": "渋谷", "口座種別": "普通", "概算金額": 52000,
        })
        s = wb.sheet()
        self.assertEqual(s.value(7, 8), "（03）1234－5678")
        self.assertEqual(s.value(8, 8), "令和8年12月1日　～　令和8年12月3日（3日間）")
        self.assertEqual(s.value(9, 8), "9時30分")
        self.assertEqual(s.value(9, 24), "18時05分")
        self.assertEqual([s.value(16, c) for c in range(8, 15)], [1, 2, 3, 4, 5, 6, 7])
        self.assertEqual(s.value(14, 20), "銀行・信用金庫渋谷支店")

    def test_ringi_stamps(self):
        f = fmap("ringi")
        self.assertEqual({k: f[k].cell for k in ("部長", "課長", "担当")},
                         {"部長": "X3", "課長": "AA3", "担当": "AD3"})
        self.assertEqual(f["部長"].section, "決裁")

    def test_multisheet_skips_example(self):
        f = fmap("multisheet")
        self.assertTrue(all(x.sheet == "本紙" for x in f.values()))
        wb = self.run_fill("multisheet", {"備考": "駐車場利用", "H6": "総務部", "本紙!来訪目的": "工事"})
        s = wb.sheet("本紙")
        self.assertEqual(s.value(9, 8), "駐車場利用")
        self.assertEqual(s.value(6, 8), "総務部", "bare cell refs go to the real form, not 記入例")
        self.assertIsNone(load(fx("multisheet")).sheet("本紙").value(6, 8))

    def test_kintou_split_and_choice_cells(self):
        f = fmap("kintou")
        self.assertEqual(f["氏名"].cell, "H6")
        self.assertEqual(f["年齢"].suffix, "歳")
        self.assertEqual(f["勤続年数"].kind, "split")
        self.assertEqual([p.get("suffix") for p in f["勤続年数"].parts], ["年", "ヶ月"])
        self.assertEqual(f["性別"].kind, "choice_cells")
        self.assertEqual(f["宛名"].cell, "B4")
        wb = self.run_fill("kintou", {
            "宛名": "株式会社〇〇", "氏名": "佐藤 花子", "年齢": 28, "勤続年数": "5年3ヶ月",
            "電話番号": "090-1234-5678", "性別": "女性", "郵便番号": "1000001",
        })
        s = wb.sheet()
        self.assertEqual((s.value(8, 8), s.value(8, 11)), (5, 3))
        self.assertEqual((s.value(9, 8), s.value(9, 13), s.value(9, 18)), ("090", 1234, 5678))
        self.assertEqual((s.value(10, 8), s.value(10, 12)), ("□", "■"))

    def test_underline_form(self):
        f = fmap("underline")
        self.assertEqual(f["住所"].range, "G9:AG9")
        self.assertNotIn("申込者情報", f)
        self.assertEqual(f["氏名"].section, "申込者情報")

    def test_refill_filled_form(self):
        out = os.path.join(self.tmp, "k.xlsx")
        rep = fill(fx("kintou"), out, {"氏名": "佐藤 花子"})
        self.assertTrue(rep.to_dict()["ok"])
        f = {x.id: x for x in detect(load(out))}
        self.assertEqual(f["氏名"].value, "佐藤 花子")
        self.assertEqual(f["氏名"].cell, "H6")


class FormatterTests(unittest.TestCase):
    def test_patterns(self):
        self.assertEqual(format_pattern("〒　　　－", "100-0001"), "〒100－0001")
        self.assertEqual(format_pattern("（　　　　）　　　　－", "03-1234-5678"), "（03）1234－5678")
        self.assertEqual(format_pattern("　　時　　分", "9:05"), "9時05分")

    def test_era_list(self):
        self.assertEqual(format_date("1990-04-02", "昭和・平成　　年　　月　　日生（満　　歳）", True, TODAY),
                         "平成2年4月2日生（満36歳）")
        self.assertEqual(format_date("1985-06-15", "明・大・昭・平・令　　年　　月　　日", True), "昭和60年6月15日")


if __name__ == "__main__":
    unittest.main()


class DogfoodRegressionTests(unittest.TestCase):
    """Problems found by an agent filling forms from plain-language requests."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.out = os.path.join(self.tmp, "o.xlsx")

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_wareki_inputs_are_parsed(self):
        from masume.wareki import parse_date
        self.assertEqual(parse_date("昭和60年6月15日"), dt.date(1985, 6, 15))
        self.assertEqual(parse_date("S60.6.15"), dt.date(1985, 6, 15))
        self.assertEqual(parse_date("平成元年1月8日"), dt.date(1989, 1, 8))
        self.assertEqual(parse_date("R8/11/19"), dt.date(2026, 11, 19))
        self.assertEqual(parse_date("２０２６／１１／１９"), dt.date(2026, 11, 19))

    def test_placeholders_and_nulls_never_overwrite(self):
        from masume.cli import template_for
        tpl = template_for(detect(load(fx("shinsei_kikan"))))
        self.assertTrue(all(v in (None, []) for v in tpl.values()), tpl)
        rep = fill(fx("shinsei_kikan"), self.out, dict(tpl, 申請日="YYYY-MM-DD", 所属=""))
        self.assertTrue(rep.to_dict()["ok"])
        self.assertEqual(rep.written, [])

    def test_unparseable_date_keeps_placeholder(self):
        rep = fill(fx("shinsei_kikan"), self.out, {"申請日": "来週"})
        self.assertTrue(rep.errors)
        d = rep.to_dict()
        self.assertFalse(d["saved"])
        self.assertIn("planned", d)
        self.assertFalse(os.path.exists(self.out))

    def test_range_formats_and_times(self):
        s = None
        rep = fill(fx("shinsei_kikan"), self.out,
                   {"出張期間": "2026-12-01/2026-12-03", "出発時刻": "09:30", "概算金額": "52000"}, today=TODAY)
        self.assertTrue(rep.to_dict()["ok"], rep.to_dict())
        s = load(self.out).sheet()
        self.assertEqual(s.value(8, 8), "令和8年12月1日　～　令和8年12月3日（3日間）")
        self.assertEqual(s.value(9, 8), "9時30分")
        self.assertEqual(s.value(11, 8), 52000)
        rep = fill(fx("shinsei_kikan"), self.out, {"出張期間": "2026-12-01"})
        self.assertTrue(rep.warnings)

    def test_readback_with_saved_fields(self):
        from masume.render import sheet_to_markdown
        fields = detect(load(fx("rirekisho")))
        rep = fill(fx("rirekisho"), self.out, {"現住所": "東京都千代田区1-1", "性別": "女",
                                                "学歴・職歴": [{"年": "2009", "月": "4", "学歴・職歴": "入学"}]})
        self.assertTrue(rep.to_dict()["ok"], rep.to_dict())
        self.assertEqual(load(self.out).sheet().value(14, 2), 2009, "numeric strings become numbers")
        md = sheet_to_markdown(load(self.out), None, False, fields)
        self.assertIn("| 現住所 | テキスト | F10:Y10 | [履歴書] 現在値: 東京都千代田区1-1 |", md)
        self.assertIn("1行記入済み", md)
        self.assertIn("現在値: 男 ・ （女）", md)


class RealFormLayoutTests(unittest.TestCase):
    """Layout habits seen in real municipal forms (made-up copy: make_frame_fixture.py)."""

    path = os.path.join(FX, "frame_todoke.xlsx")

    def test_fields(self):
        f = {x.id: x for x in detect(load(self.path))}
        self.assertEqual(set(f), {"名称（氏名）", "担当者名", "担当課", "所在地（住所）", "日付", "法人番号", "指定番号", "送付先", "変更年月日"})
        self.assertEqual(f["送付先"].template, "〒")  # printed mark, not a filled value
        self.assertEqual(len(f["指定番号"].parts), 6)  # boxes in the row under the label
        self.assertEqual(f["担当課"].cell, "U13")  # unmerged caption spills right: write under it
        self.assertEqual(f["所在地（住所）"].cell, "O5")      # open space inside the frame
        self.assertEqual(f["担当者名"].cell, "O12")
        self.assertEqual(len(f["法人番号"].parts), 13)
        self.assertEqual(f["日付"].template, "令和")

    def test_fill_era_split_date(self):
        tmp = tempfile.mkdtemp()
        try:
            out = os.path.join(tmp, "o.xlsx")
            rep = fill(self.path, out, {"日付": "2026-11-19", "法人番号": "1234567890123", "担当者名": "山田",
                                        "送付先": "730-0011 広島市中区"})
            self.assertEqual(rep.errors, [])
            sh = load(out).sheet("届出書")
            self.assertEqual(sh.value(17, 6), "〒730-0011 広島市中区")
            rep = fill(self.path, os.path.join(tmp, "q.xlsx"), {"変更年月日": "2026-12-01"})
            self.assertEqual(rep.errors, [])
            sq = load(os.path.join(tmp, "q.xlsx")).sheet("届出書")
            self.assertEqual([sq.value(21, c) for c in (18, 22, 25)], [8, 12, 1])  # 和暦: the sheet prints 令和
            self.assertEqual([sh.value(15, c) for c in (4, 7, 10)], [8, 11, 19])
            rep = fill(self.path, os.path.join(tmp, "p.xlsx"), {"日付": "2018-01-01"})
            self.assertTrue(rep.errors)  # 平成 date into a 令和 form is an error, not a wrong year
        finally:
            shutil.rmtree(tmp)


class FormControlCheckboxTests(unittest.TestCase):
    """Excel form-control check boxes (made by make_control_fixture.py, opens in Excel)."""

    path = os.path.join(FX, "checkbox_controls.xlsx")

    def test_detect(self):
        f = {x.id: x for x in detect(load(self.path))}
        self.assertEqual({k for k, v in f.items() if v.kind == "checkbox"},
                         {"名称変更/社名変更", "名称変更/法人成り", "所在地変更/事務所等の移転", "所在地変更/その他"})
        self.assertTrue(f["名称変更/法人成り"].value)
        self.assertIsNone(f["名称変更/社名変更"].value)
        self.assertIn("氏名", f)  # the captions did not turn into text fields

    def test_tick_and_untick(self):
        tmp = tempfile.mkdtemp()
        try:
            out = os.path.join(tmp, "o.xlsx")
            rep = fill(self.path, out, {"名称変更/社名変更": True, "名称変更/法人成り": False,
                                        "所在地変更/その他": "はい", "氏名": "山田"})
            self.assertEqual(rep.errors, [])
            self.assertTrue(rep.verified)
            state = {c.shape_id: c.checked for c in load(out).sheet("変更届").controls}
            self.assertEqual(state, {"1025": True, "1026": False, "1027": False, "1028": True})
            rep = fill(self.path, os.path.join(tmp, "p.xlsx"), {"名称変更/社名変更": "たぶん"})
            self.assertTrue(rep.errors)  # unclear answers are refused, not guessed
        finally:
            shutil.rmtree(tmp)


class ComparisonGridTests(unittest.TestCase):
    """「事項 | 変更前 | 変更後」 grid (made by make_henko_fixture.py)."""

    path = os.path.join(FX, "henko_todoke.xlsx")

    def test_ids_per_row_and_column(self):
        f = {x.id: x for x in detect(load(self.path))}
        self.assertEqual(len(f), 10)
        self.assertEqual(f["所在地（住所）(〒)/変更後"].cell, "S9")
        self.assertEqual(f["所在地（住所）/変更後"].cell, "R10")
        self.assertEqual(f["名称（氏名）(ふりがな)/変更前"].cell, "H6")
        self.assertEqual([p.get("suffix") for p in f["電話番号/変更前"].parts], ["）", "－", None])

    def test_fill(self):
        tmp = tempfile.mkdtemp()
        try:
            out = os.path.join(tmp, "o.xlsx")
            rep = fill(self.path, out, {"所在地（住所）/変更後": "広島市南区1-2-3", "所在地（住所）(〒)/変更後": "732-0001",
                                        "電話番号/変更後": "082-111-2222"})
            self.assertEqual(rep.errors, [])
            sh = load(out).sheet("変更届")
            self.assertEqual(sh.value(10, 18), "広島市南区1-2-3")
            self.assertEqual(sh.value(9, 19), "732-0001")
            self.assertEqual([sh.value(11, c) for c in (19, 22, 25)], ["082", 111, 2222])
        finally:
            shutil.rmtree(tmp)


class DrawnBoxTests(unittest.TestCase):
    """[ ]電子データ [ ]書面: check boxes drawn as boxed cells (make_henko_fixture.py)."""

    path = os.path.join(FX, "box_choice.xlsx")

    def test_groups_per_column(self):
        f = {x.id: x for x in detect(load(self.path))}
        self.assertEqual(set(f), {"受取方法/変更前", "受取方法/変更後", "通知方法/変更前", "通知方法/変更後"})
        self.assertEqual(f["受取方法/変更後"].options, ["電子データ", "書面"])

    def test_tick(self):
        tmp = tempfile.mkdtemp()
        try:
            out = os.path.join(tmp, "o.xlsx")
            rep = fill(self.path, out, {"受取方法/変更前": "書面", "受取方法/変更後": "電子データ"})
            self.assertEqual(rep.errors, [])
            sh = load(out).sheet("申出書")
            self.assertEqual([sh.value(4, 9), sh.value(4, 16), sh.value(4, 24), sh.value(4, 31)], [None, "✓", "✓", None])
            rep = fill(self.path, os.path.join(tmp, "p.xlsx"), {"受取方法/変更前": "郵送"})
            self.assertTrue(rep.errors)  # not an option on the form
        finally:
            shutil.rmtree(tmp)


class ReadBackTests(unittest.TestCase):
    """md --fields must report only what is really ticked (dogfood: a blank form showed every option)."""

    def test_blank_and_filled_boxes(self):
        from masume.render import current_value
        path = os.path.join(FX, "box_choice.xlsx")
        fields = detect(load(path))
        sheet = load(path).sheet("申出書")
        self.assertTrue(all(current_value(sheet, f) is None for f in fields))
        tmp = tempfile.mkdtemp()
        try:
            out = os.path.join(tmp, "o.xlsx")
            fill(path, out, {"受取方法/変更後": "電子データ"})
            sheet = load(out).sheet("申出書")
            got = {f.id: current_value(sheet, f) for f in fields}
            self.assertEqual(got["受取方法/変更後"], ["電子データ"])
            self.assertIsNone(got["受取方法/変更前"])
        finally:
            shutil.rmtree(tmp)


class PaperCheckWarningTests(unittest.TestCase):
    """Warnings for things a person should check on paper (dogfood findings)."""

    def test_short_digits_and_eraless_year(self):
        path = os.path.join(FX, "frame_todoke.xlsx")
        tmp = tempfile.mkdtemp()
        try:
            rep = fill(path, os.path.join(tmp, "o.xlsx"), {"指定番号": "12345", "変更年月日": "2026-12-01"})
            self.assertEqual(rep.errors, [])
            text = " ".join(rep.warnings)
            self.assertIn("6マスに5文字", text)
            self.assertIn("令和8年 → 8", text)
        finally:
            shutil.rmtree(tmp)


class EraYearMonthTests(unittest.TestCase):
    """dogfood: "2026-12" into 「令和[ ]年[ ]月」 became 令和2026年 with no warning."""

    path = os.path.join(FX, "frame_todoke.xlsx")

    def test_year_month_forms(self):
        tmp = tempfile.mkdtemp()
        try:
            for i, v in enumerate(["2026-12", "2026年12月", "令和8年12月", "R8.12"]):
                out = os.path.join(tmp, f"o{i}.xlsx")
                rep = fill(self.path, out, {"日付": v})
                self.assertEqual(rep.errors, [], v)
                sh = load(out).sheet("届出書")
                self.assertEqual([sh.value(15, c) for c in (4, 7, 10)], [8, 12, None], v)
            rep = fill(self.path, os.path.join(tmp, "x.xlsx"), {"日付": [2026, 12, 1]})
            self.assertTrue(rep.errors)  # a western year never goes into an era box
        finally:
            shutil.rmtree(tmp)
