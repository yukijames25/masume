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
        for k in ("記入日", "氏名(ふりがな)", "氏名", "生年月日", "性別", "現住所(ふりがな)", "現住所",
                  "現住所(2)", "電話", "E-mail", "学歴・職歴", "免許・資格"):
            self.assertIn(k, f)
        self.assertEqual(f["性別"].kind, "circle")
        self.assertEqual(f["学歴・職歴"].kind, "table")
        self.assertEqual([c["id"] for c in f["学歴・職歴"].columns], ["年", "月", "学歴・職歴"])
        self.assertEqual(len(f["学歴・職歴"].rows), 12)
        self.assertTrue(any(k.startswith("本人希望記入欄") for k in f))
        wb = self.run_fill("rirekisho", {
            "記入日": "2026-11-19", "氏名(ふりがな)": "やまだ たろう", "氏名": "山田 太郎",
            "生年月日": "1990-04-02", "性別": "男", "現住所": "123-4567", "現住所(2)": "東京都千代田区1-1",
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
        self.assertEqual(f["請求者/住所"].kind, "pattern")
        self.assertIn("住所(2)", f)
        self.assertEqual(f["請求者/氏名"].section, "請求者")
        self.assertEqual(f["必要な方/氏名"].section, "必要な方")
        self.assertNotIn("あて先）〇〇市長", " ".join(f))
        wb = self.run_fill("juminhyo", {
            "請求者/生年月日": "1985-06-15", "必要な方/生年月日": "2020-01-10",
            "請求者との関係": {"その他": "祖父"}, "使用目的": "その他（パスポート申請）",
            "必要通数": 2,
        })
        s = wb.sheet()
        self.assertEqual(s.value(11, 23), "昭60年6月15日")
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
        self.assertEqual(format_date("1985-06-15", "明・大・昭・平・令　　年　　月　　日", True), "昭60年6月15日")


if __name__ == "__main__":
    unittest.main()
