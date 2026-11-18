"""Never produce a file Excel rejects, never write where Excel would drop the value.

Inputs in fixtures_review/ were saved by real Excel. Every case here was first
reproduced against Excel itself (corrupt-file prompt or silently lost value).
"""
import json
import os
import shutil
import sys
import tempfile
import unittest
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "src"))

from masume import fill, load, ooxml  # noqa: E402

FX = os.path.join(HERE, "fixtures_review")


def fx(name):
    return os.path.join(FX, name + ".xlsx")


class SafetyTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.out = os.path.join(self.tmp, "out.xlsx")

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def ok(self, name, data):
        rep = fill(fx(name), self.out, data)
        self.assertTrue(rep.to_dict()["ok"], rep.to_dict())
        self.assertTrue(rep.verified)
        with zipfile.ZipFile(self.out) as z:
            for n in z.namelist():
                if n.startswith("xl/worksheets/sheet"):
                    self.assertEqual(ooxml.check_sheet_xml(z.read(n).decode("utf-8")), [], n)
        return load(self.out).sheet()

    def refused(self, name, data, needle):
        rep = fill(fx(name), self.out, data)
        self.assertFalse(rep.to_dict()["ok"])
        self.assertTrue(any(needle in e for e in rep.errors), rep.errors)
        self.assertFalse(os.path.exists(self.out), "no file may be written when refusing")

    # corrupt-file cases ---------------------------------------------------
    def test_new_rows_stay_ascending(self):
        s = self.ok("base", {"A11": "x", "A12": "y", "A14": "z"})
        self.assertEqual((s.value(11, 1), s.value(12, 1), s.value(14, 1)), ("x", "y", "z"))
        self.ok("empty", {"A1": "a", "B3": "b", "C2": "c"})

    def test_out_of_range_refs(self):
        self.refused("base", {"A0": "x"}, "out of Excel's range")
        self.refused("base", {"A1048577": "x"}, "out of Excel's range")
        self.refused("base", {"XFE1": "x"}, "out of Excel's range")

    def test_non_finite_and_huge_numbers(self):
        self.refused("base", json.loads('{"B1": NaN}'), "cannot store")
        self.refused("base", {"B1": float("inf")}, "cannot store")
        self.refused("base", {"B1": 10 ** 400}, "too large")

    def test_text_limits_and_illegal_chars(self):
        self.refused("base", {"B1": "x" * 32768}, "32767")
        s = self.ok("base", {"B1": "a￾b￿c\ud800d"})
        self.assertEqual(s.value(1, 2), "abcd")

    def test_excel_table_header_is_protected(self):
        self.refused("table", {"A1": "氏名"}, "header row")
        self.ok("table", {"A3": "x", "A4": "y"})

    def test_rows_without_r_attribute_refused(self):
        self.refused("norowrefs", {"A3": "x"}, "without r=")

    def test_same_path_never_destroys_template(self):
        src = os.path.join(self.tmp, "same.xlsx")
        shutil.copy(fx("base"), src)
        with open(src, "rb") as fh:
            before = fh.read()
        rep = fill(src, src, {"B1": "x"})
        self.assertTrue(rep.errors)
        with open(src, "rb") as fh:
            self.assertEqual(fh.read(), before)

    # silent data-loss cases ------------------------------------------------
    def test_new_cells_inherit_column_style(self):
        base = load(fx("colstyle")).sheet()
        s = self.ok("colstyle", {"C5": "x", "E7": "y"})
        self.assertNotEqual(s.cells[(5, 3)].style, 0, "new cell must keep the sheet/column font")
        del base

    def test_x_escape_round_trip(self):
        s = self.ok("base", {"B1": "ID_x0031_23"})
        self.assertEqual(s.value(1, 2), "ID_x0031_23")
        with zipfile.ZipFile(self.out) as z:
            self.assertIn(b"ID_x005F_x0031_23", z.read("xl/worksheets/sheet1.xml"))

    def test_array_and_spill_ranges_refused(self):
        self.refused("formulas", {"D2": 999}, "array formula")
        self.refused("spill", {"A2": 99}, "array formula")

    def test_hidden_merged_cell_refused(self):
        p = os.path.join(HERE, "fixtures", "kyuka_todoke.xlsx")
        rep = fill(p, self.out, {"I9": "x"})
        self.assertTrue(any("merged" in e and "H9" in e for e in rep.errors), rep.errors)

    def test_field_like_key_is_not_a_cell(self):
        rep = fill(fx("base"), self.out, {"TEL1": "03-1234-5678"})
        self.assertEqual(rep.unknown_keys, ["TEL1"])
        self.assertEqual(rep.written, [])

    def test_prefixed_namespace_file(self):
        self.ok("prefixed", {"B2": "x", "C9": 1})


if __name__ == "__main__":
    unittest.main()
