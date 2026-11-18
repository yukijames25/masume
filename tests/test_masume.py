import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
from xml.etree import ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "src"))

from masume import detect, fill, load  # noqa: E402
from masume import ooxml  # noqa: E402
from masume.fill import format_choice, format_date, FillError  # noqa: E402
from masume.render import sheet_to_markdown  # noqa: E402

FX = os.path.join(HERE, "fixtures")
KYUKA = os.path.join(FX, "kyuka_todoke.xlsx")
KEIHI = os.path.join(FX, "keihi_seisan.xlsx")
PLAIN = os.path.join(FX, "plain_colon.xlsx")
EXCEL_SAVED = os.path.join(FX, "excel_saved_kyuka.xlsx")  # same form, re-saved by real Excel (sharedStrings, x14ac, mc:Ignorable)


def fields_of(path):
    return {f.id: f for f in detect(load(path))}


class DetectTests(unittest.TestCase):
    def test_kyuka_fields(self):
        f = fields_of(KYUKA)
        expect = {
            "日付": ("T5", "inline_date"),
            "所属": ("H8", "text"),
            "社員番号": ("X8", "text"),
            "氏名": ("H9", "text"),
            "休暇の種類": ("H12", "choice"),
            "開始日": ("H13", "date"),
            "終了日": ("U13", "date"),
            "日数": ("H14", "text"),
            "理由": ("B16", "text"),
            "備考": ("B21", "inline_text"),
        }
        self.assertEqual(set(f), set(expect))
        for k, (cell, kind) in expect.items():
            self.assertEqual((f[k].cell, f[k].kind), (cell, kind), k)
        self.assertEqual(f["休暇の種類"].options, ["年次有給休暇", "特別休暇", "代休", "その他"])
        self.assertEqual(f["氏名"].section, "申請者")
        self.assertEqual(f["理由"].range, "B16:AG19")

    def test_keihi_table_and_formula(self):
        f = fields_of(KEIHI)
        self.assertIn("明細", f)
        t = f["明細"]
        self.assertEqual(t.kind, "table")
        self.assertEqual([c["id"] for c in t.columns], ["No", "日付", "内容", "金額"])
        self.assertEqual(t.rows, [10, 11, 12, 13, 14])
        self.assertNotIn("合計", f, "formula cell must not become an input box")
        self.assertEqual(f["支払方法"].options, ["現金", "振込"])

    def test_plain_colon_labels(self):
        f = fields_of(PLAIN)
        self.assertEqual({k: v.cell for k, v in f.items()},
                         {"氏名": "B3", "メール": "B4", "生年月日": "B5", "電話番号": "B6"})
        self.assertEqual(f["生年月日"].kind, "date")

    def test_excel_saved_file_same_fields(self):
        a = {k: (v.cell, v.kind) for k, v in fields_of(KYUKA).items()}
        b = {k: (v.cell, v.kind) for k, v in fields_of(EXCEL_SAVED).items()}
        self.assertEqual(a, b)

    def test_markdown(self):
        md = sheet_to_markdown(load(KEIHI))
        self.assertIn("| 明細 | 表 |", md)
        self.assertIn("（数式）", md)


class FormatTests(unittest.TestCase):
    def test_wareki(self):
        self.assertEqual(format_date("2019-05-01", "令和　　年　　月　　日", True), "令和元年5月1日")
        self.assertEqual(format_date("2019-04-30", "令和　　年　　月　　日", True), "平成31年4月30日")
        self.assertEqual(format_date("2026/11/19", "　　年　　月　　日", False), "2026年11月19日")
        self.assertEqual(format_date("2026-11-19", "西暦　　年　　月　　日", True), "西暦2026年11月19日")
        self.assertEqual(format_date("2026-11-19", "提出日：令和　年　月　日", True), "提出日：令和8年11月19日")
        self.assertEqual(format_date("未定", None, True), "未定")

    def test_choice(self):
        self.assertEqual(format_choice("□有　□無", "無"), "□有　■無")
        self.assertEqual(format_choice("■有　□無", ["無"], "☑"), "□有　☑無")
        with self.assertRaises(FillError):
            format_choice("□有　□無", "多分")


class XmlSurgeryTests(unittest.TestCase):
    def test_prefixed_namespace_and_insert(self):
        xml = ('<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
               '<x:sheetData><x:row r="1" spans="1:1"><x:c r="A1" s="3"/></x:row>'
               '<x:row r="5"><x:c r="C5"><x:v>1</x:v></x:c></x:row></x:sheetData></x:worksheet>')
        out = ooxml.set_cells_xml(xml, {(1, 1): "あ<&>", (1, 3): 7, (3, 2): "new", (5, 1): "z"})
        root = ET.fromstring(out)  # must stay well-formed
        ns = {"x": ooxml.NS}
        rows = [r.get("r") for r in root.findall(".//x:row", ns)]
        self.assertEqual(rows, ["1", "3", "5"])
        r1 = [c.get("r") for c in root.findall(".//x:row[@r='1']/x:c", ns)]
        self.assertEqual(r1, ["A1", "C1"])
        r5 = [c.get("r") for c in root.findall(".//x:row[@r='5']/x:c", ns)]
        self.assertEqual(r5, ["A5", "C5"])
        a1 = root.find(".//x:c[@r='A1']", ns)
        self.assertEqual(a1.get("s"), "3")
        self.assertEqual(a1.find("x:is/x:t", ns).text, "あ<&>")
        self.assertNotIn("spans=", out)

    def test_empty_sheetdata(self):
        xml = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>'
        out = ooxml.set_cells_xml(xml, {(2, 2): "x"})
        self.assertIn('<row r="2"><c r="B2" t="inlineStr">', out)

    def test_formula_refused(self):
        xml = ('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
               '<row r="1"><c r="A1"><f>1+1</f><v>2</v></c></row></sheetData></worksheet>')
        with self.assertRaises(ooxml.WriteError):
            ooxml.set_cells_xml(xml, {(1, 1): 3})

    def test_furigana_ignored(self):
        si = ET.fromstring(
            f'<si xmlns="{ooxml.NS}"><t>山田</t><rPh sb="0" eb="2"><t>ヤマダ</t></rPh>'
            '<phoneticPr fontId="1"/></si>')
        self.assertEqual(ooxml._text_of_si(si), "山田")


class FillTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_roundtrip_and_preservation(self):
        src = os.path.join(self.tmp, "src.xlsx")
        # add a fake image/drawing part to prove untouched parts survive byte-for-byte
        with zipfile.ZipFile(KYUKA) as zin, zipfile.ZipFile(src, "w", zipfile.ZIP_DEFLATED) as zout:
            for i in zin.infolist():
                zout.writestr(i, zin.read(i.filename))
            zout.writestr("xl/media/image1.png", b"\x89PNG fake stamp bytes")
        out = os.path.join(self.tmp, "out.xlsx")
        rep = fill(src, out, {
            "日付": "2026-11-19", "氏名": "山田 太郎", "社員番号": 42,
            "休暇の種類": "代休", "開始日": "2026-12-01", "備考": "なし",
        })
        self.assertTrue(rep.to_dict()["ok"], rep.to_dict())
        self.assertTrue(rep.verified)
        wb = load(out)
        s = wb.sheet()
        self.assertEqual(s.value(5, 20), "令和8年11月19日")
        self.assertEqual(s.value(9, 8), "山田 太郎")
        self.assertEqual(s.value(8, 24), 42)
        self.assertEqual(s.value(12, 8), "□年次有給休暇　□特別休暇　■代休　□その他")
        self.assertEqual(s.value(13, 8), "令和8年12月1日")
        self.assertEqual(s.value(21, 2), "備考：なし")
        self.assertEqual(len(s.merges), len(load(src).sheet().merges))
        with zipfile.ZipFile(src) as a, zipfile.ZipFile(out) as b:
            self.assertEqual(a.namelist(), b.namelist())
            for n in a.namelist():
                if n in ("xl/worksheets/sheet1.xml", "xl/workbook.xml"):
                    continue
                self.assertEqual(a.read(n), b.read(n), n)
            self.assertIn(b'fullCalcOnLoad="1"', b.read("xl/workbook.xml"))
        # styles of written cells are kept (border style index unchanged)
        self.assertEqual(s.cells[(9, 8)].style, load(src).sheet().cells[(9, 8)].style)

    def test_excel_saved_fill_keeps_namespaces(self):
        out = os.path.join(self.tmp, "x.xlsx")
        rep = fill(EXCEL_SAVED, out, {"氏名": "山田", "日付": "2026-11-19"})
        self.assertTrue(rep.to_dict()["ok"], rep.to_dict())
        with zipfile.ZipFile(EXCEL_SAVED) as a, zipfile.ZipFile(out) as b:
            head_a = a.read("xl/worksheets/sheet1.xml")[:800]
            head_b = b.read("xl/worksheets/sheet1.xml")[:800]
            self.assertEqual(head_a.split(b"<sheetData")[0], head_b.split(b"<sheetData")[0])
        self.assertEqual(load(out).sheet().value(9, 8), "山田")

    def test_table_and_errors(self):
        out = os.path.join(self.tmp, "o.xlsx")
        rep = fill(KEIHI, out, {
            "明細": [{"日付": "11/1", "内容": "交通費", "金額": 2300}, ["", "11/2", "書籍", 1800]],
            "支払方法": "現金",
        })
        self.assertTrue(rep.to_dict()["ok"], rep.to_dict())
        s = load(out).sheet()
        self.assertEqual(s.value(10, 10), "交通費")
        self.assertEqual(s.value(10, 25), 2300)
        self.assertEqual(s.value(11, 25), 1800)
        self.assertEqual(s.value(10, 2), 1, "pre-numbered index column untouched")

        rep = fill(KEIHI, out, {"明細": [{}] * 6})
        self.assertTrue(rep.errors)
        rep = fill(KEIHI, out, {"Y15": 1})
        self.assertTrue(any("formula" in e for e in rep.errors))
        rep = fill(KEIHI, out, {"存在しない欄": 1})
        self.assertEqual(rep.unknown_keys, ["存在しない欄"])
        rep = fill(KEIHI, out, {"支払方法": "小切手"})
        self.assertTrue(rep.errors)

    def test_cell_key_and_openpyxl_compat(self):
        out = os.path.join(self.tmp, "p.xlsx")
        rep = fill(PLAIN, out, {"氏名": "鈴木", "Sheet1!D10": "extra", "生年月日": "2000-01-02"})
        self.assertTrue(rep.to_dict()["ok"], rep.to_dict())
        s = load(out).sheet()
        self.assertEqual(s.value(3, 2), "鈴木")
        self.assertEqual(s.value(10, 4), "extra")
        self.assertEqual(s.value(5, 2), "2000年1月2日", "no 和暦 in sheet -> 西暦 Japanese format")
        try:
            import openpyxl
        except ImportError:
            self.skipTest("openpyxl not installed")
        ws = openpyxl.load_workbook(out).active
        self.assertEqual(ws["B3"].value, "鈴木")


class FieldsFileTests(unittest.TestCase):
    def test_saved_fields_override_detection(self):
        tmp = tempfile.mkdtemp()
        try:
            fs = [f.to_dict() for f in detect(load(KEIHI))]
            # simulate a manual correction: rename a field id and point it elsewhere
            for d in fs:
                if d["id"] == "部署名":
                    d["id"] = "所属部署"
            from masume.detect import Field
            fields = [Field.from_dict(json.loads(json.dumps(d))) for d in fs]
            out = os.path.join(tmp, "o.xlsx")
            rep = fill(KEIHI, out, {"所属部署": "総務", "明細": [{"内容": "切手", "金額": 84}]}, fields=fields)
            self.assertTrue(rep.to_dict()["ok"], rep.to_dict())
            s = load(out).sheet()
            self.assertEqual(s.value(6, 8), "総務")
            self.assertEqual(s.value(10, 25), 84)
        finally:
            shutil.rmtree(tmp, ignore_errors=True)


class McpTests(unittest.TestCase):
    def test_stdio_session(self):
        msgs = [
            {"jsonrpc": "2.0", "id": 1, "method": "initialize",
             "params": {"protocolVersion": "2025-06-18", "capabilities": {}, "clientInfo": {"name": "t", "version": "0"}}},
            {"jsonrpc": "2.0", "method": "notifications/initialized"},
            {"jsonrpc": "2.0", "id": 2, "method": "tools/list"},
            {"jsonrpc": "2.0", "id": 3, "method": "tools/call",
             "params": {"name": "inspect_form", "arguments": {"path": KYUKA}}},
            {"jsonrpc": "2.0", "id": 4, "method": "tools/call",
             "params": {"name": "fill_form", "arguments": {"path": KYUKA, "values": {}, "output_path": KYUKA}}},
        ]
        env = dict(os.environ, PYTHONPATH=os.path.join(ROOT, "src"))
        p = subprocess.run([sys.executable, "-m", "masume.mcp_server"],
                           input="\n".join(json.dumps(m, ensure_ascii=False) for m in msgs).encode("utf-8"),
                           capture_output=True, env=env, timeout=60)
        out = [json.loads(line) for line in p.stdout.decode("utf-8").splitlines() if line.strip()]
        self.assertEqual([o["id"] for o in out], [1, 2, 3, 4])
        self.assertEqual(out[0]["result"]["serverInfo"]["name"], "masume")
        self.assertEqual({t["name"] for t in out[1]["result"]["tools"]},
                         {"list_sheets", "inspect_form", "form_to_markdown", "fill_form"})
        payload = json.loads(out[2]["result"]["content"][0]["text"])
        self.assertIn("氏名", payload["template"])
        self.assertTrue(out[3]["result"]["isError"], "must refuse to overwrite the template")


if __name__ == "__main__":
    unittest.main()
