"""Build fixtures_hard/checkbox_controls.xlsx (requires openpyxl).

Real municipal forms often use Excel *form-control* check boxes (開発タブの
チェックボックス) instead of □ characters. openpyxl cannot create them, so this
script writes the grid with openpyxl and then injects the control parts the way
Excel stores them: xl/ctrlProps/ctrlPropN.xml, a legacy VML drawing and the
<controls> block of the sheet.

    python tests/make_control_fixture.py
"""
from __future__ import annotations

import os
import shutil
import tempfile
import zipfile

from openpyxl import Workbook
from openpyxl.styles import Border, Side

HERE = os.path.join(os.path.dirname(__file__), "fixtures_hard")
T = Side(style="thin")
# (shape id, col0, row0) of each box: anchors are 0-based like the XML
BOXES = [(1025, 1, 3, False), (1026, 1, 4, True), (1027, 8, 3, False), (1028, 8, 4, False)]


def _grid(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "変更届"
    ws["B2"] = "⑴　名称変更"
    ws["D4"] = "社名変更"
    ws["D5"] = "法人成り"
    ws["I2"] = "⑵　所在地変更"
    ws["K4"] = "事務所等の移転"
    ws["K5"] = "その他"
    ws.merge_cells("B8:E8"); ws["B8"] = "氏名"
    ws.merge_cells("F8:N8")
    for c in range(2, 15):
        ws.cell(8, c).border = Border(top=T, bottom=T, left=T if c in (2, 6) else None, right=T if c == 14 else None)
    wb.save(path)


def _control(sid, rid, col, row):
    return (
        '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
        '<mc:Choice Requires="x14">'
        f'<control shapeId="{sid}" r:id="{rid}" name="Check Box {sid - 1024}">'
        '<controlPr defaultSize="0" autoFill="0" autoLine="0" autoPict="0">'
        '<anchor moveWithCells="1">'
        f'<from><xdr:col>{col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row - 1}</xdr:row><xdr:rowOff>95250</xdr:rowOff></from>'
        f'<to><xdr:col>{col + 1}</xdr:col><xdr:colOff>190500</xdr:colOff><xdr:row>{row + 1}</xdr:row><xdr:rowOff>28575</xdr:rowOff></to>'
        '</anchor></controlPr></control></mc:Choice></mc:AlternateContent>'
    )


def _vml_shape(sid, col, row, checked):
    tick = "<x:Checked>1</x:Checked>" if checked else ""
    return (
        f'<v:shape id="_x0000_s{sid}" type="#_x0000_t201" style="position:absolute;'
        f'margin-left:{col * 40}pt;margin-top:{row * 15}pt;width:20pt;height:20pt;z-index:{sid - 1024}" '
        'filled="f" stroked="f" o:insetmode="auto"><v:textbox><div style="text-align:left"></div></v:textbox>'
        '<x:ClientData ObjectType="Checkbox"><x:SizeWithCells/>'
        f'<x:Anchor>{col}, 0, {row - 1}, 8, {col + 1}, 15, {row + 1}, 2</x:Anchor>'
        f'<x:AutoFill>False</x:AutoFill><x:AutoLine>False</x:AutoLine>{tick}<x:NoThreeD/></x:ClientData></v:shape>'
    )


def build(path):
    tmp = tempfile.mkdtemp()
    try:
        base = os.path.join(tmp, "base.xlsx")
        _grid(base)
        with zipfile.ZipFile(base) as zin:
            parts = {n: zin.read(n) for n in zin.namelist()}
        sheet = parts["xl/worksheets/sheet1.xml"].decode("utf-8")
        if 'xmlns:r="' not in sheet:
            sheet = sheet.replace("<worksheet ", '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ', 1)
        sheet = sheet.replace(
            "<worksheet ",
            '<worksheet xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
            'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" ', 1)
        controls = "".join(_control(sid, f"rId{i + 2}", col, row) for i, (sid, col, row, _) in enumerate(BOXES))
        tail = ('<legacyDrawing r:id="rId1"/>'
                '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
                '<mc:Choice Requires="x14"><controls>' + controls + "</controls></mc:Choice></mc:AlternateContent>")
        anchor = sheet.find("<tableParts")
        if anchor < 0:
            anchor = sheet.rfind("</worksheet>")
        sheet = sheet[:anchor] + tail + sheet[anchor:]
        parts["xl/worksheets/sheet1.xml"] = sheet.encode("utf-8")
        rels = ['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>']
        for i, (sid, col, row, checked) in enumerate(BOXES):
            rels.append(f'<Relationship Id="rId{i + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp" Target="../ctrlProps/ctrlProp{i + 1}.xml"/>')
            attr = ' checked="Checked"' if checked else ""
            parts[f"xl/ctrlProps/ctrlProp{i + 1}.xml"] = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" '
                f'objectType="CheckBox"{attr} lockText="1" noThreeD="1"/>').encode("utf-8")
        parts["xl/worksheets/_rels/sheet1.xml.rels"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join(rels) + "</Relationships>").encode("utf-8")
        parts["xl/drawings/vmlDrawing1.vml"] = (
            '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" '
            'xmlns:x="urn:schemas-microsoft-com:office:excel">'
            '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>'
            '<v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201" path="m,l,21600r21600,l21600,xe">'
            '<v:stroke joinstyle="miter"/><v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/>'
            '<o:lock v:ext="edit" shapetype="t"/></v:shapetype>'
            + "".join(_vml_shape(sid, col, row, checked) for sid, col, row, checked in BOXES)
            + "</xml>").encode("utf-8")
        ct = parts["[Content_Types].xml"].decode("utf-8")
        add = '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>'
        add += "".join(f'<Override PartName="/xl/ctrlProps/ctrlProp{i + 1}.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/>'
                       for i in range(len(BOXES)))
        parts["[Content_Types].xml"] = ct.replace("</Types>", add + "</Types>").encode("utf-8")
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
            order = ["[Content_Types].xml"] + [n for n in parts if n != "[Content_Types].xml"]
            for n in order:
                zout.writestr(n, parts[n])
    finally:
        shutil.rmtree(tmp)


if __name__ == "__main__":
    os.makedirs(HERE, exist_ok=True)
    out = os.path.join(HERE, "checkbox_controls.xlsx")
    build(out)
    print("wrote", out)
