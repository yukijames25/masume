"""Minimal, dependency-free reader/writer for .xlsx (SpreadsheetML) packages.

Reading uses ElementTree. Writing is done by *string surgery* on the sheet XML
so that everything we do not touch (shapes, images, stamps, print settings,
namespace prefixes that Excel is picky about) is preserved byte-for-byte.
"""
from __future__ import annotations

import math
import os
import posixpath
import re
import tempfile
import zipfile
from dataclasses import dataclass, field
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_RNS = "http://schemas.openxmlformats.org/package/2006/relationships"
_N = "{%s}" % NS


# --------------------------------------------------------------------------
# cell reference helpers
# --------------------------------------------------------------------------
_REF_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")
MAX_ROW = 1_048_576
MAX_COL = 16_384
MAX_TEXT = 32_767  # Excel's per-cell limit, counted in UTF-16 code units


def col_to_idx(col: str) -> int:
    n = 0
    for ch in col.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def idx_to_col(idx: int) -> str:
    s = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        s = chr(65 + rem) + s
    return s


def parse_ref(ref: str) -> tuple[int, int]:
    """'B3' -> (row=3, col=2)"""
    m = _REF_RE.match(ref.strip())
    if not m:
        raise ValueError(f"invalid cell reference: {ref!r}")
    row, col = int(m.group(2)), col_to_idx(m.group(1))
    if not (1 <= row <= MAX_ROW and 1 <= col <= MAX_COL):
        raise ValueError(f"cell reference out of Excel's range: {ref!r}")
    return row, col


def make_ref(row: int, col: int) -> str:
    return f"{idx_to_col(col)}{row}"


def parse_range(rng: str) -> tuple[int, int, int, int]:
    """'B3:D4' -> (r1, c1, r2, c2)"""
    if ":" in rng:
        a, b = rng.split(":", 1)
    else:
        a = b = rng
    r1, c1 = parse_ref(a)
    r2, c2 = parse_ref(b)
    return min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)


# --------------------------------------------------------------------------
# data model
# --------------------------------------------------------------------------
@dataclass
class Cell:
    row: int
    col: int
    value: object = None  # str | float | int | bool | None
    style: int = 0
    formula: str | None = None

    @property
    def ref(self) -> str:
        return make_ref(self.row, self.col)


@dataclass
class Control:
    """A form control (チェックボックス) floating over the grid."""
    shape_id: str
    kind: str  # CheckBox | Radio | ...
    r1: int  # 1-based cells under the control's anchor
    c1: int
    r2: int
    c2: int
    checked: bool
    prop_part: str  # xl/ctrlProps/ctrlPropN.xml
    vml_part: str | None = None
    text: str = ""  # caption drawn by the control itself

    @property
    def row(self) -> int:
        return (self.r1 + self.r2) // 2


@dataclass
class Sheet:
    name: str
    path: str  # path inside the zip, e.g. xl/worksheets/sheet1.xml
    cells: dict[tuple[int, int], Cell] = field(default_factory=dict)
    merges: list[tuple[int, int, int, int]] = field(default_factory=list)
    hidden: bool = False
    locked: list = field(default_factory=list)  # [((r1, c1, r2, c2), reason)]
    controls: list[Control] = field(default_factory=list)

    def lock_reason(self, row: int, col: int) -> str | None:
        for (r1, c1, r2, c2), why in self.locked:
            if r1 <= row <= r2 and c1 <= col <= c2:
                return why
        return None

    def merge_of(self, row: int, col: int) -> tuple[int, int, int, int] | None:
        for m in self.merges:
            if m[0] <= row <= m[2] and m[1] <= col <= m[3]:
                return m
        return None

    def value(self, row: int, col: int):
        c = self.cells.get((row, col))
        return None if c is None else c.value

    @property
    def max_row(self) -> int:
        rows = [r for r, _ in self.cells] + [m[2] for m in self.merges]
        return max(rows) if rows else 0

    @property
    def max_col(self) -> int:
        cols = [c for _, c in self.cells] + [m[3] for m in self.merges]
        return max(cols) if cols else 0


@dataclass
class Styles:
    bordered: set[int] = field(default_factory=set)  # cellXfs indices drawing any border
    sides: dict[int, frozenset] = field(default_factory=dict)  # cellXfs index -> {"left",...}
    heading: set[int] = field(default_factory=set)  # bold or large font (titles / headings)


@dataclass
class Workbook:
    path: str
    sheets: list[Sheet]
    styles: Styles

    @property
    def bordered_styles(self) -> set[int]:
        return self.styles.bordered

    def sheet(self, name: str | None = None) -> Sheet:
        if name is None:
            for s in self.sheets:
                if not s.hidden:
                    return s
            return self.sheets[0]
        for s in self.sheets:
            if s.name == name:
                return s
        raise KeyError(f"sheet not found: {name!r} (have: {[s.name for s in self.sheets]})")


# --------------------------------------------------------------------------
# reading
# --------------------------------------------------------------------------
def _text_of_si(si: ET.Element) -> str:
    """Text of a <si> / <is> element, ignoring phonetic (furigana) runs."""
    parts: list[str] = []
    for child in si:
        tag = child.tag
        if tag == _N + "t":
            parts.append(child.text or "")
        elif tag == _N + "r":
            t = child.find(_N + "t")
            if t is not None:
                parts.append(t.text or "")
        # _N+"rPh" (furigana) and _N+"phoneticPr" are intentionally skipped
    return decode_x("".join(parts))


_XESC = re.compile(r"_x([0-9A-Fa-f]{4})_")


def decode_x(text: str) -> str:
    """Decode OOXML ``_xHHHH_`` escapes the way Excel displays them."""
    if "_x" not in text:
        return text
    return _XESC.sub(lambda m: chr(int(m.group(1), 16)), text)


def _resolve(base_dir: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(base_dir, target))


def _read_styles(zf: zipfile.ZipFile) -> Styles:
    st = Styles()
    try:
        root = ET.fromstring(zf.read("xl/styles.xml"))
    except KeyError:
        return st
    border_sides: list[frozenset] = []
    borders = root.find(_N + "borders")
    if borders is not None:
        for b in borders.findall(_N + "border"):
            sides = set()
            for side in ("left", "right", "top", "bottom"):
                el = b.find(_N + side)
                if el is not None and el.get("style") not in (None, "none"):
                    sides.add(side)
            border_sides.append(frozenset(sides))
    heading_fonts: set[int] = set()
    fonts = root.find(_N + "fonts")
    if fonts is not None:
        for i, f in enumerate(fonts.findall(_N + "font")):
            b = f.find(_N + "b")
            sz = f.find(_N + "sz")
            bold = b is not None and b.get("val", "1") not in ("0", "false")
            try:
                big = sz is not None and float(sz.get("val", "11")) >= 14
            except ValueError:
                big = False
            if bold or big:
                heading_fonts.add(i)
    xfs = root.find(_N + "cellXfs")
    if xfs is not None:
        for i, xf in enumerate(xfs.findall(_N + "xf")):
            bid = int(xf.get("borderId", "0"))
            if bid < len(border_sides) and border_sides[bid]:
                st.bordered.add(i)
                st.sides[i] = border_sides[bid]
            if int(xf.get("fontId", "0")) in heading_fonts:
                st.heading.add(i)
    return st


def load(path: str) -> Workbook:
    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
        shared: list[str] = []
        if "xl/sharedStrings.xml" in names:
            sroot = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            shared = [_text_of_si(si) for si in sroot.findall(_N + "si")]

        wb = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rel_target = {
            r.get("Id"): _resolve("xl", r.get("Target", ""))
            for r in rels.findall("{%s}Relationship" % PKG_RNS)
        }
        sheets: list[Sheet] = []
        sheets_el = wb.find(_N + "sheets")
        for s in (sheets_el if sheets_el is not None else []):
            rid = s.get("{%s}id" % RNS)
            spath = rel_target.get(rid)
            if not spath or spath not in names:
                continue  # chartsheet / dialogsheet / missing
            sheet = Sheet(
                name=s.get("name", ""),
                path=spath,
                hidden=s.get("state", "visible") != "visible",
            )
            data = zf.read(spath)
            _read_sheet(data, sheet, shared)
            _read_tables(zf, names, spath, sheet)
            _read_controls(zf, names, spath, data, sheet)
            sheets.append(sheet)
        return Workbook(path=path, sheets=sheets, styles=_read_styles(zf))


def _read_tables(zf: zipfile.ZipFile, names: set, spath: str, sheet: Sheet) -> None:
    """Header rows of Excel tables (ListObjects) must match their column names."""
    d, fn = posixpath.split(spath)
    rels_path = f"{d}/_rels/{fn}.rels"
    if rels_path not in names:
        return
    rels = ET.fromstring(zf.read(rels_path))
    for r in rels.findall("{%s}Relationship" % PKG_RNS):
        if not r.get("Type", "").endswith("/table"):
            continue
        tpath = _resolve(d, r.get("Target", ""))
        if tpath not in names:
            continue
        t = ET.fromstring(zf.read(tpath))
        ref = t.get("ref")
        if not ref:
            continue
        r1, c1, r2, c2 = parse_range(ref)
        if t.get("headerRowCount", "1") != "0":
            sheet.locked.append(((r1, c1, r1, c2), f"header row of Excel table {t.get('displayName', '')}"))
        if t.get("totalsRowCount", "0") not in ("0", ""):
            sheet.locked.append(((r2, c1, r2, c2), "totals row of an Excel table"))


_XDR = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"
_VML_SHAPE = r'<v:shape\b[^>]*\b(?:id|o:spid)="_x0000_s{sid}"[^>]*>.*?</v:shape>'


def _sheet_rels(zf: zipfile.ZipFile, names: set, spath: str) -> dict[str, tuple[str, str]]:
    d, fn = posixpath.split(spath)
    rels_path = f"{d}/_rels/{fn}.rels"
    if rels_path not in names:
        return {}
    rels = ET.fromstring(zf.read(rels_path))
    return {r.get("Id"): (r.get("Type", ""), _resolve(d, r.get("Target", "")))
            for r in rels.findall("{%s}Relationship" % PKG_RNS)}


def _read_controls(zf: zipfile.ZipFile, names: set, spath: str, data: bytes, sheet: Sheet) -> None:
    """Form-control check boxes (xl/ctrlProps) and their legacy VML twins."""
    if b"<control" not in data and b":control" not in data:
        return
    rels = _sheet_rels(zf, names, spath)
    vml = next((t for (ty, t) in rels.values() if ty.endswith("/vmlDrawing") and t in names), None)
    vml_text = zf.read(vml).decode("utf-8", "replace") if vml else ""
    root = ET.fromstring(data)
    seen: set[str] = set()
    for el in root.iter(_N + "control"):
        sid = el.get("shapeId")
        rid = el.get("{%s}id" % RNS)
        anchor = el.find(f"{_N}controlPr/{_N}anchor")
        if not sid or sid in seen or anchor is None or rid not in rels:
            continue
        ty, target = rels[rid]
        if not ty.endswith("/ctrlProp") or target not in names:
            continue
        try:
            fr, to = anchor.find(_N + "from"), anchor.find(_N + "to")
            c1, r1 = int(fr.find(_XDR + "col").text) + 1, int(fr.find(_XDR + "row").text) + 1
            c2, r2 = int(to.find(_XDR + "col").text) + 1, int(to.find(_XDR + "row").text) + 1
        except (AttributeError, TypeError, ValueError):
            continue
        prop = ET.fromstring(zf.read(target))
        text = ""
        m = re.search(_VML_SHAPE.format(sid=re.escape(sid)), vml_text, re.S) if vml_text else None
        if m:
            div = re.search(r"<v:textbox\b.*?</v:textbox>", m.group(0), re.S)
            if div:
                text = re.sub(r"<[^>]+>", "", div.group(0)).strip()
        seen.add(sid)
        sheet.controls.append(Control(
            shape_id=sid, kind=prop.get("objectType", ""), r1=r1, c1=c1, r2=r2, c2=c2,
            checked=prop.get("checked") == "Checked", prop_part=target,
            vml_part=vml if m else None, text=decode_x(text)))


def set_control_prop(xml: str, checked: bool) -> str:
    """Tick / untick a <formControlPr> (x14 form control)."""
    m = re.search(r"<(\w+:)?formControlPr\b[^>]*?(/?)>", xml)
    if not m:
        raise WriteError("form control without formControlPr")
    tag = re.sub(r'\s+checked="[^"]*"', "", m.group(0))
    if checked:
        tag = re.sub(r"\s*(/?)>$", r' checked="Checked"\1>', tag)
    return xml[:m.start()] + tag + xml[m.end():]


def set_control_vml(vml: str, shape_id: str, checked: bool) -> str:
    """Same for the legacy VML copy (<x:Checked>) that older readers use."""
    m = re.search(_VML_SHAPE.format(sid=re.escape(shape_id)), vml, re.S)
    if not m:
        return vml
    shape = re.sub(r"\s*<x:Checked>[^<]*</x:Checked>", "", m.group(0))
    if checked:
        shape = shape.replace("</x:ClientData>", "<x:Checked>1</x:Checked></x:ClientData>", 1)
    return vml[:m.start()] + shape + vml[m.end():]


def _read_sheet(data: bytes, sheet: Sheet, shared: list[str]) -> None:
    root = ET.fromstring(data)
    sd = root.find(_N + "sheetData")
    if sd is not None:
        prev_r = 0
        for row_el in sd.findall(_N + "row"):
            r = int(row_el.get("r", "0")) or prev_r + 1
            prev_r = r
            next_col = 1
            for c_el in row_el.findall(_N + "c"):
                ref = c_el.get("r")
                if ref:
                    row, col = parse_ref(ref)
                else:
                    row, col = r, next_col
                next_col = col + 1
                t = c_el.get("t", "n")
                v_el = c_el.find(_N + "v")
                f_el = c_el.find(_N + "f")
                if f_el is not None and f_el.get("t") in ("array", "dataTable") and f_el.get("ref"):
                    try:
                        sheet.locked.append((parse_range(f_el.get("ref")), "inside an array formula / spill range"))
                    except ValueError:
                        pass
                raw = v_el.text if v_el is not None else None
                value: object = None
                if t == "s" and raw is not None:
                    try:
                        value = shared[int(raw)]
                    except (ValueError, IndexError):
                        value = None
                elif t == "inlineStr":
                    is_el = c_el.find(_N + "is")
                    value = _text_of_si(is_el) if is_el is not None else None
                elif t == "str":
                    value = decode_x(raw) if raw is not None else None
                elif t == "e":
                    value = raw
                elif t == "b":
                    value = raw == "1"
                elif raw is not None:
                    try:
                        f = float(raw)
                        value = int(f) if f.is_integer() and "E" not in raw.upper() else f
                    except ValueError:
                        value = raw
                sheet.cells[(row, col)] = Cell(
                    row=row,
                    col=col,
                    value=value,
                    style=int(c_el.get("s", "0")),
                    formula=(f_el.text or "") if f_el is not None else None,
                )
    mc = root.find(_N + "mergeCells")
    if mc is not None:
        for m in mc.findall(_N + "mergeCell"):
            ref = m.get("ref")
            if ref:
                sheet.merges.append(parse_range(ref))


# --------------------------------------------------------------------------
# writing (non-destructive string surgery)
# --------------------------------------------------------------------------
class WriteError(Exception):
    pass


_ILLEGAL = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff\ud800-\udfff]")


def clean_text(value: str) -> str:
    """Remove characters Excel cannot store (XML-illegal, U+FFFE/FFFF, lone surrogates)."""
    return _ILLEGAL.sub("", value)


def check_value(value) -> None:
    if isinstance(value, bool) or value is None:
        return
    if isinstance(value, (int, float)):
        try:
            f = float(value)
        except OverflowError:
            raise WriteError(f"number too large for Excel: {value!r}") from None
        if not math.isfinite(f):
            raise WriteError(f"Excel cannot store {value!r}")
        return
    text = clean_text(str(value))
    if len(text.encode("utf-16-le")) // 2 > MAX_TEXT:
        raise WriteError(f"text longer than Excel's limit of {MAX_TEXT} characters")


def _xml_text(value: str) -> str:
    value = clean_text(value)
    # protect a literal "_xHHHH_" so Excel does not decode it
    value = re.sub(r"_x(?=[0-9A-Fa-f]{4}_)", "_x005F_x", value)
    return escape(value)


def _cell_xml(prefix: str, ref: str, style: str | None, value) -> str:
    s_attr = f' s="{style}"' if style not in (None, "0") else ""
    p = prefix
    if value is None or value == "":
        return f'<{p}c r="{ref}"{s_attr}/>'
    if isinstance(value, bool):
        return f'<{p}c r="{ref}"{s_attr} t="b"><{p}v>{int(value)}</{p}v></{p}c>'
    if isinstance(value, (int, float)):
        check_value(value)
        num = repr(int(value)) if isinstance(value, int) else repr(float(value))
        return f'<{p}c r="{ref}"{s_attr}><{p}v>{num}</{p}v></{p}c>'
    check_value(value)
    text = _xml_text(str(value))
    return (
        f'<{p}c r="{ref}"{s_attr} t="inlineStr"><{p}is>'
        f'<{p}t xml:space="preserve">{text}</{p}t></{p}is></{p}c>'
    )


def _row_pattern(prefix: str) -> re.Pattern:
    p = re.escape(prefix)
    return re.compile(
        rf"<{p}row\b(?P<attrs>[^>]*?)(?:/>|>(?P<body>.*?)</{p}row>)", re.S
    )


def _cell_pattern(prefix: str) -> re.Pattern:
    p = re.escape(prefix)
    return re.compile(rf"<{p}c\b(?P<attrs>[^>]*?)(?:/>|>(?P<body>.*?)</{p}c>)", re.S)


_ATTR_R = re.compile(r'\br="([^"]+)"')
_ATTR_S = re.compile(r'\bs="(\d+)"')


def set_cells_xml(xml: str, updates: dict[tuple[int, int], object]) -> str:
    """Return sheet XML with the given (row, col) cells set to new values."""
    m = re.search(r"<(\w+:)?sheetData\b", xml)
    if not m:
        raise WriteError("sheetData not found")
    prefix = m.group(1) or ""
    p = re.escape(prefix)

    # normalise self-closing <sheetData/>
    xml = re.sub(rf"<{p}sheetData\s*/>", f"<{prefix}sheetData></{prefix}sheetData>", xml, count=1)
    sd_open = re.search(rf"<{p}sheetData\b[^>]*>", xml)
    sd_close = xml.find(f"</{prefix}sheetData>", sd_open.end())
    head, body, tail = xml[: sd_open.end()], xml[sd_open.end(): sd_close], xml[sd_close:]

    by_row: dict[int, dict[int, object]] = {}
    for (r, c), v in updates.items():
        by_row.setdefault(r, {})[c] = v

    row_re = _row_pattern(prefix)
    cell_re = _cell_pattern(prefix)

    rows = list(row_re.finditer(body))
    row_index = {}
    for rm in rows:
        ra = _ATTR_R.search(rm.group("attrs"))
        if not ra:
            raise WriteError("sheet has rows without r= attributes; re-save it in Excel first")
        row_index[int(ra.group(1))] = rm
    col_styles = _col_styles(head)

    # process rows from bottom to top so earlier offsets stay valid
    new_body = body
    for r in sorted(by_row, reverse=True):
        cells = by_row[r]
        rm = row_index.get(r)
        if rm is None:
            cell_xml = "".join(
                _cell_xml(prefix, make_ref(r, c), col_styles.get(c), v) for c, v in sorted(cells.items())
            )
            new_row = f'<{prefix}row r="{r}">{cell_xml}</{prefix}row>'
            # before the first existing row with a larger index, else at the ORIGINAL
            # end of sheetData: rows are processed bottom-up, so this keeps them ascending
            pos = len(body)
            for rr in sorted(row_index):
                if rr > r:
                    pos = row_index[rr].start()
                    break
            new_body = new_body[:pos] + new_row + new_body[pos:]
            continue

        attrs = rm.group("attrs")
        row_body = rm.group("body") or ""
        inserted = False
        row_style = None
        if re.search(r'\bcustomFormat="(1|true)"', attrs):
            sm_row = _ATTR_S.search(attrs)
            row_style = sm_row.group(1) if sm_row else None
        # update / insert each cell (right to left)
        for c in sorted(cells, reverse=True):
            value = cells[c]
            ref = make_ref(r, c)
            found = None
            later = None
            for cm in cell_re.finditer(row_body):
                ca = _ATTR_R.search(cm.group("attrs"))
                if not ca:
                    raise WriteError(f"row {r} has cells without r= attributes; unsupported")
                rr, cc = parse_ref(ca.group(1))
                if cc == c:
                    found = cm
                    break
                if cc > c and later is None:
                    later = cm
            if found is not None:
                cbody = found.group("body") or ""
                if re.search(rf"<{p}f\b", cbody):
                    raise WriteError(f"{ref} contains a formula; refusing to overwrite")
                sm = _ATTR_S.search(found.group("attrs"))
                new_cell = _cell_xml(prefix, ref, sm.group(1) if sm else None, value)
                row_body = row_body[: found.start()] + new_cell + row_body[found.end():]
            else:
                new_cell = _cell_xml(prefix, ref, row_style or col_styles.get(c), value)
                pos = later.start() if later is not None else len(row_body)
                row_body = row_body[:pos] + new_cell + row_body[pos:]
                inserted = True
        if inserted:
            attrs = re.sub(r'\s+spans="[^"]*"', "", attrs)
        attrs = attrs.rstrip().rstrip("/")
        new_row = f"<{prefix}row{attrs}>{row_body}</{prefix}row>"
        new_body = new_body[: rm.start()] + new_row + new_body[rm.end():]

    return head + new_body + tail


def _col_styles(head: str) -> dict[int, str]:
    """Column default styles from <cols> (方眼紙 forms often set the font sheet-wide)."""
    out: dict[int, str] = {}
    for m in re.finditer(r"<(?:\w+:)?col\b([^>]*)/?>", head):
        a = m.group(1)
        st = re.search(r'\bstyle="(\d+)"', a)
        mn = re.search(r'\bmin="(\d+)"', a)
        mx = re.search(r'\bmax="(\d+)"', a)
        if not (st and mn and mx) or st.group(1) == "0":
            continue
        lo, hi = int(mn.group(1)), min(int(mx.group(1)), MAX_COL)
        for c in range(lo, hi + 1):
            out[c] = st.group(1)
    return out


def check_sheet_xml(xml: str) -> list[str]:
    """Structural rules Excel enforces: rows and cells ascending, inside the grid."""
    problems = []
    m = re.search(r"<(\w+:)?sheetData\b", xml)
    prefix = (m.group(1) if m else "") or ""
    prev_r = 0
    for rm in _row_pattern(prefix).finditer(xml):
        ra = _ATTR_R.search(rm.group("attrs"))
        r = int(ra.group(1)) if ra else prev_r + 1
        if not (prev_r < r <= MAX_ROW):
            problems.append(f"row {r} out of order or range (after {prev_r})")
        prev_r = r
        prev_c = 0
        for cm in _cell_pattern(prefix).finditer(rm.group("body") or ""):
            ca = _ATTR_R.search(cm.group("attrs"))
            if not ca:
                continue
            try:
                rr, cc = parse_ref(ca.group(1))
            except ValueError as e:
                problems.append(str(e))
                continue
            if rr != r or cc <= prev_c:
                problems.append(f"cell {ca.group(1)} out of order in row {r}")
            prev_c = cc
    return problems


def _force_recalc(workbook_xml: str) -> str:
    m = re.search(r"<(\w+:)?calcPr\b([^>]*?)(/?)>", workbook_xml)
    if m:
        attrs = m.group(2)
        if "fullCalcOnLoad" in attrs:
            attrs = re.sub(r'fullCalcOnLoad="[^"]*"', 'fullCalcOnLoad="1"', attrs)
        else:
            attrs = attrs + ' fullCalcOnLoad="1"'
        return workbook_xml[: m.start()] + f"<{m.group(1) or ''}calcPr{attrs}{m.group(3)}>" + workbook_xml[m.end():]
    m = re.search(r"<(\w+:)?workbook\b", workbook_xml)
    prefix = (m.group(1) if m else "") or ""
    anchor = -1
    for tag in ("sheets", "functionGroups", "externalReferences", "definedNames"):
        i = workbook_xml.rfind(f"</{prefix}{tag}>")
        if i >= 0:
            anchor = max(anchor, i + len(f"</{prefix}{tag}>"))
    if anchor < 0:
        return workbook_xml
    return workbook_xml[:anchor] + f'<{prefix}calcPr fullCalcOnLoad="1"/>' + workbook_xml[anchor:]


def save_with_updates(
    src: str,
    dst: str,
    updates: dict[str, dict[tuple[int, int], object]],
    recalc: bool = True,
    controls: list[tuple[Control, bool]] | None = None,
) -> None:
    """Copy *src* to *dst*, replacing values of cells per sheet path.

    updates: { "xl/worksheets/sheet1.xml": {(row, col): value, ...}, ... }
    controls: [(Control, checked), ...] check boxes to tick or untick
    """
    if os.path.exists(dst) and os.path.samefile(src, dst):
        raise WriteError("output path must differ from the template")
    with zipfile.ZipFile(src) as zin:
        replaced: dict[str, bytes] = {}
        for spath, cells in updates.items():
            if not cells:
                continue
            xml = zin.read(spath).decode("utf-8")
            new_xml = set_cells_xml(xml, cells)
            problems = check_sheet_xml(new_xml)
            if problems:
                raise WriteError("refusing to write a sheet Excel would reject: " + "; ".join(problems[:3]))
            replaced[spath] = new_xml.encode("utf-8")
        for ctl, checked in controls or []:
            cur = replaced.get(ctl.prop_part) or zin.read(ctl.prop_part)
            replaced[ctl.prop_part] = set_control_prop(cur.decode("utf-8"), checked).encode("utf-8")
            if ctl.vml_part:
                raw = replaced.get(ctl.vml_part) or zin.read(ctl.vml_part)
                enc = "utf-8" if raw[:200].lower().find(b"windows-1252") < 0 else "cp1252"
                replaced[ctl.vml_part] = set_control_vml(raw.decode(enc, "replace"), ctl.shape_id,
                                                         checked).encode(enc, "replace")
        if recalc and any(k in updates for k in replaced) and "xl/workbook.xml" in zin.namelist():
            replaced["xl/workbook.xml"] = _force_recalc(
                zin.read("xl/workbook.xml").decode("utf-8")
            ).encode("utf-8")
        out_dir = os.path.dirname(os.path.abspath(dst)) or "."
        fd, tmp = tempfile.mkstemp(prefix=".masume-", suffix=".xlsx", dir=out_dir)
        os.close(fd)
        try:
            seen: set[str] = set()
            with zipfile.ZipFile(tmp, "w") as zout:
                for info in zin.infolist():
                    if info.filename in seen:
                        continue  # duplicate entries confuse Excel
                    seen.add(info.filename)
                    data = replaced.get(info.filename)
                    if data is None:
                        try:
                            data = zin.read(info.filename)
                        except NotImplementedError as e:
                            raise WriteError(f"unsupported zip compression in template: {e}") from None
                    zi = zipfile.ZipInfo(info.filename, date_time=info.date_time)
                    zi.compress_type = (info.compress_type if info.compress_type in
                                        (zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED) else zipfile.ZIP_DEFLATED)
                    zi.external_attr = info.external_attr
                    zout.writestr(zi, data)
            os.replace(tmp, dst)
        except BaseException:
            try:
                os.remove(tmp)
            except OSError:
                pass
            raise
