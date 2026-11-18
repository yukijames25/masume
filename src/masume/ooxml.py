"""Minimal, dependency-free reader/writer for .xlsx (SpreadsheetML) packages.

Reading uses ElementTree. Writing is done by *string surgery* on the sheet XML
so that everything we do not touch (shapes, images, stamps, print settings,
namespace prefixes that Excel is picky about) is preserved byte-for-byte.
"""
from __future__ import annotations

import posixpath
import re
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
    return int(m.group(2)), col_to_idx(m.group(1))


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
class Sheet:
    name: str
    path: str  # path inside the zip, e.g. xl/worksheets/sheet1.xml
    cells: dict[tuple[int, int], Cell] = field(default_factory=dict)
    merges: list[tuple[int, int, int, int]] = field(default_factory=list)
    hidden: bool = False

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
    return "".join(parts)


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
            _read_sheet(zf.read(spath), sheet, shared)
            sheets.append(sheet)
        return Workbook(path=path, sheets=sheets, styles=_read_styles(zf))


def _read_sheet(data: bytes, sheet: Sheet, shared: list[str]) -> None:
    root = ET.fromstring(data)
    sd = root.find(_N + "sheetData")
    if sd is not None:
        for row_el in sd.findall(_N + "row"):
            r = int(row_el.get("r", "0"))
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
                elif t in ("str", "e"):
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


def _xml_text(value: str) -> str:
    # strip characters that are illegal in XML 1.0
    value = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", value)
    return escape(value)


def _cell_xml(prefix: str, ref: str, style: str | None, value) -> str:
    s_attr = f' s="{style}"' if style not in (None, "0") else ""
    p = prefix
    if value is None or value == "":
        return f'<{p}c r="{ref}"{s_attr}/>'
    if isinstance(value, bool):
        return f'<{p}c r="{ref}"{s_attr} t="b"><{p}v>{int(value)}</{p}v></{p}c>'
    if isinstance(value, (int, float)):
        return f'<{p}c r="{ref}"{s_attr}><{p}v>{value!r}</{p}v></{p}c>'
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
        if ra:
            row_index[int(ra.group(1))] = rm

    # process rows from bottom to top so earlier offsets stay valid
    new_body = body
    for r in sorted(by_row, reverse=True):
        cells = by_row[r]
        rm = row_index.get(r)
        if rm is None:
            cell_xml = "".join(
                _cell_xml(prefix, make_ref(r, c), None, v) for c, v in sorted(cells.items())
            )
            new_row = f'<{prefix}row r="{r}">{cell_xml}</{prefix}row>'
            # insert before first existing row with a larger index
            pos = len(new_body)
            for rr in sorted(row_index):
                if rr > r:
                    pos = row_index[rr].start()
                    break
            new_body = new_body[:pos] + new_row + new_body[pos:]
            continue

        attrs = rm.group("attrs")
        row_body = rm.group("body") or ""
        inserted = False
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
                new_cell = _cell_xml(prefix, ref, None, value)
                pos = later.start() if later is not None else len(row_body)
                row_body = row_body[:pos] + new_cell + row_body[pos:]
                inserted = True
        if inserted:
            attrs = re.sub(r'\s+spans="[^"]*"', "", attrs)
        attrs = attrs.rstrip().rstrip("/")
        new_row = f"<{prefix}row{attrs}>{row_body}</{prefix}row>"
        new_body = new_body[: rm.start()] + new_row + new_body[rm.end():]

    return head + new_body + tail


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
) -> None:
    """Copy *src* to *dst*, replacing values of cells per sheet path.

    updates: { "xl/worksheets/sheet1.xml": {(row, col): value, ...}, ... }
    """
    with zipfile.ZipFile(src) as zin:
        replaced: dict[str, bytes] = {}
        for spath, cells in updates.items():
            if not cells:
                continue
            xml = zin.read(spath).decode("utf-8")
            replaced[spath] = set_cells_xml(xml, cells).encode("utf-8")
        if recalc and replaced and "xl/workbook.xml" in zin.namelist():
            replaced["xl/workbook.xml"] = _force_recalc(
                zin.read("xl/workbook.xml").decode("utf-8")
            ).encode("utf-8")
        with zipfile.ZipFile(dst, "w") as zout:
            for info in zin.infolist():
                data = replaced.get(info.filename)
                if data is None:
                    data = zin.read(info.filename)
                zi = zipfile.ZipInfo(info.filename, date_time=info.date_time)
                zi.compress_type = info.compress_type
                zi.external_attr = info.external_attr
                zout.writestr(zi, data)
