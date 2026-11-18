"""Heuristic detection of input fields in Japanese "Excel 方眼紙" forms.

The idea: an Excel 方眼紙 form is a grid of tiny cells where the author used
*merged ranges* to draw labels and input boxes. We rebuild those logical
"regions" and then pair every label with the empty box next to (or under) it.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Iterable

from .ooxml import Sheet, Workbook, idx_to_col, make_ref

CHECK_CHARS = "□☐"
CHECKED_CHARS = "■☑☒✓✔"

_SPACES = re.compile(r"[\s　]+")
_TRAIL = re.compile(r"[：:＊*※]+$")
_DATE_PLACEHOLDER = re.compile(
    r"^(?P<pre>.*?)(?P<era>令和|平成|昭和|西暦)?[\s　]*(?P<y>\d*)[\s　]*年"
    r"[\s　]*(?P<m>\d*)[\s　]*月[\s　]*(?P<d>\d*)[\s　]*日(?P<post>.*)$",
    re.S,
)
_INLINE_BLANK = re.compile(r"^(?P<label>[^_＿]{1,30}?)[：:]?[\s　]*(?P<blank>[_＿]{2,})(?P<post>.*)$", re.S)
_OPTION = re.compile(r"[□☐■☑☒]\s*([^□☐■☑☒\s　（）()、,，]+)")
_INDEX_HEADER = {"no", "no.", "№", "番号", "項番", "#", "ｎｏ", "ｎｏ．"}


def normalize_label(text: str) -> str:
    t = _SPACES.sub("", str(text))
    t = _TRAIL.sub("", t)
    t = t.strip("（）()【】[]「」")
    return t


@dataclass(frozen=True)
class Region:
    r1: int
    c1: int
    r2: int
    c2: int
    value: object
    merged: bool
    bordered: bool
    formula: bool = False

    @property
    def ref(self) -> str:
        return make_ref(self.r1, self.c1)

    @property
    def range(self) -> str:
        a, b = make_ref(self.r1, self.c1), make_ref(self.r2, self.c2)
        return a if a == b else f"{a}:{b}"

    @property
    def text(self) -> str:
        return "" if self.value is None else str(self.value).strip()

    @property
    def empty(self) -> bool:
        return self.text == "" and not self.formula

    @property
    def width(self) -> int:
        return self.c2 - self.c1 + 1


@dataclass
class Field:
    id: str
    label: str
    sheet: str
    cell: str
    range: str
    kind: str  # text | date | choice | inline_text | inline_date | table
    section: str | None = None
    options: list[str] = field(default_factory=list)
    columns: list[dict] = field(default_factory=list)
    rows: list[int] = field(default_factory=list)
    template: str | None = None  # original cell text for inline kinds
    value: object = None  # current value (only meaningful for filled forms)

    def to_dict(self) -> dict:
        d = {
            "id": self.id,
            "label": self.label,
            "kind": self.kind,
            "sheet": self.sheet,
            "cell": self.cell,
            "range": self.range,
        }
        if self.section:
            d["section"] = self.section
        if self.options:
            d["options"] = self.options
        if self.columns:
            d["columns"] = self.columns
            d["max_rows"] = len(self.rows)
        if self.template is not None:
            d["template"] = self.template
        return d


class Grid:
    def __init__(self, sheet: Sheet, bordered_styles: set[int]):
        self.sheet = sheet
        self.bordered_styles = bordered_styles
        self.owner: dict[tuple[int, int], Region] = {}
        self.max_row = sheet.max_row
        self.max_col = sheet.max_col
        for (r1, c1, r2, c2) in sheet.merges:
            top = sheet.cells.get((r1, c1))
            bordered = any(
                (c := sheet.cells.get((r, cc))) is not None and c.style in bordered_styles
                for r in range(r1, r2 + 1)
                for cc in (c1, c2)
            )
            reg = Region(r1, c1, r2, c2, top.value if top else None, True, bordered,
                         bool(top and top.formula is not None))
            for r in range(r1, r2 + 1):
                for cc in range(c1, c2 + 1):
                    self.owner[(r, cc)] = reg

    def at(self, row: int, col: int) -> Region | None:
        if row < 1 or col < 1 or row > self.max_row or col > self.max_col:
            return None
        reg = self.owner.get((row, col))
        if reg is not None:
            return reg
        c = self.sheet.cells.get((row, col))
        value = c.value if c else None
        bordered = bool(c and c.style in self.bordered_styles)
        return Region(row, col, row, col, value, False, bordered,
                      bool(c and c.formula is not None))

    def at_or_blank(self, row: int, col: int) -> Region:
        reg = self.at(row, col)
        return reg if reg is not None else Region(row, col, row, col, None, False, False)

    def regions(self) -> list[Region]:
        """All non-empty regions plus empty merged/bordered boxes, reading order."""
        seen: set[Region] = set()
        out: list[Region] = []
        keys = set(self.owner) | set(self.sheet.cells)
        for (r, c) in sorted(keys):
            reg = self.at(r, c)
            if reg is None or reg in seen:
                continue
            seen.add(reg)
            if reg.empty and not (reg.merged or reg.bordered):
                continue
            out.append(reg)
        out.sort(key=lambda g: (g.r1, g.c1))
        return out


def is_label(reg: Region) -> bool:
    if reg.empty or reg.formula or not isinstance(reg.value, str):
        return False
    t = reg.text
    if len(t) > 40:
        return False
    if any(ch in t for ch in CHECK_CHARS + CHECKED_CHARS):
        return False
    if _DATE_PLACEHOLDER.match(t) and _is_placeholder_date(t):
        return False
    if re.fullmatch(r"[\d,.\-+%円¥\s]+", t):
        return False
    return True


def _is_placeholder_date(text: str) -> bool:
    m = _DATE_PLACEHOLDER.match(text)
    if not m:
        return False
    seg = text[m.start("y"):m.start("post")]
    # "生年月日" is a label, "令和　年　月　日" is a placeholder
    if not re.search(r"年[\s　\d]+月[\s　\d]+日", seg):
        return False
    return not (m.group("y") and m.group("m") and m.group("d"))


def _is_box(reg: Region | None) -> bool:
    return reg is not None and reg.empty and (reg.merged or reg.bordered)


def _date_like_label(label: str) -> bool:
    return label.endswith("日") or "年月日" in label or label in {"日付", "期日"}


# --------------------------------------------------------------------------
def detect_sheet(sheet: Sheet, bordered_styles: set[int]) -> list[Field]:
    g = Grid(sheet, bordered_styles)
    regions = g.regions()
    claimed: set[Region] = set()  # input boxes already paired
    used_labels: set[Region] = set()
    fields: list[Field] = []

    def left_label(reg: Region) -> Region | None:
        c = reg.c1 - 1
        while c >= 1:
            left = g.at(reg.r1, c)
            if left is None:
                return None
            if not left.empty:
                return left if is_label(left) and left not in used_labels else None
            if left.merged or left.bordered:
                return None
            c = left.c1 - 1
            if reg.c1 - c > 3:
                return None
        return None

    # ---- 1. tables -------------------------------------------------------
    for tf, header_regs, body_regs in _detect_tables(g, regions):
        fields.append(tf)
        used_labels.update(header_regs)
        claimed.update(body_regs)

    # ---- 2. inline kinds (dates / checkboxes / underscores) ---------------
    for reg in regions:
        if reg in used_labels or reg.empty or not isinstance(reg.value, str):
            continue
        t = reg.text
        if any(ch in t for ch in CHECK_CHARS + CHECKED_CHARS):
            opts = _OPTION.findall(t)
            if not opts:
                continue
            first = min((t.find(ch) for ch in CHECK_CHARS + CHECKED_CHARS if ch in t))
            label = t[:first].strip(" 　：:")
            lab_reg = left_label(reg)
            if not label and lab_reg is not None:
                label = lab_reg.text
                used_labels.add(lab_reg)
            fields.append(Field("", label or "選択", sheet.name, reg.ref, reg.range,
                                "choice", options=opts, template=str(reg.value)))
            claimed.add(reg)
        elif _is_placeholder_date(t):
            lab_reg = left_label(reg)
            m = _DATE_PLACEHOLDER.match(t)
            label = (m.group("pre") or "").strip(" 　：:")
            if not label and lab_reg is not None:
                label = lab_reg.text
                used_labels.add(lab_reg)
            fields.append(Field("", label or "日付", sheet.name, reg.ref, reg.range,
                                "inline_date", template=str(reg.value)))
            claimed.add(reg)
        else:
            m = _INLINE_BLANK.match(t)
            if m:
                fields.append(Field("", m.group("label").strip(), sheet.name, reg.ref,
                                    reg.range, "inline_text", template=str(reg.value)))
                claimed.add(reg)
                used_labels.add(reg)

    labels = [r for r in regions if is_label(r) and r not in used_labels and r not in claimed]

    # ---- 3. label -> box on the right ------------------------------------
    for lab in labels:
        box = None
        c = lab.c2 + 1
        skipped = 0
        while c <= g.max_col:
            reg = g.at(lab.r1, c)
            if reg is None or not reg.empty:
                break
            if (reg.merged or reg.bordered) and reg not in claimed:
                box = reg
                break
            if reg in claimed:
                break
            # plain blank grid cell (方眼紙 spacer)
            skipped += 1
            if skipped > 2:
                break
            c = reg.c2 + 1
        if box is None and lab.text.endswith(("：", ":")):
            nxt = g.at_or_blank(lab.r1, lab.c2 + 1)
            if nxt.empty and nxt not in claimed:
                box = nxt
        if box is not None:
            claimed.add(box)
            used_labels.add(lab)
            fields.append(_box_field(sheet, lab, box))

    # ---- 4. label -> box below -------------------------------------------
    for lab in labels:
        if lab in used_labels:
            continue
        below = g.at(lab.r2 + 1, lab.c1)
        if _is_box(below) and below not in claimed and below.c2 >= lab.c2:
            claimed.add(below)
            used_labels.add(lab)
            fields.append(_box_field(sheet, lab, below))

    # ---- 5. sections, ids --------------------------------------------------
    headings = _headings(g, regions, used_labels, claimed)
    for f in fields:
        r = int(re.sub(r"\D", "", f.cell))
        sec = None
        for hr, text in headings:
            if hr <= r:
                sec = text
            else:
                break
        f.section = sec
    fields.sort(key=lambda f: (int(re.sub(r"\D", "", f.cell)), _col_of(f.cell)))
    _assign_ids(fields)
    return fields


def _col_of(ref: str) -> int:
    from .ooxml import col_to_idx

    return col_to_idx(re.sub(r"\d", "", ref))


def _box_field(sheet: Sheet, lab: Region, box: Region) -> Field:
    label = lab.text
    kind = "date" if _date_like_label(normalize_label(label)) else "text"
    return Field("", label, sheet.name, box.ref, box.range, kind)


def _headings(g: Grid, regions: Iterable[Region], used: set[Region], claimed: set[Region]):
    by_row: dict[int, list[Region]] = {}
    for reg in regions:
        if not reg.empty:
            by_row.setdefault(reg.r1, []).append(reg)
    out = []
    for r in sorted(by_row):
        regs = by_row[r]
        if len(regs) != 1:
            continue
        reg = regs[0]
        if reg in used or reg in claimed or not is_label(reg):
            continue
        t = reg.text
        if len(t) > 30:
            continue
        out.append((r, normalize_label(t)))
    return out


def _assign_ids(fields: list[Field]) -> None:
    base = [normalize_label(f.label) or f.cell for f in fields]
    counts: dict[str, int] = {}
    for b in base:
        counts[b] = counts.get(b, 0) + 1
    seen: dict[str, int] = {}
    for f, b in zip(fields, base):
        fid = b
        if counts[b] > 1 and f.section:
            fid = f"{f.section}/{b}"
        n = seen.get(fid, 0) + 1
        seen[fid] = n
        if n > 1:
            fid = f"{fid}#{n}"
        f.id = fid
    # second pass: section-qualified ids could still collide with #1 entries
    ids = [f.id for f in fields]
    for f in fields:
        if ids.count(f.id) > 1:
            pass  # extremely rare; ids stay usable via cell refs


# --------------------------------------------------------------------------
def _detect_tables(g: Grid, regions: list[Region]):
    """Yield (Field, header_regions, body_regions) for grid-like tables."""
    by_row: dict[int, list[Region]] = {}
    for reg in regions:
        by_row.setdefault(reg.r1, []).append(reg)
    consumed_rows: set[int] = set()
    for r in sorted(by_row):
        if r in consumed_rows:
            continue
        row_regs = sorted(by_row[r], key=lambda x: x.c1)
        # find runs of horizontally contiguous label regions
        run: list[Region] = []
        runs: list[list[Region]] = []
        for reg in row_regs:
            if is_label(reg) and (not run or reg.c1 == run[-1].c2 + 1):
                run.append(reg)
            else:
                if len(run) >= 2:
                    runs.append(run)
                run = [reg] if is_label(reg) else []
        if len(run) >= 2:
            runs.append(run)
        for hdr in runs:
            res = _table_body(g, hdr)
            if res is None:
                continue
            body_rows, body_regs = res
            cols = []
            for h in hdr:
                cols.append({"id": normalize_label(h.text), "label": h.text,
                             "col": idx_to_col(h.c1)})
            label = "表"
            above = g.at(hdr[0].r1 - 1, hdr[0].c1)
            if above is not None and is_label(above):
                label = above.text
            first = hdr[0]
            last_r = body_rows[-1]
            f = Field("", label, g.sheet.name, make_ref(body_rows[0], first.c1),
                      f"{make_ref(body_rows[0], first.c1)}:{make_ref(last_r, hdr[-1].c2)}",
                      "table", columns=cols, rows=body_rows)
            consumed_rows.update(range(hdr[0].r1, last_r + 1))
            yield f, set(hdr), body_regs


def _table_body(g: Grid, hdr: list[Region]):
    start = max(h.r2 for h in hdr) + 1
    rows: list[int] = []
    body_regs: set[Region] = set()
    r = start
    while r <= g.max_row and len(rows) < 500:
        regs = [g.at(r, h.c1) for h in hdr]
        if any(x is None for x in regs):
            break
        ok = True
        for i, (h, x) in enumerate(zip(hdr, regs)):
            if x.empty:
                if not (x.merged or x.bordered):
                    ok = False
                    break
                continue
            # allow a pre-numbered index column
            if i == 0 and normalize_label(h.text).lower() in _INDEX_HEADER and re.fullmatch(r"\d+", x.text):
                continue
            ok = False
            break
        if not ok:
            break
        rows.append(r)
        body_regs.update(regs)
        r = max(x.r2 for x in regs) + 1
    if len(rows) < 2:
        return None
    return rows, body_regs


def detect(wb: Workbook, sheet: str | None = None) -> list[Field]:
    sheets = [wb.sheet(sheet)] if sheet else [s for s in wb.sheets if not s.hidden]
    out: list[Field] = []
    for s in sheets:
        out.extend(detect_sheet(s, wb.bordered_styles))
    if len({f.sheet for f in out}) > 1:
        # make ids unique across sheets
        seen: dict[str, int] = {}
        for f in out:
            seen[f.id] = seen.get(f.id, 0) + 1
        for f in out:
            if seen[f.id] > 1:
                f.id = f"{f.sheet}!{f.id}"
    return out
