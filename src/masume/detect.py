"""Heuristic detection of input fields in Japanese "Excel 方眼紙" forms.

An Excel 方眼紙 form is a grid of tiny cells where the author used *merged
ranges* and borders to draw labels and input boxes. We rebuild those logical
"regions" and pair every label with the box next to (or under) it, and we
recognise the conventions Japanese forms print *inside* boxes:

* □ checkboxes (in one cell, or one □ per cell), 男 ・ 女 circle-one choices
* date placeholders 「令和　年　月　日」, era lists 「昭和・平成　年…」, ranges 「～」
* fill-in templates 「〒　　－」「（　）　－」「　時　分」, underscores 「＿＿」
* split boxes 「[ ]年[ ]ヶ月」「[ ]－[ ]－[ ]」, one-digit-per-cell boxes
* repeating tables (明細, 学歴・職歴), honorific boxes 「[ ]御中」
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Iterable

from .ooxml import Sheet, Styles, Workbook, col_to_idx, idx_to_col, make_ref

CHECK_CHARS = "□☐"
CHECKED_CHARS = "■☑☒✓✔"
BOX_CHARS = CHECK_CHARS + CHECKED_CHARS
HEADING_MARKS = "■●◆◇▼▽▲△★☆◎"

_SP = r"[\s　]"
_SPACES = re.compile(_SP + "+")
_ERA = r"(?:明治|大正|昭和|平成|令和|明|大|昭|平|令)"
_DATE_ONE = re.compile(
    rf"(?P<era>西暦|{_ERA}(?:{_SP}*[・･、/／]{_SP}*{_ERA})*)?{_SP}*(?P<y>\d*){_SP}*年"
    rf"(?P<g1>{_SP}*)(?P<m>\d*){_SP}*月(?P<g2>{_SP}*)(?P<d>\d*){_SP}*日"
)
# kept for backwards compatibility (single placeholder, with pre/post text)
_DATE_PLACEHOLDER = re.compile(
    rf"^(?P<pre>.*?)(?P<era>西暦|{_ERA}(?:{_SP}*[・･、/／]{_SP}*{_ERA})*)?{_SP}*(?P<y>\d*){_SP}*年"
    rf"{_SP}*(?P<m>\d*){_SP}*月{_SP}*(?P<d>\d*){_SP}*日(?P<post>.*)$",
    re.S,
)
_INLINE_BLANK = re.compile(r"^(?P<label>[^_＿]{1,30}?)[：:]?[\s　]*(?P<blank>[_＿]{2,})(?P<post>.*)$", re.S)
_OPTION = re.compile(r"[□☐■☑☒✓✔]\s*([^□☐■☑☒✓✔\s　（）()、,，]+)")
_INDEX_HEADER = {"no", "no.", "№", "番号", "項番", "#", "ｎｏ", "ｎｏ．"}
_TOTAL_WORDS = {"合計", "小計", "計", "総計", "以上", "合計金額"}
_EXAMPLE_SHEET = re.compile(r"記入例|見本|サンプル|example|sample", re.I)
_SEP_CHARS = "－-ー‐−―/／:：.．"
_TEMPLATE_CHARS = set("〒－-ー‐−()（）:：時分秒年月日歳満生・～〜/／.．") | set(_SEP_CHARS)
UNIT_TOKENS = {
    "年", "月", "日", "歳", "才", "円", "万円", "千円", "ヶ月", "か月", "カ月", "箇月", "ケ月",
    "時", "分", "秒", "人", "名", "枚", "件", "通", "部", "冊", "台", "回", "号", "番", "番地",
    "丁目", "階", "泊", "日間", "週間", "年度", "%", "％", "kg", "cm", "円/月",
    "－", "-", "ー", "‐", "−", "―", "〒", "～", "〜", "・", "：", ":", "／", "/", "(", ")", "（", "）",
    "様", "御中", "殿", "印", "㊞",
}
HONORIFICS = {"様", "御中", "殿"}
CIRCLE_VOCAB = {
    "男", "女", "男性", "女性", "有", "無", "あり", "なし", "有り", "無し", "要", "不要", "可", "否",
    "はい", "いいえ", "済", "未", "新規", "更新", "変更", "継続", "自宅", "勤務先", "本人", "代理人",
    "明", "大", "昭", "平", "令", "明治", "大正", "昭和", "平成", "令和", "同意する", "同意しない",
    "普通", "当座", "既婚", "未婚", "持家", "借家", "賛成", "反対", "出席", "欠席",
}
_FURIGANA = {"ふりがな", "フリガナ", "よみがな", "ヨミガナ"}
# filled in by the office, not the applicant (受付印, ※市記入欄, 処理欄 ...)
_OFFICE_USE = re.compile(r"受付印?$|受付番号|受理|記入欄$|処理欄|事務処理|整理番号|記入しないで|使用欄|^宛先$|^あて先$|問い?合わ?せ先|^提出先")
# the printed addressee 「広島市長」「〇〇税務署長」 is not a label
_ADDRESSEE = re.compile(r"(?:市|区|町|村|都|道|府|県|署|所|局|庁)長$|知事$")
_ERA_NAMES = {"令和", "平成", "昭和", "西暦"}
# 「⑴ 名称変更」「(2) 所在地変更」: the group a check box belongs to
_GROUP_NO = re.compile(r"^(?:[⑴-⒇①-⑳]|[（(][0-9０-９]{1,2}[)）])")


def normalize_label(text: str) -> str:
    t = _SPACES.sub("", str(text))
    t = t.lstrip("※＊*")
    t = re.sub(r"[：:＊*※]+$", "", t)
    pairs = {"（": "）", "(": ")", "【": "】", "[": "]", "「": "」", "〔": "〕"}
    # strip brackets only when ONE matching pair encloses the whole label
    while len(t) >= 2 and t[0] in pairs and t[-1] == pairs[t[0]] and t.count(t[0]) == 1:
        t = t[1:-1]
    for ch in HEADING_MARKS:
        if t.startswith(ch):
            t = t[1:]
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
    style: int = 0
    sides: frozenset = frozenset()

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

    @property
    def height(self) -> int:
        return self.r2 - self.r1 + 1


@dataclass
class Field:
    id: str
    label: str
    sheet: str
    cell: str
    range: str
    kind: str  # text | date | choice | choice_cells | circle | inline_text | inline_date | pattern | split | digits | table
    section: str | None = None
    options: list[str] = field(default_factory=list)
    columns: list[dict] = field(default_factory=list)
    rows: list[int] = field(default_factory=list)
    parts: list[dict] = field(default_factory=list)  # split / digits / choice_cells
    template: str | None = None  # original cell text for inline kinds
    suffix: str | None = None  # unit printed after the box (円, 歳, 様 ...)
    value: object = None  # current value when the form is already filled
    id_hint: str | None = None

    @classmethod
    def from_dict(cls, d: dict) -> "Field":
        rows = d.get("rows") or []
        if not rows and d.get("kind") == "table" and d.get("max_rows"):
            start = int(re.sub(r"\D", "", d["cell"]))
            rows = list(range(start, start + int(d["max_rows"])))
        return cls(
            id=d["id"], label=d.get("label", d["id"]), sheet=d["sheet"], cell=d["cell"],
            range=d.get("range", d["cell"]), kind=d.get("kind", "text"), section=d.get("section"),
            options=list(d.get("options") or []), columns=list(d.get("columns") or []),
            rows=list(rows), parts=list(d.get("parts") or []), template=d.get("template"),
            suffix=d.get("suffix"), value=d.get("value"),
        )

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
            d["rows"] = self.rows
        if self.parts:
            d["parts"] = self.parts
        if self.template is not None:
            d["template"] = self.template
        if self.suffix:
            d["suffix"] = self.suffix
        if self.value not in (None, ""):
            d["value"] = self.value
        return d


class Grid:
    def __init__(self, sheet: Sheet, styles: Styles | set[int]):
        if not isinstance(styles, Styles):
            styles = Styles(bordered=set(styles))
        self.sheet = sheet
        self.styles = styles
        self.owner: dict[tuple[int, int], Region] = {}
        self.max_row = sheet.max_row
        self.max_col = sheet.max_col
        for (r1, c1, r2, c2) in sheet.merges:
            top = sheet.cells.get((r1, c1))
            sides: set[str] = set()
            for r in range(r1, r2 + 1):
                for cc in range(c1, c2 + 1):
                    c = sheet.cells.get((r, cc))
                    if c is not None:
                        sides |= styles.sides.get(c.style, frozenset())
            reg = Region(r1, c1, r2, c2, top.value if top else None, True, bool(sides),
                         bool(top and top.formula is not None), top.style if top else 0, frozenset(sides))
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
        style = c.style if c else 0
        sides = self.styles.sides.get(style, frozenset())
        return Region(row, col, row, col, value, False, bool(sides),
                      bool(c and c.formula is not None), style, sides)

    def at_or_blank(self, row: int, col: int) -> Region:
        reg = self.at(row, col)
        return reg if reg is not None else Region(row, col, row, col, None, False, False)

    def heading_styled(self, reg: Region) -> bool:
        return reg.style in self.styles.heading

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


# --------------------------------------------------------------------------
# text classifiers
# --------------------------------------------------------------------------
def placeholder_dates(text: str) -> list[re.Match]:
    """Date placeholders (with blank y/m/d) inside *text*."""
    tokens = [t for t in _SPACES.split(str(text).strip()) if t]
    if len(tokens) >= 5 and all(len(t) == 1 for t in tokens):
        return []  # 均等割付 caption 「変　更　年　月　日」, not 「　年　月　日」
    out = []
    for m in _DATE_ONE.finditer(text):
        if not (m.group("g1") or m.group("m")) or not (m.group("g2") or m.group("d")):
            continue  # "生年月日" is a label, not a placeholder
        if m.group("y") and m.group("m") and m.group("d"):
            continue  # already filled
        out.append(m)
    return out


def _is_placeholder_date(text: str) -> bool:
    return bool(placeholder_dates(text))


def is_unit(text: str) -> bool:
    return normalize_label(text) in UNIT_TOKENS or text.strip() in UNIT_TOKENS


def is_symbol_template(text: str) -> bool:
    raw = str(text)
    if not re.search(r"[\s　]{2,}|^[\s　]+\S|[\s　]\S*$", raw):
        return False
    residue = _SPACES.sub("", raw)
    return bool(residue) and len(residue) <= 20 and all(ch in _TEMPLATE_CHARS for ch in residue)


def is_template(text: str) -> bool:
    """「〒　　－」「（　）　－」「　時　分」「銀行・信用金庫　　支店」"""
    raw = str(text)
    if not re.search(r"[\s　]{2,}|^[\s　]+\S", raw):
        return False
    residue = _SPACES.sub("", raw)
    if not residue or len(residue) > 20:
        return False
    if all(ch in _TEMPLATE_CHARS for ch in residue):
        return True
    # words separated by wide blanks, e.g. 「銀行・信用金庫　　　支店」
    tokens = [t for t in _SPACES.split(raw.strip()) if t]
    if all(len(t) == 1 for t in tokens):
        return False  # 均等割付 label such as 「氏　　名」
    if any(re.fullmatch(r"[（(].+[）)]", t) for t in tokens):
        return False  # 「名　　称」+「（氏名）」 is a two-line label
    return len(tokens) >= 2 and all(len(t) <= 8 for t in tokens) and bool(re.search(r"[\s　]{3,}", raw))


def circle_options(text: str) -> list[str] | None:
    """「男 ・ 女」「有・無」 -> options, else None."""
    t = text.strip()
    if any(ch in t for ch in BOX_CHARS) or len(t) > 30 or not re.search(r"[・･/／]", t):
        return None
    parts = [p.strip(" 　") for p in re.split(r"[・･/／]", t)]
    if not 2 <= len(parts) <= 6 or any(not p or len(p) > 6 or _SPACES.search(p) for p in parts):
        return None
    spaced = bool(re.search(r"[\s　][・･/／]|[・･/／][\s　]", t))
    if spaced or all(p in CIRCLE_VOCAB for p in parts):
        return parts
    return None


def choice_options(text: str) -> list[str]:
    return _OPTION.findall(text)


def is_label(reg: Region, maxlen: int = 40) -> bool:
    if reg.empty or reg.formula or not isinstance(reg.value, str):
        return False
    t = reg.text
    if len(t) > maxlen:
        return False
    n_box = sum(t.count(ch) for ch in BOX_CHARS)
    if n_box >= 2 or (n_box == 1 and t[0] not in BOX_CHARS):
        return False
    if n_box == 1 and t[0] in CHECK_CHARS and len(t.strip()) <= 1:
        return False
    if _is_placeholder_date(t) or is_symbol_template(t) or circle_options(t):
        return False
    if re.fullmatch(r"[\d０-９,，.．\-－+＋%％円¥￥\s　]+", t):
        return False
    if is_unit(t):
        return False
    n = normalize_label(t)
    if _OFFICE_USE.search(n) or _ADDRESSEE.search(n):
        return False
    return True


def is_frame(reg: Region | None) -> bool:
    """A single empty cell carrying one rule line of a larger frame (not a box)."""
    return (reg is not None and reg.empty and not reg.merged and bool(reg.sides)
            and reg.sides in ({"top"}, {"left"}, {"right"}, {"left", "right"}))


def has_heading_cue(text: str) -> bool:
    t = text.strip()
    return bool(t) and (t[0] in HEADING_MARKS or t.startswith("【") or bool(re.match(r"^第.{1,3}[章節部]", t))
                        or bool(re.match(r"^[0-9０-９]{1,2}[\s　]+\S", t) and len(_SPACES.sub("", t)) >= 6))
    # 「１　変更理由等」 is a numbered heading; 「１　氏名」 stays a label


def _is_box(reg: Region | None) -> bool:
    return reg is not None and reg.empty and (reg.merged or reg.bordered) and not is_frame(reg)


def _date_like_label(label: str) -> bool:
    return label.endswith("日") or "年月日" in label or label in {"日付", "期日"}


def _row_of(ref: str) -> int:
    return int(re.sub(r"\D", "", ref))


def _col_of(ref: str) -> int:
    return col_to_idx(re.sub(r"\d", "", ref))


# --------------------------------------------------------------------------
def detect_sheet(sheet: Sheet, styles: Styles | set[int]) -> list[Field]:
    g = Grid(sheet, styles)
    regions = g.regions()
    claimed: set[Region] = set()  # input boxes already paired
    used_labels: set[Region] = set()
    fields: list[Field] = []

    def free_label(reg: Region | None) -> bool:
        return (reg is not None and is_label(reg) and reg not in used_labels
                and reg not in claimed and not has_heading_cue(reg.text))

    def left_label(reg: Region, row: int | None = None) -> Region | None:
        row = reg.r1 if row is None else row
        c = reg.c1 - 1
        while c >= 1:
            left = g.at(row, c)
            if left is None:
                return None
            if not left.empty:
                if free_label(left) and not (g.heading_styled(left) and not left.bordered):
                    return left
                return None
            if left.merged or left.bordered:
                return None
            c = left.c1 - 1
            if reg.c1 - c > 3:
                return None
        return None

    def new(label, reg, kind, **kw) -> Field:
        return Field("", label, sheet.name, reg.ref, reg.range, kind, **kw)

    # ---- 0. check boxes drawn as boxed cells: [ ]電子データ [ ]書面 ----------
    fields.extend(_drawn_box_fields(g, sheet, regions, used_labels, claimed))

    # ---- 0a. 「有 （[ ]年[ ]月[ ]日） ・ 無」: options in separate cells ----------
    fields.extend(_circle_cell_fields(g, sheet, regions, used_labels, claimed))

    # ---- 0b. comparison grids: 事項 | 変更前 | 変更後 -----------------------
    fields.extend(_comparison_fields(g, sheet, regions, used_labels, claimed))

    # ---- 1. tables -------------------------------------------------------
    for tf, header_regs, body_regs in _detect_tables(g, regions):
        if any(b in claimed for b in body_regs) or any(h in used_labels for h in header_regs):
            continue  # already read as a comparison grid
        fields.append(tf)
        used_labels.update(header_regs)
        claimed.update(body_regs)

    # ---- 2. things printed inside cells ----------------------------------
    for reg in regions:
        if reg in used_labels or reg in claimed or reg.empty or not isinstance(reg.value, str):
            continue
        t = reg.text
        n_box = sum(t.count(ch) for ch in BOX_CHARS)
        if n_box == 1 and len(t) == 1:
            res = _choice_cells(g, reg, claimed, used_labels)
            if res is None:
                continue
            parts, option_regs = res
            lab = left_label(reg)
            label = lab.text if lab else "選択"
            if lab:
                used_labels.add(lab)
            f = new(label, reg, "choice_cells", options=[p["option"] for p in parts], parts=parts)
            f.range = f"{reg.ref}:{make_ref(option_regs[-1].r2, option_regs[-1].c2)}" if option_regs else reg.range
            f.value = [p["option"] for p in parts if p.get("checked")] or None
            for p in parts:
                p.pop("checked", None)
            fields.append(f)
            claimed.add(reg)
            claimed.update(option_regs)
            continue
        if n_box:
            opts = choice_options(t)
            if not opts:
                continue
            lab = left_label(reg)
            first = min(t.find(ch) for ch in BOX_CHARS if ch in t)
            label = t[:first].strip(" 　：:")
            if len(opts) < 2 and lab is None and not label:
                continue  # 「■ 申込者情報」 is a heading, not a checkbox
            if len(opts) < 2 and "ください" in t:
                continue  # 「該当する項目に☑してください。」 is an instruction
            if not label and lab is not None:
                label = lab.text
            if lab is not None and label == lab.text:
                used_labels.add(lab)
            checked = [o for o in opts if re.search("[" + CHECKED_CHARS + r"]\s*" + re.escape(o), t)]
            fields.append(new(label or "選択", reg, "choice", options=opts, template=str(reg.value),
                              value=checked or None))
            claimed.add(reg)
            continue
        if _is_placeholder_date(t):
            lab = left_label(reg)
            m = _DATE_PLACEHOLDER.match(t)
            label = (m.group("pre") or "").strip(" 　：:") if m else ""
            label = re.sub(_ERA + r"$", "", label).strip(" 　：:・")
            if not label and lab is not None:
                label = lab.text
                used_labels.add(lab)
            if not label:
                label = "記入日" if "現在" in t else "日付"
            fields.append(new(label, reg, "inline_date", template=str(reg.value)))
            claimed.add(reg)
            continue
        lab = left_label(reg)
        opts = circle_options(t) if lab is not None else None
        if opts:
            used_labels.add(lab)
            fields.append(new(lab.text, reg, "circle", options=opts, template=str(reg.value)))
            claimed.add(reg)
            continue
        if lab is not None and (reg.merged or reg.bordered) and is_template(t):
            used_labels.add(lab)
            fields.append(new(lab.text, reg, "pattern", template=str(reg.value)))
            claimed.add(reg)
            continue
        if (lab is None and (reg.merged or reg.bordered) and is_template(t) and not is_symbol_template(t)
                and not _is_box(g.at(reg.r1, reg.c2 + 1))):
            last = [x for x in _SPACES.split(t) if x][-1]
            fields.append(new(last, reg, "pattern", template=str(reg.value)))
            claimed.add(reg)
            used_labels.add(reg)
            continue
        m = _INLINE_BLANK.match(t)
        if m:
            fields.append(new(m.group("label").strip(), reg, "inline_text", template=str(reg.value)))
            claimed.add(reg)
            used_labels.add(reg)

    # ---- 2b. form-control check boxes floating over the grid ---------------
    fields.extend(_control_fields(g, sheet, used_labels, claimed))

    labels = [r for r in regions if free_label(r)]

    # ---- 3. label -> box(es) on the right --------------------------------
    for lab in [r for r in regions if is_label(r) and not has_heading_cue(r.text) and r not in claimed
                and not _is_vertical(r.text)]:
        if lab in used_labels and lab.height == 1:
            continue
        if lab not in used_labels and lab not in labels:
            continue
        found_any = lab in used_labels
        prev_cols = None
        zip_row = False  # a 〒 box was found: the next row is the address itself
        for row in range(lab.r1 + (1 if found_any else 0), lab.r2 + 1):
            res = _scan_right(g, lab, row, claimed, used_labels)
            if res is None:
                continue
            kind, boxes, prefix, suffixes = res
            cols = [(b.c1, b.c2) for b in boxes]
            if row > lab.r1:
                above = g.at(boxes[0].r1 - 1, boxes[0].c1)
                if above is not None and above.r1 >= lab.r1 and is_label(above):
                    continue  # box belongs to the header above it (決裁 | 部長 / stamp)
                if cols == prev_cols and not any("top" in b.sides and "bottom" in b.sides for b in boxes):
                    claimed.update(boxes)
                    continue  # lower half of the same boxes drawn over two rows (法人番号)
            prev_cols = cols
            label = lab.text if not found_any or zip_row else f"{lab.text}({row - lab.r1 + 1})"
            zip_row = False
            if normalize_label(lab.text) in _ERA_NAMES and kind == "text" and suffixes and suffixes[0] in ("年度", "年"):
                # 「令和[ ]年度」: the box is the year, so the field is 「年度」, not 「令和」
                f = _make_box_field(sheet, suffixes[0], kind, boxes, prefix, suffixes)
            elif normalize_label(lab.text) in _ERA_NAMES and kind == "split" and suffixes[0] == "年":
                f = _make_box_field(sheet, "日付", kind, boxes, prefix, suffixes)
                f.template = normalize_label(lab.text)  # 「令和[ ]年[ ]月[ ]日」 split over cells
            else:
                f = _make_box_field(sheet, label, kind, boxes, prefix, suffixes)
                if prefix == "〒" and not found_any and lab.r2 > row:
                    f.id_hint = f"{normalize_label(lab.text)}(〒)"  # 「現住所(〒)」, the address goes in 「現住所」
                    zip_row = True
            fields.append(f)
            claimed.update(boxes)
            found_any = True
        if found_any:
            used_labels.add(lab)
            continue
        if lab.text.endswith(("：", ":")):
            nxt = g.at_or_blank(lab.r1, lab.c2 + 1)
            if nxt.empty and nxt not in claimed:
                claimed.add(nxt)
                used_labels.add(lab)
                fields.append(_make_box_field(sheet, lab.text, "text", [nxt], None, [None]))

    # ---- 3c. one-character boxes in the row under a label: 「指定番号」 -------
    for lab in [r for r in regions if free_label(r)]:
        row = lab.r2 + 1
        edge = Region(row, lab.c1 - 1, row, lab.c1 - 1, None, False, False)
        res = _scan_right(g, edge, row, claimed, used_labels)
        if res is None or res[0] != "digits" or res[1][0].c1 > lab.c1 + 1:
            continue
        kind, boxes, prefix, suffixes = res
        heads = [g.at(row - 1, b.c1) for b in boxes]
        if any(h is not None and h != lab and not h.empty for h in heads):
            continue  # each box has its own header (決裁: 部長 | 課長 | 担当): not one number
        fields.append(_make_box_field(sheet, lab.text, kind, boxes, prefix, suffixes))
        claimed.update(boxes)
        used_labels.add(lab)

    # ---- 3d. 「変更年月日」 with [ ]年[ ]月[ ]日 in the row under it ------------
    for lab in [r for r in regions if free_label(r) and _date_like_label(normalize_label(r.text))]:
        row = lab.r2 + 1
        edge = Region(row, lab.c1 - 1, row, lab.c1 - 1, None, False, False)
        res = _scan_right(g, edge, row, claimed, used_labels, stop_col=lab.c2)
        if res is None or res[0] != "split" or res[3][:2] != ["年", "月"]:
            continue
        kind, boxes, prefix, suffixes = res
        f = _make_box_field(sheet, lab.text, kind, boxes, prefix, suffixes)
        f.template = "日付"  # fill takes a date; the year is 和暦 when the sheet uses 和暦
        fields.append(f)
        claimed.update(boxes)
        used_labels.add(lab)

    # ---- 3b. [box] 御中 / 様 / 殿 ------------------------------------------
    for reg in regions:
        if reg.empty or not isinstance(reg.value, str) or reg.text not in HONORIFICS:
            continue
        left = g.at(reg.r1, reg.c1 - 1)
        if _is_box(left) and left not in claimed:
            claimed.add(left)
            fields.append(_make_box_field(sheet, "宛名", "text", [left], None, [reg.text]))

    # ---- 4. label -> box below -------------------------------------------
    long_labels = [r for r in regions if r not in used_labels and r not in claimed
                   and is_label(r, maxlen=120) and not has_heading_cue(r.text)]
    for lab in long_labels:
        if lab in used_labels:
            continue
        below = g.at(lab.r2 + 1, lab.c1)
        if (_is_box(below) and below not in claimed and below.c1 == lab.c1
                and below.c2 >= lab.c2):
            claimed.add(below)
            used_labels.add(lab)
            fields.append(_make_box_field(sheet, lab.text, "text", [below], None, [None]))

    # ---- 5. filled forms: label -> non-empty box on the right --------------
    for lab in labels:
        if lab in used_labels or lab in claimed or _is_sentence(lab.text):
            continue
        right = g.at(lab.r1, lab.c2 + 1)
        if (right is not None and right.text in _BOX_PREFIXES and right.merged and right not in claimed
                and right not in used_labels):
            # an address box with 〒 printed in its corner: the value is written after the mark
            claimed.add(right)
            used_labels.add(lab)
            f = _make_box_field(sheet, lab.text, "text", [right], None, [None])
            f.template = right.text
            fields.append(f)
            continue
        if right is not None and isinstance(right.value, str) and _is_printed(right.text):
            continue
        if (right is None or right.empty or right in claimed or right in used_labels
                or not (right.merged or right.bordered) or not isinstance(right.value, (str, int, float))
                or g.heading_styled(right) or has_heading_cue(str(right.value))
                or re.fullmatch(r"[（(][^（(]+[）)]", str(right.value).strip())):
            continue  # 「申請者」+「（特別徴収義務者）」 is one label, not a filled value
        if isinstance(right.value, str) and is_label(right):
            after = g.at(right.r1, right.c2 + 1)
            if _is_box(after) and after not in claimed:
                continue  # right is itself a label with its own box
        claimed.add(right)
        used_labels.add(lab)
        f = _make_box_field(sheet, lab.text, "text", [right], None, [None])
        f.value = right.value
        fields.append(f)

    # ---- 5b. label inside a ruled frame, value written in the open space right of it ----
    for lab in labels:
        if lab in used_labels or lab in claimed:
            continue
        n = normalize_label(lab.text)
        if not 2 <= len(n) <= 10 or re.search(r"[。、，,]", n) or g.heading_styled(lab):
            continue
        under = None
        if not lab.merged:
            # an unmerged caption spills over the cells to its right, so its value goes under it;
            # only inside a ruled frame (a rule right above it), never for free-floating text
            above = g.at(lab.r1 - 1, lab.c1)
            ruled = "top" in lab.sides or (above is not None and "bottom" in above.sides)
            under = g.at(lab.r2 + 1, lab.c1)
            if (not ruled or under is None or not under.empty or under.merged or under in claimed
                    or "top" in under.sides or "bottom" in lab.sides):
                continue
        upper = g.at(lab.r1 - 1, lab.c1)
        if (lab.text.strip()[:1] in "（(" and upper is not None and upper in used_labels
                and (upper.c1, upper.c2) == (lab.c1, lab.c2)):
            used_labels.add(lab)  # 「所在地」+「（住所）」 on the next line: one label
            for f in fields:
                if f.label == upper.text:
                    f.label = upper.text + lab.text
            continue
        c, free = lab.c2 + 1, 0
        while c <= g.max_col and free < 4:
            reg = g.at(lab.r1, c)
            if reg is None or reg.merged or not reg.empty or reg in claimed or (
                    reg.bordered and not is_frame(reg)):
                break
            free += 1
            c = reg.c2 + 1
        if free >= 4:
            phone = _phone_template(g, lab, claimed) if _PHONE_LABEL.search(n) else None
            if phone is not None:
                fields.append(phone)  # 「電話」 with 「（　）　－」 printed in the rows under it
                used_labels.add(lab)
                continue
            box = under if under is not None else g.at(lab.r1, lab.c2 + 1)
            claimed.add(box)
            used_labels.add(lab)
            fields.append(_make_box_field(sheet, lab.text, "text", [box], None, [None]))

    # ---- 6. furigana ids ----------------------------------------------------
    for f in fields:
        if normalize_label(f.label) in _FURIGANA and not f.id_hint:
            lab_reg = next((r for r in regions if r.text == f.label and r.r1 == _row_of(f.cell)), None)
            if lab_reg is not None:
                below = g.at(lab_reg.r2 + 1, lab_reg.c1)
                if below is not None and is_label(below):
                    f.id_hint = f"{normalize_label(below.text)}(ふりがな)"

    # ---- 6b. 〒 box + address row under one label: 「現住所(〒)」「現住所」 ------
    by_label = {f.label: f for f in fields}
    for f in fields:
        if f.kind == "pattern" and (f.template or "").lstrip(" 　").startswith("〒"):
            second = by_label.get(f"{f.label}(2)")
            base = f.label
            f.id_hint = f"{normalize_label(base)}(〒)"
            if second is not None:
                second.label = base

    # ---- 7. sections, ids ---------------------------------------------------
    headings = _headings(g, regions, used_labels, claimed)
    groups = _group_labels(g, regions, used_labels, claimed)
    for f in fields:
        r, c = _row_of(f.cell), _col_of(f.cell)
        sec = None
        for hr, text in headings:
            if hr <= r:
                sec = text
            else:
                break
        for grp, text in groups:
            if grp.r1 <= r <= grp.r2 and grp.c2 < c:
                sec = text
        f.section = sec
    fields.sort(key=lambda f: (_row_of(f.cell), _col_of(f.cell)))
    _assign_ids(fields)
    return fields


def _scan_right(g: Grid, lab: Region, row: int, claimed: set, used: set, stop_col: int | None = None):
    """Find box(es) right of *lab* on *row*.

    Returns (kind, boxes, prefix, suffixes) or None. Handles 〒 prefixes,
    unit suffixes (円, 歳), split boxes ([ ]年[ ]ヶ月, [ ]－[ ]－[ ]) and
    one-character-per-cell boxes (口座番号).
    """
    c = lab.c2 + 1
    prefix = None
    boxes: list[Region] = []
    suffixes: list[str | None] = []
    spacer = 0
    last_col = g.max_col if stop_col is None else min(stop_col, g.max_col)
    while c <= last_col:
        reg = g.at(row, c)
        if reg is None or reg in claimed or reg.c2 > last_col:
            break
        if reg.r1 != row and not (reg.r1 <= row <= reg.r2):
            break
        if reg.empty:
            if (is_frame(reg) and "right" in reg.sides and boxes and not boxes[-1].merged
                    and reg.c1 == boxes[-1].c2 + 1
                    and suffixes[-1] is None and "right" in boxes[-1].sides):
                boxes.append(reg)  # last digit box: its other sides come from neighbours
                suffixes.append(None)
                break
            if (reg.merged or reg.bordered) and not is_frame(reg):
                if boxes and suffixes[-1] is None and not reg.merged and not boxes[-1].merged:
                    # adjacent single cells: digits or underline run
                    boxes.append(reg)
                    suffixes.append(None)
                    c = reg.c2 + 1
                    continue
                if (boxes and suffixes[-1] is None and reg.merged and boxes[-1].merged
                        and reg.width == boxes[-1].width <= 6 and (reg.r1, reg.r2) == (boxes[-1].r1, boxes[-1].r2)):
                    boxes.append(reg)  # [ ][ ][ ] merged per character
                    suffixes.append(None)
                    c = reg.c2 + 1
                    continue
                if (boxes and suffixes[-1] is None and reg.merged and boxes[-1].merged
                        and "left" not in reg.sides and "right" not in boxes[-1].sides
                        and (reg.r1, reg.r2) == (boxes[-1].r1, boxes[-1].r2)):
                    c = reg.c2 + 1  # no line between them: the same box drawn as two merged cells
                    continue
                if boxes and suffixes[-1] is None:
                    break  # two different boxes with nothing between: separate fields
                boxes.append(reg)
                suffixes.append(None)
                spacer = 0
                c = reg.c2 + 1
                continue
            spacer += 1
            if spacer > 2:
                break
            c = reg.c2 + 1
            continue
        # non-empty region
        if isinstance(reg.value, str) and is_unit(reg.text) and reg.text not in HONORIFICS or (
                isinstance(reg.value, str) and reg.text in HONORIFICS and boxes):
            if not boxes:
                if prefix is None and len(reg.text) <= 2:
                    prefix = reg.text
                    c = reg.c2 + 1
                    continue
                break
            if suffixes[-1] is None:
                suffixes[-1] = reg.text
                c = reg.c2 + 1
                spacer = 0
                continue
            break
        if (boxes and suffixes[-1] is None and isinstance(reg.value, str)
                and reg.text[:1] in "年月日" and len(boxes) >= 1 and any(suffixes)):
            suffixes[-1] = reg.text[:1]  # 「[ ]日提出」「[ ]月分（…）」
            used.add(reg)
        break
    if not boxes:
        return None
    # trim trailing boxes that are not followed by a unit when we have a split
    single_cells = [b for b in boxes if not b.merged]
    if (len(boxes) >= 2 and not single_cells and all(s is None for s in suffixes)
            and len({(b.width, b.r1, b.r2) for b in boxes}) == 1):
        # [ ][ ][ ] one merged box per character (法人番号); few or wide boxes: keep the first only
        if len(boxes) >= (3 if boxes[0].width <= 3 else 8):
            return "digits", boxes, prefix, suffixes
        boxes, suffixes = boxes[:1], suffixes[:1]
    if len(boxes) >= 3 and len(single_cells) == len(boxes) and all(s is None for s in suffixes):
        vertical = sum(1 for b in boxes if {"left", "right"} & b.sides)
        if vertical >= len(boxes) - 1:
            return "digits", boxes, prefix, suffixes
        # underline drawn on unmerged cells -> one logical box
        merged = Region(boxes[0].r1, boxes[0].c1, boxes[-1].r2, boxes[-1].c2, None, False, True)
        return "underline", [merged] + boxes[1:], prefix, [None] * len(boxes)
    if len(boxes) >= 2:
        return "split", boxes, prefix, suffixes
    b = boxes[0]
    if not b.merged and b.sides != {"bottom"} and len(b.sides) < 3:
        return None  # a corner of a frame (┘ ┐), not a box; an underline 「＿」 is fine
    return "text", boxes, prefix, suffixes


_BOX_PREFIXES = {"〒"}


def _is_vertical(t: str) -> bool:
    """Vertical caption 「給/与/支/払/者」 (one character per line): a group heading, not a field label."""
    lines = [x for x in str(t).split("\n") if x.strip()]
    return len(lines) >= 3 and all(len(x.strip()) <= 1 for x in lines)


def _is_sentence(t: str) -> bool:
    """Instructions and notes 「…してください。」「←…記載」, not field labels."""
    t = t.strip()
    return t.endswith(("。", "、")) or "ください" in t or t.startswith(("←", "→", "※")) or len(t) > 40


def _is_printed(t: str) -> bool:
    """Text the form itself prints: 均等割付 「一　括　徴　収」, multi-line captions, sentences."""
    if "\n" in t or _is_sentence(t):
        return True
    tokens = [x for x in _SPACES.split(t.strip()) if x]
    return len(tokens) >= 3 and all(len(x) == 1 for x in tokens)


def _comparison_fields(g: Grid, sheet: Sheet, regions: list[Region], used: set, claimed: set) -> list[Field]:
    """「事項 | 変更前 | 変更後」: row labels on the left, 2-4 wide column headers on top.

    Every row label gets one field per column (id 「名称（氏名）/変更後」), read with
    the same box scanner as ordinary rows but kept inside the column.
    """
    out: list[Field] = []
    by_row: dict[int, list[Region]] = {}
    for reg in regions:
        by_row.setdefault(reg.r1, []).append(reg)
    for r, regs in sorted(by_row.items()):
        heads = sorted((x for x in regs if is_label(x) and x.merged and x.width >= 8
                        and x not in used and x not in claimed), key=lambda x: x.c1)
        run: list[Region] = []
        for h in heads:
            if run and (h.c1 != run[-1].c2 + 1 or (h.r1, h.r2) != (run[0].r1, run[0].r2)):
                if len(run) >= 2:
                    break
                run = []
            run.append(h)
        if not 2 <= len(run) <= 4:
            continue
        widths = [h.width for h in run]
        if max(widths) > 1.5 * min(widths):
            continue
        first = run[0]

        def row_label(row: int) -> Region | None:
            c = first.c1 - 1
            while c >= 1:
                reg = g.at(row, c)
                if reg is None:
                    return None
                if not reg.empty:
                    return reg if is_label(reg) and reg.r1 <= row else None
                c = reg.c1 - 1
            return None

        body: list[tuple[Region, list]] = []
        taken: set[Region] = set()
        row = first.r2 + 1
        while row <= g.max_row:
            lab = row_label(row)
            if lab is None or lab in claimed:
                break
            cols = []
            for rr in range(lab.r1, lab.r2 + 1):
                for h in run:
                    edge = Region(rr, h.c1 - 1, rr, h.c1 - 1, None, False, False)
                    res = _scan_right(g, edge, rr, claimed, used, stop_col=h.c2)
                    if res is None:
                        continue
                    kind, boxes, prefix, suffixes = res
                    keep = [i for i, b in enumerate(boxes) if b.c2 <= h.c2]
                    if not keep:
                        continue
                    boxes = [boxes[i] for i in keep]
                    suffixes = [suffixes[i] for i in keep]
                    if kind in ("split", "digits") and len(boxes) == 1:
                        kind = "text"
                    if any(b in taken for b in boxes):
                        continue  # a box spanning two rows of a tall label
                    taken.update(boxes)
                    cols.append((rr, h, kind, boxes, prefix, suffixes))
            if not cols:
                break
            body.append((lab, cols))
            row = lab.r2 + 1
        if len(body) < 2:
            continue
        labels_in_order = [lab for lab, _ in body]
        for i, (lab, cols) in enumerate(body):
            base = normalize_label(lab.text)
            if base in _FURIGANA and i + 1 < len(labels_in_order):
                base = f"{normalize_label(labels_in_order[i + 1].text)}(ふりがな)"
            rows_seen: dict[str, int] = {}
            for rr, h, kind, boxes, prefix, suffixes in cols:
                col = normalize_label(h.text)
                if prefix == "〒":
                    name = base + "(〒)"  # the address itself goes in the row below: 「所在地」
                else:
                    n = rows_seen.get(col, 0) + 1
                    rows_seen[col] = n
                    name = base + (f"({n})" if n > 1 else "")
                f = _make_box_field(sheet, lab.text, kind, boxes, None, suffixes)
                if prefix and prefix != "〒":
                    f.template = prefix  # 「（ ） -」 printed before the first box
                f.id_hint = f"{name}/{col}"
                out.append(f)
                claimed.update(boxes)
            used.add(lab)
            claimed.add(lab)  # no second pass over its other rows
        used.update(run)
        corner = g.at(first.r1, first.c1 - 1)
        while corner is not None and corner.empty and corner.c1 > 1:
            corner = g.at(first.r1, corner.c1 - 1)
        if corner is not None and is_label(corner):
            used.add(corner)  # 「事項」
    return out


_ALL_SIDES = frozenset({"left", "right", "top", "bottom"})
_TICKS = set(CHECKED_CHARS) | {"レ", "○", "〇"}


_SEPARATORS = {"・", "･", "／", "/"}
_PHONE_LABEL = re.compile(r"電話|ＴＥＬ|TEL|Tel|ファクス|ＦＡＸ|FAX")
_DASHES = set("－-ー‐−―")


def _phone_template(g: Grid, lab: Region, claimed: set) -> "Field | None":
    """「電話」 with 「（　　）　　－」 printed in separate cells in the rows just under it."""
    for r in range(lab.r2 + 1, lab.r2 + 4):
        marks = []
        c = lab.c1
        while c <= min(g.max_col, lab.c1 + 40):
            reg = g.at(r, c)
            if reg is None:
                break
            t = reg.text if isinstance(reg.value, str) else ""
            if t in ("（", "("):
                marks = [reg]
            elif marks and t in ("）", ")") and len(marks) == 1:
                marks.append(reg)
            elif len(marks) == 2 and t in _DASHES:
                marks.append(reg)
                break
            elif not reg.empty:
                marks = []
            c = reg.c2 + 1
        if len(marks) != 3:
            continue
        slots = []
        for m in marks:
            s = g.at(r, m.c2 + 1)
            if s is None or not s.empty or s in claimed:
                break
            slots.append(s)
        if len(slots) != 3:
            continue
        parts = [{"cell": s.ref, "suffix": suf} for s, suf in zip(slots, ("）", "－", None))]
        parts[-1].pop("suffix")
        f = Field("", lab.text, g.sheet.name, slots[0].ref, f"{marks[0].ref}:{make_ref(r, min(g.max_col, slots[2].c2 + 4))}",
                  "split", parts=parts, template="（")
        claimed.update(slots)
        return f
    return None


def _circle_cell_fields(g: Grid, sheet: Sheet, regions: list[Region], used: set, claimed: set) -> list[Field]:
    """「有 （[ ]年[ ]月[ ]日） ・ 無」: circle-one options printed in their own cells.

    One choice field (the chosen option is written as 「（無）」) plus, when an option
    is followed by date boxes, a date field 「…(有の場合の年月日)」.
    """
    out: list[Field] = []
    by_row: dict[int, list[Region]] = {}
    for reg in regions:
        if not reg.empty and isinstance(reg.value, str) and reg not in claimed:
            by_row.setdefault(reg.r1, []).append(reg)
    for r, regs in sorted(by_row.items()):
        regs.sort(key=lambda x: x.c1)
        opts = [x for x in regs if normalize_label(x.text) in CIRCLE_VOCAB and x not in used]
        seps = [x for x in regs if x.text.strip() in _SEPARATORS]
        chain: list[Region] = []
        for a in opts:
            if chain and not any(chain[-1].c2 < s.c1 and s.c2 < a.c1 for s in seps):
                chain = []
            chain.append(a)
        if len(chain) < 2:
            continue
        lab = None
        c = chain[0].c1 - 1
        while c >= 1:
            reg = g.at(r, c)
            if reg is None:
                break
            if not reg.empty:
                lab = reg if is_label(reg, maxlen=120) else None
                break
            c = reg.c1 - 1
        if lab is None:
            continue
        base = normalize_label(lab.text.split("\n")[0])
        names = [normalize_label(x.text) for x in chain]
        parts = [{"cell": x.ref, "option": n, "circle": True} for x, n in zip(chain, names)]
        f = Field("", lab.text, sheet.name, chain[0].ref, f"{chain[0].ref}:{make_ref(chain[-1].r2, chain[-1].c2)}",
                  "choice_cells", options=names, parts=parts)
        f.id_hint = base
        out.append(f)
        used.add(lab)
        claimed.add(lab)
        for x in chain:
            used.add(x)
            claimed.add(x)
        for s in seps:
            used.add(s)
        for i, x in enumerate(chain):
            stop = chain[i + 1].c1 - 1 if i + 1 < len(chain) else None
            res = _scan_right(g, x, r, claimed, used, stop_col=stop)
            if res is None or res[0] != "split" or res[3][:2] != ["年", "月"]:
                continue
            kind, boxes, prefix, suffixes = res
            d = _make_box_field(sheet, lab.text, kind, boxes, None, suffixes)
            d.template = "日付"  # 「有 （[ ]年[ ]月[ ]日）」: the date that goes with 有
            d.id_hint = f"{base}({names[i]}の場合の年月日)"
            out.append(d)
            claimed.update(boxes)
    return out


def _drawn_box_fields(g: Grid, sheet: Sheet, regions: list[Region], used: set, claimed: set) -> list[Field]:
    """[□] drawn as a small boxed cell, its caption to the right: 「[ ] 電子データ  [ ] 書面」.

    Two or more on one row make one choice field; a repeated caption starts a new
    group (変更前 / 変更後), named after the column header above when there are several.
    """
    pairs = []
    for box in regions:
        if (not box.merged or box in claimed or box.width > 6 or box.height > 4 or box.sides != _ALL_SIDES
                or not (box.empty or box.text in _TICKS)):
            continue
        cap = None
        c = box.c2 + 1
        while c <= box.c2 + 3:
            reg = g.at(box.r1, c)
            if reg is None:
                break
            if not reg.empty:
                if (isinstance(reg.value, str) and reg.r1 <= box.r1 and reg.r2 >= box.r2 and is_label(reg)
                        and len(normalize_label(reg.text)) <= 12 and reg not in claimed):
                    cap = reg
                break
            if reg.merged or (reg.bordered and not is_frame(reg)):
                break
            c = reg.c2 + 1
        if cap is not None:
            pairs.append((box, cap))
    by_row: dict[int, list] = {}
    for box, cap in pairs:
        by_row.setdefault(box.r1, []).append((box, cap))
    out: list[Field] = []
    for r, items in sorted(by_row.items()):
        items.sort(key=lambda x: x[0].c1)
        groups: list[list] = [[]]
        for box, cap in items:
            if any(normalize_label(cap.text) == normalize_label(c2.text) for _, c2 in groups[-1]):
                groups.append([])
            groups[-1].append((box, cap))
        groups = [gr for gr in groups if len(gr) >= 2]
        if not groups:
            continue
        first_box = groups[0][0][0]
        row_lab = None
        c = first_box.c1 - 1
        while c >= 1:
            reg = g.at(r, c)
            if reg is None:
                break
            if not reg.empty:
                row_lab = reg if is_label(reg) else None
                break
            c = reg.c1 - 1
        base = normalize_label(row_lab.text) if row_lab is not None else "選択"
        for gr in groups:
            name = base
            if len(groups) > 1 and row_lab is not None:
                head = None
                for rr in range(row_lab.r1 - 1, max(row_lab.r1 - 20, 0), -1):
                    reg = g.at(rr, gr[0][0].c1)
                    if reg is not None and not reg.empty and isinstance(reg.value, str):
                        head = reg
                        break
                if head is not None:
                    used.add(head)  # a column header, never a filled value
                    h = normalize_label(head.text).split("※")[0]
                    h = re.sub(r"[（(][^（(）)]*[）)]$", "", h) or h
                    name = f"{base}/{h}"
            parts = [{"cell": b.ref, "option": normalize_label(cp.text), "box": True} for b, cp in gr]
            last = gr[-1][1]
            f = Field("", row_lab.text if row_lab is not None else "選択", sheet.name, gr[0][0].ref,
                      f"{gr[0][0].ref}:{make_ref(last.r2, last.c2)}", "choice_cells",
                      options=[p["option"] for p in parts], parts=parts,
                      value=[normalize_label(cp.text) for b, cp in gr if b.text in _TICKS] or None)
            f.id_hint = name
            out.append(f)
            for b, cp in gr:
                claimed.add(b)
                claimed.add(cp)
                used.add(cp)
        if row_lab is not None:
            used.add(row_lab)
            claimed.add(row_lab)  # no second pass over its other rows
    return out


def _control_fields(g: Grid, sheet: Sheet, used: set, claimed: set) -> list[Field]:
    """One field per form-control check box; its caption is the text right of it."""
    out: list[Field] = []
    captions: set[Region] = set()
    boxes = [c for c in sheet.controls if c.kind == "CheckBox"]
    for ctl in boxes:
        row = ctl.row
        lab = None
        for c in range(ctl.c1, ctl.c2 + 6):
            reg = g.at(row, c)
            if reg is not None and not reg.empty and isinstance(reg.value, str) and reg.c1 >= ctl.c1:
                lab = reg
                break
        label = ctl.text or (lab.text if lab is not None else "")
        if lab is not None and not ctl.text:
            captions.add(lab)
            nxt = g.at(lab.r2 + 1, lab.c1)
            if (label.endswith("、") and nxt is not None and not nxt.empty and isinstance(nxt.value, str)
                    and not any(b.row == nxt.r1 and abs(b.c1 - ctl.c1) <= 2 for b in boxes)):
                label += nxt.text  # 「旧社名の法人は登記上存続し、」+「社名変更」
                captions.add(nxt)
        label = _SPACES.sub("", label) or "チェック"
        f = Field("", label, sheet.name, make_ref(row, ctl.c1),
                  f"{make_ref(ctl.r1, ctl.c1)}:{make_ref(ctl.r2, ctl.c2)}", "checkbox",
                  options=[label], parts=[{"control": ctl.shape_id}], value=True if ctl.checked else None)
        out.append(f)
    for f, ctl in zip(out, boxes):
        # the numbered group heading above it: 「⑵ 所在地変更」
        for r in range(ctl.row - 1, max(ctl.row - 9, 0), -1):
            hit = None
            for c in range(max(ctl.c1 - 4, 1), ctl.c1 + 2):
                reg = g.at(r, c)
                if (reg is not None and reg.r1 == r and not reg.empty and isinstance(reg.value, str)
                        and _GROUP_NO.match(normalize_label(reg.text))):
                    hit = reg
                    break
            if hit is not None:
                f.id_hint = _GROUP_NO.sub("", normalize_label(hit.text)) + "/" + normalize_label(f.label)
                used.add(hit)
                break
    used.update(captions)
    claimed.update(captions)
    return out


def _make_box_field(sheet: Sheet, label: str, kind: str, boxes: list[Region], prefix, suffixes) -> Field:
    first = boxes[0]
    lab_norm = normalize_label(label)
    if kind == "underline":
        f = Field("", label, sheet.name, first.ref, first.range, "text")
    elif kind in ("split", "digits"):
        last = boxes[-1]
        parts = []
        for b, s in zip(boxes, suffixes):
            p = {"cell": b.ref}
            if s:
                p["suffix"] = s
            parts.append(p)
        f = Field("", label, sheet.name, first.ref, f"{first.ref}:{make_ref(last.r2, last.c2)}", kind,
                  parts=parts)
    else:
        kind2 = "date" if _date_like_label(lab_norm) else "text"
        f = Field("", label, sheet.name, first.ref, first.range, kind2, suffix=suffixes[0] if suffixes else None)
    return f


def _choice_cells(g: Grid, start: Region, claimed: set, used: set):
    """[□][男性][□][女性] -> parts."""
    parts = []
    option_regs = []
    row = start.r1
    c = start.c1
    while True:
        box = g.at(row, c)
        gap = 0
        while box is not None and box.empty and gap < 2:
            gap += 1
            c = box.c2 + 1
            box = g.at(row, c)
        if box is None or box.empty:
            break
        t = box.text
        if not (len(t) == 1 and t in BOX_CHARS):
            break
        nxt = g.at(row, box.c2 + 1)
        if nxt is None or nxt.empty or not isinstance(nxt.value, str) or any(ch in nxt.text for ch in BOX_CHARS):
            break
        parts.append({"cell": box.ref, "option": nxt.text.strip(), "checked": t in CHECKED_CHARS})
        option_regs.append(nxt)
        claimed.add(box)
        c = nxt.c2 + 1
    if not parts:
        return None
    return parts, option_regs


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
        if reg in used or reg in claimed or not isinstance(reg.value, str):
            continue
        t = reg.text
        if len(t) > 30:
            continue
        if not (has_heading_cue(t) or g.heading_styled(reg)):
            continue
        out.append((r, normalize_label(t)))
    return out


def _group_labels(g: Grid, regions: Iterable[Region], used: set[Region], claimed: set[Region]):
    """Tall, narrow labels on the left (縦書き 請求者 / 必要な方) act as sections."""
    out = []
    for reg in regions:
        if reg in used or reg in claimed or not is_label(reg) or reg.height < 3 or reg.width > 4 or _is_sentence(reg.text):
            continue
        out.append((reg, normalize_label(reg.text)))
    return out


def _assign_ids(fields: list[Field]) -> None:
    base = [f.id_hint or normalize_label(f.label) or f.cell for f in fields]
    counts: dict[str, int] = {}
    sections: dict[str, set] = {}
    for f, b in zip(fields, base):
        counts[b] = counts.get(b, 0) + 1
        sections.setdefault(b, set()).add(f.section)
    seen: dict[str, int] = {}
    for f, b in zip(fields, base):
        fid = b
        if counts[b] > 1 and f.section and len(sections[b]) > 1:
            fid = f"{f.section}/{b}"  # only when the section tells the duplicates apart
        n = seen.get(fid, 0) + 1
        seen[fid] = n
        if n > 1:
            fid = f"{fid}#{n}"
        f.id = fid


# --------------------------------------------------------------------------
def _detect_tables(g: Grid, regions: list[Region]):
    """Yield (Field, header_regions, body_regions) for grid-like tables."""
    by_row: dict[int, list[Region]] = {}
    for reg in regions:
        by_row.setdefault(reg.r1, []).append(reg)
    consumed_rows: set[int] = set()
    used_names: set[str] = set()
    for r in sorted(by_row):
        if r in consumed_rows:
            continue
        row_regs = sorted(by_row[r], key=lambda x: x.c1)
        run: list[Region] = []
        runs: list[list[Region]] = []
        for reg in row_regs:
            if _is_header(reg) and (not run or (reg.c1 == run[-1].c2 + 1 and reg.r2 == run[0].r2)):
                run.append(reg)
            else:
                if len(run) >= 2:
                    runs.append(run)
                run = [reg] if _is_header(reg) else []
        if len(run) >= 2:
            runs.append(run)
        for hdr in runs:
            res = _table_body(g, hdr)
            if res is None:
                continue
            body_rows, body_regs = res
            cols = [{"id": normalize_label(h.text), "label": h.text, "col": idx_to_col(h.c1)} for h in hdr]
            label = None
            above = g.at(hdr[0].r1 - 1, hdr[0].c1)
            span = hdr[-1].c2 - hdr[0].c1 + 1
            if above is not None and is_label(above) and above.width * 2 >= span:
                label = above.text
            if label is None:
                label = max((h.text for h in hdr), key=len)
            if normalize_label(label) in used_names:
                label = f"{label}#{len(used_names) + 1}"
            used_names.add(normalize_label(label))
            first = hdr[0]
            last_r = body_rows[-1]
            f = Field("", label, g.sheet.name, make_ref(body_rows[0], first.c1),
                      f"{make_ref(body_rows[0], first.c1)}:{make_ref(last_r, hdr[-1].c2)}",
                      "table", columns=cols, rows=body_rows)
            consumed_rows.update(range(hdr[0].r1, last_r + 1))
            yield f, set(hdr), body_regs


def _is_header(reg: Region) -> bool:
    if reg.empty or reg.formula or not isinstance(reg.value, str):
        return False
    t = reg.text
    if len(t) > 20 or any(ch in t for ch in BOX_CHARS) or _is_placeholder_date(t) or is_symbol_template(t):
        return False
    if circle_options(t) or re.fullmatch(r"[\d０-９,，.．\-－+＋%％\s　]+", t):
        return False
    return t not in HONORIFICS and t not in ("－", "-", "〒", "～", "・")


def _table_body(g: Grid, hdr: list[Region]):
    start = max(h.r2 for h in hdr) + 1
    rows: list[int] = []
    body_regs: set[Region] = set()
    empty_rows = 0
    r = start
    while r <= g.max_row and len(rows) < 500:
        regs = [g.at(r, h.c1) for h in hdr]
        if any(x is None for x in regs):
            break
        if any(x.r1 != r or x.c1 != h.c1 for x, h in zip(regs, hdr)):
            break
        ok = True
        for i, (h, x) in enumerate(zip(hdr, regs)):
            if x.formula:
                ok = False
                break
            if x.empty:
                if not (x.merged or x.bordered):
                    ok = False
                    break
                continue
            if i == 0 and normalize_label(h.text).lower() in _INDEX_HEADER and re.fullmatch(r"\d+", x.text):
                continue
            # filled cells are allowed inside bordered tables, totals rows end it
            if (x.merged or x.bordered) and normalize_label(x.text) not in _TOTAL_WORDS:
                continue
            ok = False
            break
        if not ok:
            break
        rows.append(r)
        body_regs.update(regs)
        if all(x.empty or (i == 0 and normalize_label(hdr[0].text).lower() in _INDEX_HEADER)
               for i, x in enumerate(regs)):
            empty_rows += 1
        r = max(x.r2 for x in regs) + 1
    # a real table has room left: at least one completely empty row
    if len(rows) < 2 or empty_rows == 0:
        return None
    return rows, body_regs


def is_example_sheet(name: str) -> bool:
    return bool(_EXAMPLE_SHEET.search(name))


def form_sheets(wb: Workbook) -> list[Sheet]:
    visible = [s for s in wb.sheets if not s.hidden]
    main = [s for s in visible if not is_example_sheet(s.name)]
    return main or visible or wb.sheets[:1]


def detect(wb: Workbook, sheet: str | None = None) -> list[Field]:
    sheets = [wb.sheet(sheet)] if sheet else form_sheets(wb)
    out: list[Field] = []
    for s in sheets:
        out.extend(detect_sheet(s, wb.styles))
    if len({f.sheet for f in out}) > 1:
        seen: dict[str, int] = {}
        for f in out:
            seen[f.id] = seen.get(f.id, 0) + 1
        for f in out:
            if seen[f.id] > 1:
                f.id = f"{f.sheet}!{f.id}"
    return out
