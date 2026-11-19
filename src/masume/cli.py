"""masume command line interface.

    masume inspect  form.xlsx [--sheet S] [--json]
    masume template form.xlsx [--sheet S]          # JSON skeleton to fill in
    masume md       form.xlsx [--sheet S] [--no-layout]
    masume fill     form.xlsx data.json -o out.xlsx [--sheet S] [--mark ■] [--fields f.json]
    masume convert  old.xls [-o new.xlsx]
"""
from __future__ import annotations

import argparse
import zipfile
import json
import sys

from . import __version__, ooxml
from .detect import detect
from .fill import fill
from .render import KIND_JA, sheet_to_markdown, workbook_to_markdown


def _stdout_utf8() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    except Exception:
        pass


def template_for(fields) -> dict:
    out: dict = {}
    for f in fields:
        if f.kind == "table":
            out[f.id] = []
        else:
            out[f.id] = None  # null = leave untouched
    return out


def _load_fields(path: str):
    from .detect import Field

    with open(path, encoding="utf-8-sig") as fh:
        raw = json.load(fh)
    if isinstance(raw, dict):
        raw = raw.get("fields", [])
    return [Field.from_dict(d) for d in raw]


def _brief(report: dict) -> dict:
    """One entry per field instead of one per cell (法人番号 is 13 cells)."""
    key = "written" if "written" in report else "planned"
    per: dict[str, list] = {}
    for w in report.get(key, []):
        per.setdefault(w.get("field") or w["key"], []).append(w["value"])
    def join(v):
        if len(v) == 1:
            return v[0]
        t = ["" if x is None else str(x) for x in v]
        return "".join(t) if all(len(x) <= 1 for x in t) else " / ".join(t)  # 1文字1マス vs 分割欄
    summary = {k: join(v) for k, v in per.items()}
    out = dict(report)
    out[key] = summary
    return out


def main(argv: list[str] | None = None) -> int:
    _stdout_utf8()
    ap = argparse.ArgumentParser(prog="masume", description="Excel方眼紙フォームをAIが読み書きするためのツール")
    ap.add_argument("--version", action="version", version=f"masume {__version__}")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("inspect", help="入力欄を検出して一覧表示")
    p.add_argument("xlsx")
    p.add_argument("--sheet")
    p.add_argument("--json", action="store_true")

    p = sub.add_parser("template", help="記入用JSONの雛形を出力")
    p.add_argument("xlsx")
    p.add_argument("--sheet")

    p = sub.add_parser("md", help="LLM向けMarkdownに変換")
    p.add_argument("xlsx")
    p.add_argument("--sheet")
    p.add_argument("--no-layout", action="store_true")
    p.add_argument("--fields", help="記入前に保存した inspect --json を使う（記入済みファイルの確認用）")

    p = sub.add_parser("fill", help="JSONの値をフォームに書き込む（書式は非破壊）")
    p.add_argument("xlsx")
    p.add_argument("data", help="JSONファイル（- で標準入力）")
    p.add_argument("-o", "--output", required=True)
    p.add_argument("--sheet")
    p.add_argument("--mark", default="■", help="チェックボックスの塗り記号（既定: ■）")
    p.add_argument("--strict", action="store_true", help="未知のキーがあれば失敗")
    p.add_argument("--fields", help="inspect --json の出力（手で修正したもの）を検出結果の代わりに使う")
    p.add_argument("--brief", action="store_true", help="書いた内容を欄ごとに1行にまとめて出力する")

    p = sub.add_parser("convert", help="旧形式 .xls を .xlsx に変換（Excel または LibreOffice を使用）")
    p.add_argument("xls")
    p.add_argument("-o", "--output")

    a = ap.parse_args(argv)
    try:
        if a.cmd == "convert":
            from .convert import ConvertError, xls_to_xlsx

            try:
                print(xls_to_xlsx(a.xls, a.output))
            except ConvertError as e:
                print(f"error: {e}", file=sys.stderr)
                return 2
            return 0
        if a.cmd in ("inspect", "template"):
            wb = ooxml.load(a.xlsx)
            fields = detect(wb, a.sheet)
            if a.cmd == "template":
                print(json.dumps(template_for(fields), ensure_ascii=False, indent=2))
            elif a.json:
                print(json.dumps([f.to_dict() for f in fields], ensure_ascii=False, indent=2))
            else:
                for f in fields:
                    extra = ""
                    if f.options:
                        extra = "  選択肢=" + "/".join(f.options)
                    elif f.columns:
                        extra = "  列=" + "/".join(c["id"] for c in f.columns) + f" 最大{len(f.rows)}行"
                    sec = f"[{f.section}] " if f.section else ""
                    print(f"{f.sheet}!{f.range:<12} {KIND_JA.get(f.kind, f.kind):<10} {sec}{f.id}{extra}")
                print(f"\n{len(fields)} fields", file=sys.stderr)
        elif a.cmd == "md":
            wb = ooxml.load(a.xlsx)
            fields = _load_fields(a.fields) if a.fields else None
            if a.sheet:
                print(sheet_to_markdown(wb, a.sheet, not a.no_layout, fields))
            else:
                print(workbook_to_markdown(wb, not a.no_layout, fields))
        elif a.cmd == "fill":
            raw = sys.stdin.read() if a.data == "-" else open(a.data, encoding="utf-8-sig").read()
            data = json.loads(raw)
            fields = _load_fields(a.fields) if a.fields else None
            rep = fill(a.xlsx, a.output, data, sheet=a.sheet, mark=a.mark, strict=a.strict, fields=fields)
            out = rep.to_dict()
            if a.brief:
                out = _brief(out)
            print(json.dumps(out, ensure_ascii=False, indent=2))
            return 0 if not rep.errors else 1
    except (OSError, KeyError, ValueError, zipfile.BadZipFile, ooxml.WriteError) as e:
        print(f"error: {e}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
