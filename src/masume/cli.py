"""masume command line interface.

    masume inspect  form.xlsx [--sheet S] [--json]
    masume template form.xlsx [--sheet S]          # JSON skeleton to fill in
    masume md       form.xlsx [--sheet S] [--no-layout]
    masume fill     form.xlsx data.json -o out.xlsx [--sheet S] [--mark ■]
"""
from __future__ import annotations

import argparse
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
            out[f.id] = [{c["id"]: "" for c in f.columns}]
        elif f.kind == "choice":
            out[f.id] = f.options[0] if f.options else ""
        elif f.kind in ("date", "inline_date"):
            out[f.id] = "YYYY-MM-DD"
        else:
            out[f.id] = ""
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

    p = sub.add_parser("fill", help="JSONの値をフォームに書き込む（書式は非破壊）")
    p.add_argument("xlsx")
    p.add_argument("data", help="JSONファイル（- で標準入力）")
    p.add_argument("-o", "--output", required=True)
    p.add_argument("--sheet")
    p.add_argument("--mark", default="■", help="チェックボックスの塗り記号（既定: ■）")
    p.add_argument("--strict", action="store_true", help="未知のキーがあれば失敗")

    a = ap.parse_args(argv)
    try:
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
            if a.sheet:
                print(sheet_to_markdown(wb, a.sheet, not a.no_layout))
            else:
                print(workbook_to_markdown(wb, not a.no_layout))
        elif a.cmd == "fill":
            raw = sys.stdin.read() if a.data == "-" else open(a.data, encoding="utf-8-sig").read()
            data = json.loads(raw)
            rep = fill(a.xlsx, a.output, data, sheet=a.sheet, mark=a.mark, strict=a.strict)
            print(json.dumps(rep.to_dict(), ensure_ascii=False, indent=2))
            return 0 if not rep.errors else 1
    except (OSError, KeyError, ValueError) as e:
        print(f"error: {e}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
