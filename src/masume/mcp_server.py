"""Zero-dependency MCP (Model Context Protocol) stdio server for masume.

Register with Claude Code:
    claude mcp add masume -- python -m masume.mcp_server
"""
from __future__ import annotations

import json
import os
import sys
import traceback

from . import __version__, ooxml
from .cli import template_for
from .detect import detect
from .fill import fill
from .render import sheet_to_markdown, workbook_to_markdown

PROTOCOL_VERSIONS = ["2025-06-18", "2025-03-26", "2024-11-05"]

TOOLS = [
    {
        "name": "list_sheets",
        "description": "List the worksheets of an .xlsx/.xlsm file.",
        "inputSchema": {
            "type": "object",
            "properties": {"path": {"type": "string", "description": "Absolute path to the .xlsx file"}},
            "required": ["path"],
        },
    },
    {
        "name": "inspect_form",
        "description": (
            "Detect the input fields of a Japanese Excel form (Excel 方眼紙 / 神エクセル): labels paired with "
            "the empty merged boxes next to or under them, checkbox choices (□), date placeholders "
            "(令和　年　月　日), underscores (＿＿) and repeating tables. Returns the fields and a JSON "
            "template you can fill and pass to fill_form."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "path": {"type": "string", "description": "Absolute path to the .xlsx file"},
                "sheet": {"type": "string", "description": "Sheet name (default: all visible sheets)"},
            },
            "required": ["path"],
        },
    },
    {
        "name": "form_to_markdown",
        "description": (
            "Render an Excel 方眼紙 form as compact Markdown (field list + logical layout rows) so that the "
            "form can be read and understood without the hundreds of empty grid columns."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "sheet": {"type": "string"},
                "layout": {"type": "boolean", "default": True},
                "fields": {
                    "type": "array",
                    "items": {"type": "object"},
                    "description": "inspect_form's 'fields' saved BEFORE filling: keeps ids stable and shows current values (use this to verify a filled form)",
                },
            },
            "required": ["path"],
        },
    },
    {
        "name": "fill_form",
        "description": (
            "Write values into the detected fields and save a NEW .xlsx. Only the target cells are changed; "
            "borders, merges, images, stamps and print settings are preserved. Keys may be field ids from "
            "inspect_form, labels, or cell references like 'C5' / 'Sheet1!C5'. Dates may be given as "
            "YYYY-MM-DD and are converted to 和暦 when the form uses it. Choice fields take an option name "
            "(or a list). Tables take a list of row objects keyed by column id."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "path": {"type": "string", "description": "Template .xlsx (never modified)"},
                "values": {"type": "object", "description": "{field id | label | cell: value}"},
                "output_path": {"type": "string", "description": "Where to save (default: <name>_filled.xlsx)"},
                "sheet": {"type": "string"},
                "mark": {"type": "string", "default": "■", "description": "Mark for checked boxes"},
                "fields": {
                    "type": "array",
                    "items": {"type": "object"},
                    "description": "Optional corrected field list (inspect_form's 'fields'); skips auto-detection",
                },
            },
            "required": ["path", "values"],
        },
    },
]


def _text(obj) -> dict:
    if not isinstance(obj, str):
        obj = json.dumps(obj, ensure_ascii=False, indent=2)
    return {"content": [{"type": "text", "text": obj}], "isError": False}


def _error(msg: str) -> dict:
    return {"content": [{"type": "text", "text": msg}], "isError": True}


def call_tool(name: str, args: dict) -> dict:
    path = args.get("path")
    if not path or not os.path.isfile(path):
        return _error(f"file not found: {path}")
    if name == "list_sheets":
        wb = ooxml.load(path)
        return _text([{"name": s.name, "hidden": s.hidden} for s in wb.sheets])
    if name == "inspect_form":
        wb = ooxml.load(path)
        fields = detect(wb, args.get("sheet"))
        return _text({"fields": [f.to_dict() for f in fields], "template": template_for(fields)})
    if name == "form_to_markdown":
        wb = ooxml.load(path)
        layout = bool(args.get("layout", True))
        fields = None
        if isinstance(args.get("fields"), list):
            from .detect import Field

            fields = [Field.from_dict(d) for d in args["fields"]]
        if args.get("sheet"):
            return _text(sheet_to_markdown(wb, args["sheet"], layout, fields))
        return _text(workbook_to_markdown(wb, layout, fields))
    if name == "fill_form":
        values = args.get("values")
        if not isinstance(values, dict):
            return _error("values must be an object")
        out = args.get("output_path")
        if not out:
            root, ext = os.path.splitext(path)
            out = f"{root}_filled{ext or '.xlsx'}"
        if os.path.abspath(out) == os.path.abspath(path):
            return _error("output_path must differ from the template path")
        fields = None
        if isinstance(args.get("fields"), list):
            from .detect import Field

            fields = [Field.from_dict(d) for d in args["fields"]]
        rep = fill(path, out, values, sheet=args.get("sheet"), mark=args.get("mark") or "■", fields=fields)
        d = rep.to_dict()
        d["output_path"] = out if not rep.errors else None
        res = _text(d)
        res["isError"] = bool(rep.errors)
        return res
    return _error(f"unknown tool: {name}")


def handle(msg: dict) -> dict | None:
    mid = msg.get("id")
    method = msg.get("method")
    if mid is None:  # notification
        return None
    try:
        if method == "initialize":
            req = (msg.get("params") or {}).get("protocolVersion")
            ver = req if req in PROTOCOL_VERSIONS else PROTOCOL_VERSIONS[0]
            result = {
                "protocolVersion": ver,
                "capabilities": {"tools": {"listChanged": False}},
                "serverInfo": {"name": "masume", "version": __version__},
                "instructions": (
                    "Use inspect_form first to learn the fields of a Japanese Excel form, then fill_form. "
                    "Never overwrite the user's template; write to a new file."
                ),
            }
        elif method == "ping":
            result = {}
        elif method == "tools/list":
            result = {"tools": TOOLS}
        elif method == "tools/call":
            params = msg.get("params") or {}
            try:
                result = call_tool(params.get("name", ""), params.get("arguments") or {})
            except Exception as e:  # tool errors are reported in-band
                result = _error(f"{type(e).__name__}: {e}")
        else:
            return {"jsonrpc": "2.0", "id": mid, "error": {"code": -32601, "message": f"method not found: {method}"}}
        return {"jsonrpc": "2.0", "id": mid, "result": result}
    except Exception as e:
        traceback.print_exc(file=sys.stderr)
        return {"jsonrpc": "2.0", "id": mid, "error": {"code": -32603, "message": str(e)}}


def main() -> None:
    stdin = sys.stdin.buffer
    stdout = sys.stdout.buffer
    for line in stdin:
        line = line.strip()
        if not line:
            continue
        try:
            msg = json.loads(line.decode("utf-8"))
        except json.JSONDecodeError:
            resp = {"jsonrpc": "2.0", "id": None, "error": {"code": -32700, "message": "parse error"}}
        else:
            if isinstance(msg, list):
                resps = [r for r in (handle(m) for m in msg) if r is not None]
                resp = resps or None
            else:
                resp = handle(msg)
        if resp is not None:
            stdout.write(json.dumps(resp, ensure_ascii=False).encode("utf-8") + b"\n")
            stdout.flush()


if __name__ == "__main__":
    main()
