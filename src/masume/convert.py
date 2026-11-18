"""Convert legacy .xls forms to .xlsx using whatever is installed locally.

Tries, in order: Microsoft Excel (Windows, via PowerShell COM) and LibreOffice
(``soffice --headless``). Nothing is sent over the network.
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile


class ConvertError(Exception):
    pass


def _excel_com(src: str, dst: str) -> bool:
    if sys.platform != "win32" or not shutil.which("powershell"):
        return False
    ps = (
        "$ErrorActionPreference='Stop';"
        "$xl=New-Object -ComObject Excel.Application;$xl.DisplayAlerts=$false;$xl.Visible=$false;"
        "try{$wb=$xl.Workbooks.Open($env:MASUME_SRC);$wb.SaveAs($env:MASUME_DST,51);$wb.Close($false)}"
        "finally{$xl.Quit()}"
    )
    env = dict(os.environ, MASUME_SRC=os.path.abspath(src), MASUME_DST=os.path.abspath(dst))
    try:
        p = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
                           env=env, capture_output=True, timeout=180)
    except (OSError, subprocess.TimeoutExpired):
        return False
    return p.returncode == 0 and os.path.isfile(dst)


def _soffice_candidates() -> list[str]:
    c = [shutil.which("soffice"), shutil.which("libreoffice")]
    for base in (os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)")):
        if base:
            c.append(os.path.join(base, "LibreOffice", "program", "soffice.exe"))
    c.append("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    return [x for x in c if x and os.path.isfile(x)]


def _libreoffice(src: str, dst: str) -> bool:
    for exe in _soffice_candidates():
        with tempfile.TemporaryDirectory() as td:
            try:
                subprocess.run([exe, "--headless", "--convert-to", "xlsx", "--outdir", td, src],
                               capture_output=True, timeout=180)
            except (OSError, subprocess.TimeoutExpired):
                continue
            out = os.path.join(td, os.path.splitext(os.path.basename(src))[0] + ".xlsx")
            if os.path.isfile(out):
                shutil.move(out, dst)
                return True
    return False


def xls_to_xlsx(src: str, dst: str | None = None) -> str:
    if not os.path.isfile(src):
        raise ConvertError(f"file not found: {src}")
    dst = dst or os.path.splitext(src)[0] + ".xlsx"
    if os.path.abspath(dst) == os.path.abspath(src):
        raise ConvertError("output must differ from input")
    if _excel_com(src, dst) or _libreoffice(src, dst):
        return dst
    raise ConvertError(
        "変換できませんでした。Excel か LibreOffice が必要です。"
        "Excel で開いて「名前を付けて保存 → Excel ブック (*.xlsx)」でも変換できます。"
    )
