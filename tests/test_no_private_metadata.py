"""Every workbook in the repo must be free of author names and local paths.

Excel writes the account owner's name (docProps/core.xml) and the folder the file was
saved in (xl/workbook.xml x15ac:absPath). Files saved by real Excel for tests must be
scrubbed before they are committed.
"""
import glob
import os
import re
import unittest
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ALLOWED_AUTHORS = {"", "openpyxl", "masume"}


class NoPrivateMetadataTests(unittest.TestCase):
    def test_workbooks_have_no_author_or_path(self):
        files = glob.glob(os.path.join(ROOT, "tests", "**", "*.xlsx"), recursive=True)
        self.assertTrue(files)
        for path in files:
            with zipfile.ZipFile(path) as z:
                names = z.namelist()
                core = z.read("docProps/core.xml").decode("utf-8") if "docProps/core.xml" in names else ""
                book = z.read("xl/workbook.xml").decode("utf-8")
            for tag in ("dc:creator", "cp:lastModifiedBy"):
                for who in re.findall(rf"<{tag}>([^<]*)</{tag}>", core):
                    self.assertIn(who, ALLOWED_AUTHORS, f"{os.path.relpath(path, ROOT)}: {tag}")
            self.assertNotIn("absPath", book, os.path.relpath(path, ROOT))
