"""Regression coverage for PC5 withholding remarks split by OCR."""
import ast
import json
from pathlib import Path
import re
import unittest


class TaxRemarkTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        path = Path(__file__).resolve().parents[1] / "03_430000848.py"
        function = next(
            node for node in ast.parse(path.read_text()).body
            if isinstance(node, ast.FunctionDef) and node.name == "extract_tax_remark"
        )
        namespace = {"re": re}
        exec(compile(ast.Module(body=[function], type_ignores=[]), str(path), "exec"), namespace)
        cls.extract = staticmethod(namespace["extract_tax_remark"])

    def test_recorded_pc5_headers(self):
        path = Path(__file__).parent / "fixtures" / "pc5_address_headers.json"
        headers = json.loads(path.read_text())
        for page in ("15", "16"):
            with self.subTest(page=page):
                text = "\n".join(headers[page]["lines"])
                self.assertEqual(self.extract(None, {"full_text": text}),
                                 "PLEASE DO NOT DEDUCT WITHHOLDING TAX")

    def test_line_end_noise_requires_intact_remark(self):
        for prefix, expected in (
            ("PLEASE DO NOT DEDUCT J", "PLEASE DO NOT DEDUCT WITHHOLDING TAX"),
            ("PLEASE DO NOT DEDUCT /", "PLEASE DO NOT DEDUCT WITHHOLDING TAX"),
            ("PLEASE DO NOT DEDUCT OTHER", ""),
            ("PLEASE DO DEDUCT J", ""),
        ):
            with self.subTest(prefix=prefix):
                text = prefix + "\nWITHHOLDING TAX"
                self.assertEqual(self.extract(None, {"full_text": text}), expected)
