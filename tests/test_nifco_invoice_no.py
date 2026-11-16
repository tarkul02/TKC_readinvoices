"""NIFCO invoice prefixes depend on evidence in the full OCR text."""
import ast
from pathlib import Path
import re
import unittest


class NifcoInvoiceTests(unittest.TestCase):
    @staticmethod
    def formatter():
        path = Path(__file__).resolve().parents[1] / '03_430000848.py'
        tree = ast.parse(path.read_text())
        nodes = [n for n in tree.body if isinstance(n, ast.FunctionDef)
                 and n.name == 'format_nifco_invoice_no']
        ns = {'re': re}
        exec(compile(ast.Module(body=nodes, type_ignores=[]), str(path), 'exec'), ns)
        return ns['format_nifco_invoice_no']

    def test_oth_token_or_invoice_prefix_adds_prefix_once(self):
        format_number = self.formatter()
        for text in ('NIFCO OTH', 'NIFCO OTH/', 'NIFCO OTH2606200023',
                     'NIFCO OTH/2606200023', 'นิฟโก้ oth : 2606200023',
                     'NIFCO\nInvoice No.\nOTH\n2606200023'):
            for value in ('2606200023', 'OTH2606200023', 'OTH/2606200023', 'OTHOTH2606200023'):
                with self.subTest(text=text, value=value):
                    result = format_number(value, text)
                    self.assertEqual(result, 'OTH/2606200023')
                    self.assertEqual(format_number(result, text), result)

    def test_no_oth_removes_old_prefix_and_ignores_other_words(self):
        format_number = self.formatter()
        for text in ('NIFCO (THAILAND) CO., LTD.', 'NIFCO OTHER',
                     'NIFCO BOTH', 'NIFCO CLOTH', 'NIFCO OTHERS', '', None):
            for value in ('2606200023', 'OTH2606200023', 'OTH/2606200023'):
                with self.subTest(text=text, value=value):
                    self.assertEqual(format_number(value, text), '2606200023')

    def test_missing_invoice_stays_empty(self):
        format_number = self.formatter()
        for text in ('NIFCO OTH/', 'NIFCO (THAILAND)'):
            for value in (None, '', ' ', 'OTH', 'OTH/'):
                self.assertEqual(format_number(value, text), '')
