"""NIFCO invoice prefixes must agree with the Text Report."""
import ast
from pathlib import Path
import re
import unittest


class NifcoInvoiceTests(unittest.TestCase):
    def test_prefix_is_added_once_and_empty_invoice_stays_empty(self):
        path = Path(__file__).resolve().parents[1] / '03_430000848.py'
        tree = ast.parse(path.read_text())
        nodes = [n for n in tree.body if isinstance(n, ast.FunctionDef)
                 and n.name == 'format_nifco_invoice_no']
        ns = {'re': re}
        exec(compile(ast.Module(body=nodes, type_ignores=[]), str(path), 'exec'), ns)
        format_number = ns['format_nifco_invoice_no']
        for number in ('2606200023', '2606200022', '2606200019', '2606190094'):
            for value in (number, 'OTH'+number, 'OTH/'+number, 'OTHOTH'+number):
                with self.subTest(value=value):
                    result = format_number(value)
                    self.assertEqual(result, 'OTH/' + number)
                    self.assertEqual(format_number(result), result)
        for value in (None, '', ' ', 'OTH', 'OTH/'):
            self.assertEqual(format_number(value), '')
