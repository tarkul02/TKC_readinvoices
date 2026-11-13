"""Buyer TaxID extraction recognises the captions on KATATA invoices."""
import ast
from pathlib import Path
import re
import unittest


class CustomerTaxIDTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        path = Path(__file__).resolve().parents[1] / '03_430000848.py'
        tree = ast.parse(path.read_text())
        cls.ns = {'re': re}
        funcs = [n for n in tree.body if isinstance(n, ast.FunctionDef)
                 and n.name in {'extract_company_tax_id_from_pages', 'clean_tax_id'}]
        exec(compile(ast.Module(body=funcs, type_ignores=[]), str(path), 'exec'), cls.ns)

    def extract(self, lines):
        return self.ns['extract_company_tax_id_from_pages'](
            None, {'lines': lines}, '0105539027476')

    def test_katata_buyer_tax_id_for_both_invoices(self):
        for invoice in ['50172600537', '50172600545']:
            lines = ['KATATA ELECTRIC (THAILAND) CO., LTD.',
                     'TAX IDENTIFICATION NO. 0105539027476', invoice,
                     'MESSERS :', 'ขายให้ :', 'THAI KOITO COMPANY LIMITED (HEAD OFFICE)',
                     '370 MOO 17 T.BANGSAOTHONG A.BANGSAOTHONG', 'SAMUTPRAKARN 10570',
                     'TAX ID NO.0105529030059', 'DELIVERY TO:', 'DESCRIPTION']
            self.assertEqual(self.extract(lines), '0105529030059')

    def test_each_new_caption_independently(self):
        for caption in ['MESSERS :', 'Messrs.', 'ขายให้ :', 'นามผู้ซื้อ :']:
            self.assertEqual(self.extract([caption, 'THAI KOITO', 'TAX ID NO.0105529030059']),
                             '0105529030059')

    def test_split_label_and_spaced_number(self):
        self.assertEqual(self.extract(['ขายให้', 'THAI KOITO', 'TAX ID NO.',
                                      '010 552 903 0059']), '0105529030059')

    def test_existing_customer_caption_still_works(self):
        self.assertEqual(self.extract(['BILL TO', 'THAI KOITO', 'TAX ID 0105529030059']),
                         '0105529030059')

    def test_seller_tax_id_is_not_used(self):
        self.assertEqual(self.extract(['TAX ID 0105539027476', 'MESSERS',
                                      'THAI KOITO', 'TAX ID 0105539027476']), '')
        self.assertEqual(self.extract(['TAX ID 0105539027476', 'บริษัท ผู้ขาย จำกัด']), '')


if __name__ == '__main__':
    unittest.main()
