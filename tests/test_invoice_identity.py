"""Missing invoice numbers must never cause independent documents to disappear."""
import ast
import json
from pathlib import Path
import re
import unittest
from po_patterns import load_po_patterns, is_po_value

ROOT = Path(__file__).resolve().parents[1]


class InvoiceIdentityTests(unittest.TestCase):
    def setUp(self):
        tree = ast.parse((ROOT / '03_430000848.py').read_text())
        names = {'extract_tax_invoice_no_from_pages','find_first_by_patterns',
                 'format_nifco_invoice_no','add_or_merge_row','append_msg','clean_tax_id'}
        self.merges = []
        self.ns = {'re':re, 'is_po_value':is_po_value,
                   'PO_PATTERNS':[p for ps in load_po_patterns().values() for p in ps],
                   'INVOICE_PATTERNS':[p['regex'] for p in json.loads((ROOT/'patternsInvoice.json').read_text())['invoice_patterns']],
                   'merge_invoice_row':lambda a,b:self.merges.append((a,b))}
        exec(compile(ast.Module(body=[n for n in tree.body if isinstance(n,ast.FunctionDef)
                                     and n.name in names],type_ignores=[]),'invoice identity','exec'),self.ns)

    def extract(self,text):
        return self.ns['extract_tax_invoice_no_from_pages'](None,{'full_text':text})

    def test_siam_split_label_and_number(self):
        self.assertEqual(self.extract('วันที่ 17/06/2569\nเลขที่ใบกำกับ\n69-06-1698\nPO4100071153'), '69-06-1698')
        self.assertEqual(self.extract('เลขที่ใบกำกับภาษี : 69-06-1698'), '69-06-1698')

    def test_date_alone_is_not_the_new_invoice_pattern(self):
        self.assertEqual(self.extract('วันที่ 17-06-2569'), '')

    def test_label_fallback_works_without_json_patterns(self):
        self.ns['INVOICE_PATTERNS']=[]
        self.assertEqual(self.extract('เลขที่ใบกำกับ\n69-06-1698'), '69-06-1698')

    def test_missing_invoice_kept_separate_even_for_same_vendor(self):
        for vendor in ['SIAM','SHIMODA']:
            rows=[{'TaxInvoiceNo':'D2606-0129','Assignment':'no.2_.pdf','SupplierName':'SHIMODA'}]
            new={'TaxInvoiceNo':'','Assignment':'no.2_.pdf','SupplierName':vendor}
            self.ns['add_or_merge_row'](rows,new)
            self.assertEqual(len(rows),2)
            self.assertIn('kept as separate record',new['Emessage'])
        self.assertFalse(self.merges)

    def test_same_invoice_same_vendor_still_merges(self):
        old={'TaxInvoiceNo':'69-06-1698','VendorTaxId':'0115545004318'}
        new=dict(old)
        rows=[old];self.ns['add_or_merge_row'](rows,new)
        self.assertEqual(len(rows),1)
        self.assertEqual(len(self.merges),1)

    def test_same_invoice_different_vendor_does_not_merge(self):
        old={'TaxInvoiceNo':'12345','VendorTaxId':'0115545004318'}
        new={'TaxInvoiceNo':'12345','VendorTaxId':'0105539027476'}
        rows=[old];self.ns['add_or_merge_row'](rows,new)
        self.assertEqual(len(rows),2)
        self.assertFalse(self.merges)

    def test_same_invoice_without_vendor_evidence_does_not_merge(self):
        rows=[{'TaxInvoiceNo':'12345'}]
        self.ns['add_or_merge_row'](rows,{'TaxInvoiceNo':'12345'})
        self.assertEqual(len(rows),2)


if __name__ == '__main__':
    unittest.main()
