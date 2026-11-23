"""XE-MODERNPLAS 2607194: buyer name separated from caption and TaxID."""
import ast
from pathlib import Path
import re
import unittest

LINES=['บริษัท ซี-โมเดิร์นพลาส จำกัด','นามผู้ซื้อ',
       'บริษัท ไทย โคะอิโท จำกัด (สาขาที่ 1)',
       'เลขที่ (INV No.) 2607194','Name',
       '555 หมู่ 7 ตำบลท่าตูม อำเภอศรีมหาโพธิ','วันที่ (Date) 15/7/2026 /',
       'ที่อยู่','จังหวัดปราจีนบุรี 25140','เงื่อนไข (Terms) 30 day','Address',
       'เลขประจำตัวผู้เสียภาษี 0105529030059']


class LabelledBuyerName(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        tree=ast.parse((Path(__file__).resolve().parents[1]/'03_430000848.py').read_text())
        ns={'re':re}
        exec(compile(ast.Module(body=[n for n in tree.body if isinstance(n,ast.FunctionDef)
                                     and n.name in {'recover_labelled_buyer_name','clean_company_name',
                                                    'is_valid_customer_company_name'}],type_ignores=[]),'buyer name','exec'),ns)
        cls.recover=staticmethod(ns['recover_labelled_buyer_name'])

    def test_missing_or_fragment_name(self):
        for old in ['',None,'THAI']:
            self.assertEqual(self.recover(old,'00001',LINES,'บริษัท ซี-โมเดิร์นพลาส จำกัด'),
                             'บริษัท ไทย โคะอิโท จำกัด')

    def test_existing_complete_name_preserved(self):
        self.assertEqual(self.recover('THAI KOITO CO., LTD.','00001',LINES),'THAI KOITO CO., LTD.')

    def test_bang_phli_preserved(self):
        for old in ['',None,'THAI']:
            self.assertEqual(self.recover(old,'00000',LINES),old)

    def test_same_line_and_optional_name_caption(self):
        for lines in [['นามผู้ซื้อ: บริษัท ไทย โคะอิโท จำกัด'],
                      ['นามผู้ซื้อ','Name','','บริษัท ไทย โคะอิโท จำกัด']]:
            self.assertEqual(self.recover('','00001',lines),'บริษัท ไทย โคะอิโท จำกัด')

    def test_no_crossing_address_invoice_or_item_boundary(self):
        for boundary in ['ที่อยู่','เลขที่ (INV No.) 2607194','จำนวนเงิน','สินค้า']:
            self.assertEqual(self.recover('','00001',['นามผู้ซื้อ',boundary,'บริษัท ผู้ขาย จำกัด']),'')

    def test_supplier_and_unlabelled_names_rejected(self):
        vendor='บริษัท ผู้ขาย จำกัด'
        self.assertEqual(self.recover('','00001',['นามผู้ซื้อ',vendor],vendor),'')
        self.assertEqual(self.recover('','00001',[vendor]),'')

    def test_conflicting_buyer_labels_not_guessed(self):
        lines=['นามผู้ซื้อ','บริษัท หนึ่ง จำกัด','นามผู้ซื้อ','บริษัท สอง จำกัด']
        self.assertEqual(self.recover('','00001',lines),'')

if __name__=='__main__':unittest.main()
