"""SANAC 26/07-02-014: seller ID must not populate buyer TaxID on 00001."""
import ast
from pathlib import Path
import re
import unittest

SELLER='0105554071146'
BUYER='0105529030059'
LINES=['SANAC TRADING (THAILAND) CO.,LTD.',
       'เลขประจำตัวผู้เสียภาษีอากร',SELLER,'ลูกค้า :','THAI KOITO CO., LTD.',
       'เลขที่ :','26/07-02-014','Customer','BRANCH NO.1 : 555 MOO 7 T.THATOOM',
       'PRACHINBURI 25140','สถานที่ส่งสินค้า :','TAX ID : '+BUYER]


class BuyerTaxCollision(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        source=Path(__file__).resolve().parents[1]/'03_430000848.py'
        tree=ast.parse(source.read_text());ns={'re':re}
        exec(compile(ast.Module(body=[n for n in tree.body if isinstance(n,ast.FunctionDef)
                                     and n.name in {'resolve_customer_tax_id','clean_tax_id',
                                                    'extract_company_tax_id_from_pages'}],type_ignores=[]),'customer tax','exec'),ns)
        cls.resolve=staticmethod(ns['resolve_customer_tax_id'])
        cls.extract=staticmethod(ns['extract_company_tax_id_from_pages'])

    def test_sanac_collision_recovers_buyer_id(self):
        self.assertEqual(self.resolve(SELLER,SELLER,None,{'lines':LINES},'00001'),BUYER)

    def test_valid_distinct_azure_value_kept(self):
        self.assertEqual(self.resolve(BUYER,SELLER,None,{'lines':[]},'00001'),BUYER)

    def test_bang_phli_valid_id_behavior_unchanged(self):
        self.assertEqual(self.resolve(SELLER,SELLER,None,{'lines':LINES},'00000'),SELLER)

    def test_collision_with_no_buyer_evidence_leaves_empty(self):
        for lines in [['Customer','TAX ID '+SELLER],['TAX ID '+SELLER],[]]:
            self.assertEqual(self.resolve(SELLER,SELLER,None,{'lines':lines},'00001'),'')

    def test_invalid_and_empty_values_still_follow_old_fallback(self):
        for branch in ['00000','00001']:
            for value in [None,'','123']:
                self.assertEqual(self.resolve(value,SELLER,None,{'lines':LINES},branch),
                                 self.extract(None,{'lines':LINES},SELLER))

    def test_not_hardcoded_to_thai_koito(self):
        another='0105539027476'
        lines=['Customer Name: Another Company','TAX ID '+another]
        self.assertEqual(self.resolve(SELLER,SELLER,None,{'lines':lines},'00001'),another)

if __name__=='__main__':unittest.main()
