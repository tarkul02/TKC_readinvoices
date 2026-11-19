"""Amount corrections isolated to verified 00001 layouts; 00000 is identical."""
import ast
from copy import deepcopy
from decimal import Decimal
from pathlib import Path
from types import SimpleNamespace as Obj
import unittest

from branch00001_amounts import reconcile_branch_amounts, written_baht_amounts
from invoice_amounts import reconcile_page_amounts, resolve_amounts, merge_amounts


SCMP_LINES = ['รวมเงิน/NET AMOUNT', '16,589.28',
              'ภาษีมูลค่าเพิ่ม/VAT 7%', '1,161.25',
              'รวมเงินทั้งสิ้น/TOTAL AMOUNT', '17,750.53']
PROSPACK_LINES = ['TOTAL AMOUNT 661.20', 'VAT 7.00% 46.28', 'NET AMOUNT -707.48',
                  'SEVEN HUNDRED SEVEN BAHT AND FORTY EIGHT STANGS.']


def record(vendor, base, tax, gross):
    return {'VendorTaxId':vendor,'TotalAmount':base,'VATAmount':tax,'AmountIncVat':gross,
            'Emessage':'Original note','Name':'Thai Koito Co.,Ltd.',
            'Address':'555 Moo 7 Thatoom 25140','TaxID':'0105529030059',
            'PurchaseOrderNo':'4300039681'}


def amounts(row):
    return tuple(row[field] for field in ('TotalAmount','VATAmount','AmountIncVat'))


class PrachinburiAmountTests(unittest.TestCase):
    def test_scmp_one_is_replaced_by_explicit_summary_base(self):
        row=record('0105528036509',1,1161.25,17750.53)
        reconcile_branch_amounts(row,'00001',SCMP_LINES)
        self.assertEqual(amounts(row),(16589.28,1161.25,17750.53))
        self.assertEqual(row['Emessage'],'Original note')
        self.assertEqual(row['_amount_pages'][0]['profile'],'00001_scmp_summary')
        # Subsequent validation/merging must retain the evidence.
        resolve_amounts(row)
        self.assertEqual(row['TotalAmount'],16589.28)

    def test_scmp_recorded_ocr_with_stray_one_before_subtotal(self):
        lines=['AMOUNT','1','T22-AT07 105102',
               'รวมเงิน/NET AMOUNT','1 16,589.28',
               '(2) ใบเสร็จรับเงินฉบับนี้ จะสมบูรณ์ต่อเมื่อบริษัทฯ เรียกเก็บเงินจากธนาคารเรียบร้อยแล้ว',
               'ภาษีมูลค่าเพิ่ม/VAT','7%','1,161.25',
               '(หนึ่งหมื่นเจ็ดพันเจ็ดร้อยห้าสิบบาทห้าสิบสามสตางค์)',
               'รวมเงินทั้งสิ้น/TOTAL AMOUNT','17,750.53']
        row=record('0105528036509',1,1161.25,17750.53)
        reconcile_branch_amounts(row,'00001',lines)
        self.assertEqual(amounts(row),(16589.28,1161.25,17750.53))
        self.assertEqual(row['Emessage'],'Original note')
        # The same exact noisy OCR still follows the old path on 00000.
        original=record('0105528036509',1,1161.25,17750.53)
        self.assertEqual(reconcile_branch_amounts(deepcopy(original),'00000',lines),
                         reconcile_page_amounts(deepcopy(original),lines))

    def test_scmp_uses_equation_not_first_or_largest_number(self):
        row=record('0105528036509',99999,1161.25,17750.53)
        lines=['NET AMOUNT 99999.00 1 16589.28','VAT 7% 1161.25','TOTAL AMOUNT 17750.53']
        reconcile_branch_amounts(row,'00001',lines)
        self.assertEqual(row['TotalAmount'],16589.28)

    def test_scmp_table_rows_and_other_amounts_not_hardcoded(self):
        for base,tax,gross in [(3029.76,212.08,3241.84),(100,0,100),(-100,-7,-107)]:
            cells=[]
            for i,(label,value) in enumerate([('NET AMOUNT',base),('VAT 7%',tax),('TOTAL AMOUNT',gross)]):
                cells.extend([Obj(row_index=i,column_index=0,content=label),
                              Obj(row_index=i,column_index=1,content=str(value))])
            row=record('0105528036509',1,tax,gross)
            reconcile_branch_amounts(row,'00001',[],[Obj(cells=cells)])
            self.assertEqual(amounts(row),(base,tax,gross))

    def test_scmp_ambiguous_or_inconsistent_summary_is_not_guessed(self):
        for lines in [SCMP_LINES+['NET AMOUNT 100.00','VAT 7% 7.00','TOTAL AMOUNT 107.00'],
                      ['NET AMOUNT 100','VAT 7% 7','TOTAL AMOUNT 120'],
                      ['NET AMOUNT 16589.28','TOTAL AMOUNT 17750.53']]:
            row=record('0105528036509',1,1161.25,17750.53)
            reconcile_branch_amounts(row,'00001',lines)
            self.assertNotIn('profile',row['_amount_pages'][0])
            self.assertEqual(row['TotalAmount'],1)
            self.assertIn('Amount review:',row['Emessage'])

    def test_prospack_slash_sign_repaired_only_with_written_total(self):
        row=record('0145534000022',661.20,46.28,-707.48)
        reconcile_branch_amounts(row,'00001',PROSPACK_LINES)
        self.assertEqual(amounts(row),(661.20,46.28,707.48))
        self.assertEqual(row['Emessage'],'Original note')
        self.assertEqual(row['_amount_pages'][0]['profile'],'00001_prospack_written_total')

    def test_prospack_no_words_or_wrong_words_does_not_flip_sign(self):
        for words in ['', 'SEVEN HUNDRED EIGHT BAHT ONLY',
                      'MINUS SEVEN HUNDRED SEVEN BAHT AND FORTY EIGHT STANGS',
                      'NEGATIVE SEVEN HUNDRED SEVEN BAHT AND FORTY EIGHT STANGS',
                      'SEVEN HUNDRED SEVEN BAHT AND FORTY EIGHT STANGS THREE HUNDRED BAHT ONLY']:
            row=record('0145534000022',661.20,46.28,-707.48)
            reconcile_branch_amounts(row,'00001',PROSPACK_LINES[:-1]+[words])
            self.assertEqual(row['AmountIncVat'],-707.48)
            self.assertIn('Amount review:',row['Emessage'])

    def test_prospack_credit_note_or_negative_base_is_not_flipped(self):
        for caption in ['CREDIT NOTE','CREDIT MEMO','ใบลดหนี้']:
            row=record('0145534000022',661.20,46.28,-707.48)
            reconcile_branch_amounts(row,'00001',PROSPACK_LINES+[caption])
            self.assertEqual(row['AmountIncVat'],-707.48)
        row=record('0145534000022',-661.20,-46.28,-707.48)
        reconcile_branch_amounts(row,'00001',PROSPACK_LINES[-1:])
        self.assertEqual(amounts(row),(-661.20,-46.28,-707.48))

    def test_branch00000_results_are_exactly_the_original_path(self):
        cases=[(record('0105528036509',1,1161.25,17750.53),SCMP_LINES),
               (record('0145534000022',661.20,46.28,-707.48),PROSPACK_LINES),
               (record('0115549009357',127692.1,8938.44,136630.4),[]),
               (record('anything',100,7,107),['VAT 7%'])]
        for row,lines in cases:
            with self.subTest(vendor=row['VendorTaxId']):
                expected=reconcile_page_amounts(deepcopy(row),lines)
                actual=reconcile_branch_amounts(deepcopy(row),'00000',lines)
                self.assertEqual(actual,expected)

    def test_unknown_vendor_on_00001_is_unchanged_from_original(self):
        row=record('9999999999999',1,1161.25,17750.53)
        self.assertEqual(reconcile_branch_amounts(deepcopy(row),'00001',SCMP_LINES),
                         reconcile_page_amounts(deepcopy(row),SCMP_LINES))

    def test_non_amount_fields_unchanged_and_no_fixed_vat_rate(self):
        row=record('0105528036509',1,10,110)
        before=deepcopy(row)
        reconcile_branch_amounts(row,'00001',['NET AMOUNT 100','VAT 10% 10','TOTAL AMOUNT 110'])
        self.assertEqual(amounts(row),(100,10,110))
        for field in ['Name','Address','TaxID','VendorTaxId','PurchaseOrderNo']:
            self.assertEqual(row[field],before[field])

    def test_written_amount_parser(self):
        for phrase,expected in [
            ('SEVEN HUNDRED SEVEN BAHT AND FORTY EIGHT STANGS.',Decimal('707.48')),
            ('THREE HUNDRED FIFTY THREE BAHT AND SEVENTY FOUR STANGS.',Decimal('353.74')),
            ('ONE THOUSAND SEVEN HUNDRED BAHT ONLY',Decimal('1700')),
            ('ONE MILLION TWO HUNDRED THOUSAND BAHT ONLY',Decimal('1200000')),
        ]:
            self.assertEqual(written_baht_amounts([phrase]),{expected})
        self.assertEqual(written_baht_amounts(['SEVEN HUNDRED SEVEN BAHT AND FORTY EIGHTY SATANG']),set())
        for invalid in ['ONE TWO BAHT ONLY', 'ONE HUNDRED AND BAHT ONLY',
                        'ONE THOUSAND FIVE THOUSAND BAHT ONLY']:
            self.assertEqual(written_baht_amounts([invalid]),set())

    def test_main_passes_cli_branch_not_vendor_branch(self):
        root=Path(__file__).resolve().parents[1]
        tree=ast.parse((root/'03_430000848.py').read_text())
        calls=[node for node in ast.walk(tree) if isinstance(node,ast.Call)
               and isinstance(node.func,ast.Name) and node.func.id=='reconcile_branch_amounts']
        self.assertEqual(len(calls),1)
        self.assertEqual(calls[0].args[1].id,'branch_email')


if __name__ == '__main__':
    unittest.main()
