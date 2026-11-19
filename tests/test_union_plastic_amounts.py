"""UNION PLASTIC BK260715048: recover the missing subtotal from OCR evidence."""
from copy import deepcopy
from types import SimpleNamespace as Obj
import unittest

from branch00001_amounts import reconcile_branch_amounts
from invoice_amounts import reconcile_page_amounts

LINES=['รวมราคาทั้งสิ้น','15,879.06',
       '(หนึ่งหมื่นหกพันเก้าร้อยเก้าสิบบาท ห้าสิบเก้าสตางค์)',
       'จำนวนภาษีมูลค่าเพิ่ม 7%','1,111.53',
       'บิลค่าจ้างทำของใบนี้จะต้องถูกหักภาษี ณ ที่จ่ายร้อยละ 3',
       'จำนวนเงินรวมทั้งสิ้น','16,990.59']


def row():
    return dict(VendorTaxId='0107537000998', TotalAmount=None,
                VATAmount=1111.53, AmountIncVat=16990.59, Name='THAI KOITO',
                VendorBranch='00000', PurchaseOrderNo='4300039551')


class UnionPlasticAmounts(unittest.TestCase):
    def test_missing_subtotal_recovered(self):
        r=row();reconcile_branch_amounts(r,'00001',LINES)
        self.assertEqual([r[f] for f in ('TotalAmount','VATAmount','AmountIncVat')],
                         [15879.06,1111.53,16990.59])
        self.assertEqual(r['_amount_pages'][0]['profile'],'00001_union_plastic_summary')
        self.assertEqual(r['Name'],'THAI KOITO')
        self.assertEqual(r['VendorBranch'],'00000')

    def test_buyer_00000_exact_legacy_result(self):
        self.assertEqual(reconcile_branch_amounts(row(),'00000',LINES),
                         reconcile_page_amounts(row(),LINES))

    def test_other_vendor_exact_legacy_result(self):
        r=row();r['VendorTaxId']='9999999999999'
        self.assertEqual(reconcile_branch_amounts(deepcopy(r),'00001',LINES),
                         reconcile_page_amounts(deepcopy(r),LINES))

    def test_table_summary(self):
        table=Obj(cells=[Obj(row_index=i,column_index=j,content=text)
                        for i, pair in enumerate([('รวมราคาทั้งสิ้น','15,879.06'),
                                                  ('จำนวนภาษีมูลค่าเพิ่ม 7%','1,111.53'),
                                                  ('จำนวนเงินรวมทั้งสิ้น','16,990.59')])
                        for j,text in enumerate(pair)])
        r=row();reconcile_branch_amounts(r,'00001',[],[table])
        self.assertEqual(r['TotalAmount'],15879.06)

    def test_no_repair_without_complete_unambiguous_consistent_evidence(self):
        for lines in [LINES[:-2],
                      [x.replace('16,990.59','16,999.59') for x in LINES],
                      LINES+['รวมราคาทั้งสิ้น','15,000.00'],
                      ['CREDIT NOTE']+LINES,
                      ['รวมราคาทั้งสิ้น','ITEM X','15,879.06']+LINES[2:]]:
            r=row();reconcile_branch_amounts(r,'00001',lines)
            self.assertNotEqual(r.get('_amount_pages',[{}])[0].get('profile'),
                                '00001_union_plastic_summary')

    def test_zero_vat_without_fixed_rate(self):
        lines=['รวมราคาทั้งสิ้น 100.00','จำนวนภาษีมูลค่าเพิ่ม 0% 0.00',
               'จำนวนเงินรวมทั้งสิ้น 100.00']
        r=row();reconcile_branch_amounts(r,'00001',lines)
        self.assertEqual((r['TotalAmount'],r['VATAmount'],r['AmountIncVat']),(100,0,100))

if __name__=='__main__':unittest.main()
