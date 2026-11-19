"""SANKO SET2607-192 office suffix must not swallow the city/postcode."""
import unittest
from branch00001_review import select_address
from vendor_address import select_vendor_address

STREET='461 MOO 7 TAMBOL THATOOM AMPHUR SRIMAHAPHOTE'
CITY='PRACHINBURI 25140 THAILAND'
LINES=['3+','SANKO ELECTRONICS (THAILAND) CO., LTD.','ต้นฉบับ',STREET,
       CITY+' (HEAD OFFICE)','ORIGINAL','SANKO',
       'TEL: 66(0) 3748 1095-6 FAX : 66(0) 3748 1097',
       'เลขประจำตัวผู้เสียภาษี / Tax ID 0255552000135','CLIENT / ADDRESS',
       'T014 THAI KOITO CO., LTD (BRANCH NO.00001)',
       '555 Moo 7 T.Thatoom A.Srimahaphote','Prachinburi 25140',
       'Tax ID: 0105529030059']


class SankoAddress(unittest.TestCase):
    def test_recorded_ocr_retains_city_and_removes_office_caption(self):
        address,status=select_address('','00000',LINES,buyer_tax_ids=['0105529030059'],
                                      buyer_name='T014 THAI KOITO CO., LTD')
        self.assertEqual((address,status),(STREET+' '+CITY,'selected'))
        self.assertNotIn('HEAD OFFICE',address)
        self.assertNotIn('555',address)

    def test_whole_address_before_suffix(self):
        address=STREET+' '+CITY
        self.assertEqual(select_address('','00000',[address+' (HEAD OFFICE)']),
                         (address,'selected'))

    def test_legacy_buyer_00000_is_unchanged(self):
        address,status=select_vendor_address('','00000',LINES)
        self.assertEqual((address,status),(STREET,'selected'))

    def test_suffix_does_not_assign_other_branch_street(self):
        lines=['สาขา 00002',STREET,CITY+' (HEAD OFFICE)']
        self.assertEqual(select_address('','00000',lines),('','unresolved'))

    def test_selected_vendor_branch_still_required(self):
        for branch in ['','00001']:
            self.assertEqual(select_address('',branch,LINES),('','unresolved'))

    def test_buyer_boundary_still_blocks_suffix_address(self):
        self.assertEqual(select_address('','00000',['SANKO','Sold to:',STREET,CITY+' (HEAD OFFICE)']),
                         ('','unresolved'))

if __name__=='__main__':unittest.main()
