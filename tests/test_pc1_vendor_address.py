"""Buyer 00001 VPI/UMC address shapes, without another OCR engine."""
import unittest

from branch00001_review import select_address
from vendor_address import select_vendor_address

VPI_TH='88/8-88/12 ม.1 ต.ขุนแก้ว อ.นครชัยศรี จ.นครปฐม 73120'
VPI_EN='88/8-88/12 Moo 1 Khunkaew, Nakornchaisri, Nakornpathom 73120'
HQ='99 Moo 5 Kanjanapisek Rd Saothonghin Bangyai Nonthaburi 11140'
VPI=['VPI INDUSTRY CO., LTD.','HEAD OFFICE : '+HQ,
     'สาขา 00002 : '+VPI_TH,'สำหรับลูกค้า','BRANCH 00002 : '+VPI_EN,
     'รหัสลูกค้า CODE','THAI KOITO','555 Moo 7 Thatoom Prachinburi 25140']
UMC='TFD Industrial Estate 1/23 Moo 5 Tambol Tha Sa-an, Amphur BangPakong, Chachoengsao 24130 Thailand.'
UMC_LINES=['UMC ELECTRONICS (THAILAND) LTD.',UMC,'Tel: 038-989-828',
           'TAXIDNo. 0105554000737 Branch : Head Office','TAX INVOICE',
           'Sold to:','THAI KOITO CO., LTD.','555 Moo 7 Thatoom Prachinburi 25140']


class AddressShapes(unittest.TestCase):
    def test_vpi_range_with_confirmed_branch(self):
        address,status=select_address('', '00002',VPI)
        self.assertEqual((address,status),(VPI_TH+' '+VPI_EN,'selected'))
        self.assertNotIn('11140',address)
        self.assertNotIn('555',address)

    def test_vpi_each_language_works_independently(self):
        for address in [VPI_TH,VPI_EN]:
            self.assertEqual(select_address('','00002',['สาขา 00002 : '+address,'รหัสลูกค้า']),
                             (address,'selected'))

    def test_vpi_buyer_00000_legacy_parser_unchanged(self):
        for address in [VPI_TH,VPI_EN]:
            self.assertEqual(select_vendor_address('','00002',['สาขา 00002 : '+address]),('','unresolved'))

    def test_umc_estate_prefix_and_all_three_invoices(self):
        for number in ['UMC-2026070178','UMC-2026070177','UMC-2026070166']:
            self.assertEqual(select_address('','00000',UMC_LINES+[number]),(UMC,'selected'))

    def test_umc_buyer_00000_legacy_parser_unchanged(self):
        self.assertEqual(select_vendor_address('','00000',UMC_LINES),('','unresolved'))

    def test_estate_prefix_in_buyer_block_is_not_supplier(self):
        self.assertEqual(select_address('','00000',['UMC','Sold to:',UMC]),('','unresolved'))

    def test_true_dates_and_missing_postcode_not_address(self):
        for text in ['15/07/2026 Moo 5 Thailand 24130','88/8-88/12 Moo 1 Khunkaew']:
            self.assertEqual(select_address('','00002',['สาขา 00002',text]),('','unresolved'))

    def test_no_assumption_of_vendor_branch(self):
        self.assertEqual(select_address('','',UMC_LINES),('','unresolved'))

if __name__=='__main__':unittest.main()
