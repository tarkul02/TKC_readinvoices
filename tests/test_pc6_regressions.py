"""PC6 summary/footer/address regressions; buyer 00000 retains legacy behavior."""
from copy import deepcopy
import json
from pathlib import Path
import unittest
from unittest.mock import patch

from branch00001_amounts import reconcile_branch_amounts
from branch00001_review import select_address, reread_buyer_tax_id
from invoice_amounts import reconcile_page_amounts
from vendor_address import select_vendor_address

OCR=json.loads((Path(__file__).parent/'fixtures/pc6_ocr_excerpts.json').read_text())


def row(base,tax,gross):
    return dict(VendorTaxId='0105522000600',TotalAmount=base,VATAmount=tax,AmountIncVat=gross)


class FocusTests(unittest.TestCase):
    def test_extra_gross_digit(self):
        r=row(10968.52,767.8,411736.32)
        reconcile_branch_amounts(r,'00001',OCR['10'])
        self.assertEqual(r['AmountIncVat'],11736.32)

    def test_swapped_base_and_gross_and_separator(self):
        r=row(8590.82,562.02,8028.8)
        reconcile_branch_amounts(r,'00001',OCR['11'])
        self.assertEqual([r[f] for f in ('TotalAmount','VATAmount','AmountIncVat')],[8028.8,562.02,8590.82])

    def test_00000_and_unknown_vendor_exactly_legacy(self):
        for branch,vendor in [('00000','0105522000600'),('00001','9999999999999')]:
            r=row(8590.82,562.02,8028.8);r['VendorTaxId']=vendor
            self.assertEqual(reconcile_branch_amounts(deepcopy(r),branch,OCR['11']),
                             reconcile_page_amounts(deepcopy(r),OCR['11']))

    def test_missing_caption_or_credit_no_repair(self):
        for lines in [[x for x in OCR['10'] if 'รวมราคาทั้งสิ้น' not in x],['CREDIT NOTE']+OCR['10']]:
            r=row(10968.52,767.8,411736.32)
            reconcile_branch_amounts(r,'00001',lines)
            self.assertNotEqual(r.get('_amount_pages',[{}])[0].get('profile'),'00001_focus_summary')


class AddressTests(unittest.TestCase):
    def test_conflicting_houses_need_independent_read(self):
        for p in ['19','23']:
            with patch('branch00001_review._local_header',return_value='Head Office: 700/869 Moo 5 Chonburi 20160'):
                address,status=select_address('','00000',OCR[p],pdf_path='sample.pdf')
            self.assertEqual(status,'selected')
            self.assertIn('700/869',address)
            self.assertNotIn('700/369',address)

    def test_unconfirmed_or_ambiguous_house_is_not_guessed(self):
        for independent in ['', '700/369 700/869']:
            with patch('branch00001_review._local_header',return_value=independent):
                self.assertEqual(select_address('','00000',OCR['19']),('','unresolved'))

    def test_footer_address_and_no_buyer_leak(self):
        name='บริษัท บางกอกพัฒนามอเตอร์ จำกัด Bangkok Metropolis Motor Co., Ltd'
        for p in ['26','28']:
            address,status=select_address('555 Moo 7 Prachinburi 25140','00000',OCR[p],supplier_name=name,buyer_tax_ids=['0105529030059'])
            self.assertEqual(status,'selected')
            self.assertIn('74110',address)
            self.assertNotIn('555',address)
            self.assertNotIn('25140',address)

    def test_raw_buyer_address_is_rejected_without_footer(self):
        self.assertEqual(select_address('555 Moo 7 Prachinburi 25140','00000',OCR['26'][:20]),('','unresolved'))

    def test_wrong_company_or_branch_footer_rejected(self):
        for name,branch in [('บริษัท คนละบริษัท จำกัด','00000'),('บริษัท บางกอกพัฒนามอเตอร์ จำกัด','00001')]:
            self.assertEqual(select_address('',branch,OCR['26'],supplier_name=name),('','unresolved'))

    def test_standalone_postcode_joins_english_address(self):
        address,status=select_address('','00000',OCR['29'])
        self.assertEqual(status,'selected')
        self.assertIn('38/20 Moo 1',address)
        self.assertIn('24140 Thailand.',address)
        self.assertNotIn('555',address)
        self.assertEqual(select_vendor_address('','00000',OCR['29']),('','unresolved'))

    def test_tax_id_not_corrected_without_independent_evidence(self):
        with patch('branch00001_review.shutil.which',return_value=None):
            self.assertEqual(reread_buyer_tax_id('missing.pdf',None,'0105529030055',['0105529030059']),'0105529030055')

if __name__=='__main__':unittest.main()
