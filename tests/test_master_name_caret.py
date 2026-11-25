"""Master bilingual name alternatives use ^, never commas in legal names."""
import tempfile
from pathlib import Path
import unittest
from openpyxl import Workbook
from Matching_TextFile import company_name_matches,split_master_name_variants,load_master_supplier

EN='THAI KOITO COMPANY LIMITED^THAI KOITO CO., LTD.'
TH='บริษัท ไทย โคะอิโท จำกัด^บจก. ไทย โคะอิโท'


class MasterNameCaret(unittest.TestCase):
    def test_each_master_alternative(self):
        for name in [*EN.split('^'),*TH.split('^')]:
            self.assertTrue(company_name_matches(name,EN,TH),name)

    def test_split_keeps_commas_and_ignores_empty_duplicate_names(self):
        self.assertEqual(split_master_name_variants(' ^ THAI KOITO CO., LTD. ^^ THAI KOITO CO., LTD. ^'),
                         ('THAI KOITO CO., LTD.',))
        self.assertEqual(split_master_name_variants(None),())

    def test_does_not_join_alternatives_or_accept_unlisted_spellings(self):
        for name in ['บริษัท ไทย โกะอิโท จำกัด','THAI KOLTO CO., LTD.','ไทย โคะอิโท',
                     'THAI KOITO','THAI KOITO TRADING LIMITED','THAI KOITO COMPANY LIMITED THAI KOITO CO., LTD.']:
            self.assertFalse(company_name_matches(name,EN,TH),name)

    def test_existing_normalization_and_either_language(self):
        for name in ['thai koito co ltd','THẠI KOITO CO., LTD.',
                     'บริษัท  ไทย โคะอิโท จำกัด','บริษัท ไทย โกะอิโท จำกัด THAI KOITO CO., LTD.']:
            self.assertTrue(company_name_matches(name,EN,TH),name)

    def test_explicit_list_uses_master_not_hardcoded_thai_alias(self):
        self.assertFalse(company_name_matches('บจก. ไทย โคะอิโท','OTHER^ANOTHER', 'บริษัท ไทย โคะอิโท จำกัด'))

    def test_single_value_master_keeps_previous_behavior(self):
        self.assertTrue(company_name_matches('THAI KOITO CO., LTD.','THAI KOITO COMPANY LIMITED',''))
        self.assertTrue(company_name_matches('บจก. ไทย โคะอิโท','','บริษัท ไทย โคะอิโท จำกัด'))

    def test_master_file_retains_branch_address_and_matches_names(self):
        wb=Workbook();ws=wb.active;ws.append(['No','Vendor account','Prones','Tax ID','Company Name (Eng)','Company Name (TH)','Branch','Street','Postcode'])
        ws.append([1,None,None,'0105529030059',EN,TH,'00001','555 Moo 7','25140'])
        with tempfile.TemporaryDirectory() as folder:
            path=Path(folder)/'master.xlsx';wb.save(path)
            records=load_master_supplier(path)['0105529030059']
        wb.close()
        self.assertEqual(len(records),1)
        row=records[0];self.assertEqual(row['BranchCode'],'00001')
        self.assertEqual(row['StreetNo'],'555 Moo 7')
        for name in ['THAI KOITO CO., LTD.','บจก. ไทย โคะอิโท']:
            self.assertTrue(company_name_matches(name,row['SupplierName_Master_EN'],row['SupplierName_Master_TH']))

if __name__=='__main__':unittest.main()
