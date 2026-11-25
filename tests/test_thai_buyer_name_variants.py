"""Accept only Thai buyer names explicitly present in Master."""
import unittest
from Matching_TextFile import company_name_matches,supplier_names_match

FULL='บริษัท ไทย โคะอิโท จำกัด'
SHORT='บจก. ไทย โคะอิโท'


class ThaiBuyerAliases(unittest.TestCase):
    def test_each_name_requires_corresponding_master_value(self):
        for name in [FULL,SHORT]:
            for master in [FULL,SHORT]:
                self.assertEqual(company_name_matches(name,'',master),name == master)

    def test_punctuation_and_extra_spaces_supported(self):
        self.assertTrue(company_name_matches('บจก ไทย  โคะอิโท','',SHORT))

    def test_other_spelling_or_missing_legal_form_not_added(self):
        for name in ['ไทย โคะอิโท','บริษัท ไทย โกะอิโท จำกัด','บจก. ไทย โตะอิโท',
                     'บริษัท ไทย โคะอิโท จำกัด สำนักงานใหญ่']:
            self.assertFalse(company_name_matches(name,'',FULL))
        self.assertFalse(company_name_matches(SHORT,'','บริษัท อื่น จำกัด'))

    def test_supplier_comparison_not_relaxed(self):
        self.assertFalse(supplier_names_match(FULL,SHORT))

if __name__=='__main__':unittest.main()
