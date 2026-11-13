"""Logo fragments must not be included in the final SupplierName."""
import unittest
from supplier_name import remove_supplier_logo_fragments

THAI = 'บริษัท พรอสเพค อุตสาหกรรม จำกัด'
ENGLISH = 'PROSPACK INDUSTRY CO., LTD.'
LINES = [THAI, 'PR', 'SPACK', ENGLISH, 'สำนักงานใหญ่ : 31/5 SALALOY AYUDHAYA 13130',
         'INDUSTRY', 'CO .,LTD', 'CUSTOMER', 'THAI KOITO COMPANY LIMITED']


class SupplierNameTests(unittest.TestCase):
    def test_four_user_results_have_one_canonical_company_name(self):
        for name in [THAI+' '+ENGLISH, THAI+' SPACK '+ENGLISH,
                     THAI+' SPACK '+ENGLISH, THAI+' '+ENGLISH]:
            self.assertEqual(remove_supplier_logo_fragments(name, LINES), THAI+' '+ENGLISH)

    def test_prefix_is_removed_using_header_not_a_hardcoded_brand(self):
        thai = 'บริษัท ตัวอย่าง จำกัด'
        english = 'EXAMPLE INDUSTRY CO., LTD.'
        self.assertEqual(remove_supplier_logo_fragments(thai+' LOGO '+english,
                         [thai, 'LOGO', english]), thai+' '+english)

    def test_english_only_logo_prefix(self):
        self.assertEqual(remove_supplier_logo_fragments('SPACK '+ENGLISH, LINES), ENGLISH)

    def test_no_complete_header_name_preserves_original(self):
        name = THAI+' SPACK '+ENGLISH
        self.assertEqual(remove_supplier_logo_fragments(name, [THAI, 'PR', 'SPACK', 'INDUSTRY CO., LTD.']), name)

    def test_legitimate_brand_word_is_not_deleted(self):
        english = 'SPACK INDUSTRY CO., LTD.'
        self.assertEqual(remove_supplier_logo_fragments(THAI+' '+english, [THAI, english]), THAI+' '+english)

    def test_buyer_names_cannot_replace_vendor_name(self):
        name = THAI+' SPACK '+ENGLISH
        self.assertEqual(remove_supplier_logo_fragments(name, [THAI, 'Buyer', ENGLISH]), name)
        self.assertEqual(remove_supplier_logo_fragments(name, [THAI, 'TAX ID 0105529030059', ENGLISH],
                         buyer_tax_ids={'0105529030059'}), name)

    def test_different_thai_company_is_not_rebuilt(self):
        name = 'บริษัท อื่น จำกัด SPACK '+ENGLISH
        self.assertEqual(remove_supplier_logo_fragments(name, LINES), name)

    def test_unrelated_complete_names_are_ambiguous(self):
        name = 'FIRST CO., LTD. SECOND CO., LTD.'
        self.assertEqual(remove_supplier_logo_fragments(name, ['FIRST CO., LTD.', 'SECOND CO., LTD.']), name)

    def test_buyer_name_boundary(self):
        name = THAI+' SPACK '+ENGLISH
        self.assertEqual(remove_supplier_logo_fragments(name, [THAI, ENGLISH], buyer_name=ENGLISH), name)


class SupplierOfficeLabelTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import ast
        import re
        from pathlib import Path

        path = Path(__file__).resolve().parents[1] / '03_430000848.py'
        tree = ast.parse(path.read_text())
        functions = [node for node in tree.body if isinstance(node, ast.FunctionDef)
                     and node.name == 'clean_supplier_legal_ending']
        cls.ns = {'re': re}
        exec(compile(ast.Module(body=functions, type_ignores=[]), str(path), 'exec'), cls.ns)

    def test_office_label_does_not_pollute_bilingual_supplier_name(self):
        thai = 'บริษัท ซิก้า (ประเทศไทย) จำกัด'
        english = 'ka (Thailand) Limited'
        for label in ('สำนักงานใหญ่', '(สำนักงานใหญ่)', '( สำนักงานใหญ่ )'):
            with self.subTest(label=label):
                self.assertEqual(self.ns['clean_supplier_legal_ending'](
                    thai + ' ' + label + ' ' + english), thai + ' ' + english)

    def test_english_only_and_label_free_names(self):
        clean = self.ns['clean_supplier_legal_ending']
        self.assertEqual(clean('EXAMPLE LTD. สำนักงานใหญ่'), 'EXAMPLE LTD.')
        name = 'บริษัท ตัวอย่าง จำกัด (มหาชน) EXAMPLE PUBLIC COMPANY LIMITED'
        self.assertEqual(clean(name), name)


if __name__ == '__main__':
    unittest.main()
