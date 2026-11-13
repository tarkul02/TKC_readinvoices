"""Customer Name export excludes head-office captions, including Thai abbreviations."""
import ast
from pathlib import Path
import re
import unittest


class CompanyNameTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        path = Path(__file__).resolve().parents[1] / '03_430000848.py'
        tree = ast.parse(path.read_text())
        funcs = [n for n in tree.body if isinstance(n, ast.FunctionDef)
                 and n.name in {'clean_company_name', 'build_excel_row'}]
        cls.ns = {'re': re}
        exec(compile(ast.Module(body=funcs, type_ignores=[]), str(path), 'exec'), cls.ns)

    def test_user_abbreviated_thai_name(self):
        clean = self.ns['clean_company_name']
        self.assertEqual(clean('บจก. ไทย โคะอิโท/   สำนักงานใหญ่'), 'บจก. ไทย โคะอิโท')

    def test_office_caption_spellings_and_separators(self):
        for suffix in [' / สำนักงานใหญ่', ' (สำนักงานใหญ่)', ' / สำหนักงานใหญ่',
                       ' / สำ นัก งาน ใหญ่', ' (HEAD OFFICE)', ' / HeadOffice']:
            with self.subTest(suffix=suffix):
                self.assertEqual(self.ns['clean_company_name']('บจก. ไทย โคะอิโท'+suffix),
                                 'บจก. ไทย โคะอิโท')

    def test_existing_names_are_preserved(self):
        for name in ['THAI KOITO COMPANY LIMITED', 'THAI KOITO CO., LTD.',
                     'บริษัท ไทย โคะอิโท จำกัด', 'บริษัท ตัวอย่าง จำกัด (มหาชน)', 'บจก. ไทย โคะอิโท']:
            self.assertEqual(self.ns['clean_company_name'](name), name)
        self.assertEqual(self.ns['clean_company_name']('THAI KOITO CO., LTD. (HEAD OFFICE)'),
                         'THAI KOITO CO., LTD.')

    def test_office_word_inside_company_name_is_not_deleted(self):
        name = 'บจก. สำนักงานใหญ่ เซอร์วิส'
        self.assertEqual(self.ns['clean_company_name'](name), name)

    def test_excel_export_cleans_fallback_name_without_changing_branch(self):
        record = {'CompanyName': 'บจก. ไทย โคะอิโท/ สำนักงานใหญ่',
                  'CompanyBranch': '00000', 'CompanyAddress': '370 Road 10570',
                  'SupplierName': 'Seller HEAD OFFICE'}
        row = self.ns['build_excel_row'](record)
        self.assertEqual(row['Name'], 'บจก. ไทย โคะอิโท')
        self.assertEqual(row['Branch'], '00000')
        self.assertEqual(row['Address'], '370 Road 10570')
        self.assertEqual(row['SupplierName'], 'Seller HEAD OFFICE')
        self.assertEqual(record['CompanyName'], 'บจก. ไทย โคะอิโท/ สำนักงานใหญ่')


if __name__ == '__main__':
    unittest.main()
