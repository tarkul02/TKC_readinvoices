"""Branch-scoped seller address selection; synthetic text follows the supplied PDF."""
import unittest
from vendor_address import select_vendor_address


HEAD = '358-358/1 หมู่ 17 นิคมอุตสาหกรรมบางพลี จังหวัดสมุทรปราการ 10570'
HEAD_EN = '358-358/1 MOO 17 BANGPLEE INDUSTRIAL ESTATE SAMUTPRAKARN THAILAND 10570'
BRANCH = '7/288 หมู่ 6 นิคมอุตสาหกรรมอมตะซิตี้ ระยอง 21140'
BRANCH_EN = '7/288 MOO 6 AMATA CITY RAYONG THAILAND 21140'
RAW = ' '.join([HEAD, HEAD_EN, BRANCH, BRANCH_EN])
LINES = ['AAPICO Plastics Public Company Limited',
         '☑ สำนักงานใหญ่ ' + HEAD, HEAD_EN, 'TEL: 0-2315-3450 FAX: 0-2315-3334',
         '□ สาขาที่ 2 ' + BRANCH, BRANCH_EN, 'TEL: 0-3865-0888',
         'เลขประจำตัวผู้เสียภาษีอากร 0107537000131',
         'ชื่อผู้ซื้อ THAI KOITO', '370 MOO 17 SAMUTPRAKARN 10570',
         'เลขประจำตัวผู้เสียภาษี 0105529030059 สำนักงานใหญ่']


class VendorAddressTests(unittest.TestCase):
    def test_unresolved_address_is_highlighted_in_matched_workbook(self):
        import tempfile
        from pathlib import Path
        from openpyxl import Workbook, load_workbook
        from Matching_TextFile import highlight_excel_errors
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'result.xlsx'
            wb = Workbook()
            ws = wb.active
            ws.append(['VendorAddress', 'VendorBranch', 'Address', 'Emessage'])
            ws.append(['', '00000', 'Buyer address',
                       'VendorAddress could not be resolved for VendorBranch'])
            wb.save(path)
            highlight_excel_errors(path)
            result = load_workbook(path)
            ws = result.active
            self.assertEqual(ws['A2'].fill.patternType, 'solid')
            self.assertEqual(ws['B2'].fill.patternType, 'solid')
            self.assertIsNone(ws['C2'].fill.patternType)
            result.close()

    def test_document_title_on_separate_line_ends_address(self):
        lines = ['สำนักงานใหญ่ ' + HEAD, HEAD_EN, 'TAX INVOICE / INVOICE',
                 'Unrelated invoice text', 'สาขาที่ 2 ' + BRANCH]
        self.assertEqual(select_vendor_address(RAW, '00000', lines),
                         (HEAD + ' ' + HEAD_EN, 'selected'))

    def test_document_title_on_address_line_preserves_postcode(self):
        for heading in ['TAX INVOICE / INVOICE', 'INVOICE', 'ใบกำกับภาษี / ใบส่งของ']:
            lines = ['สำนักงานใหญ่ ' + HEAD, HEAD_EN + ' ' + heading,
                     'Unrelated invoice text', 'สาขาที่ 2 ' + BRANCH]
            self.assertEqual(select_vendor_address(RAW, '00000', lines),
                             (HEAD + ' ' + HEAD_EN, 'selected'))

    def test_document_title_in_inline_branch_address(self):
        self.assertEqual(select_vendor_address(HEAD + ' TAX INVOICE / INVOICE',
                         '00000', ['สำนักงานใหญ่ ' + HEAD + ' TAX INVOICE / INVOICE']),
                         (HEAD, 'selected'))

    def test_single_office_raw_address_also_removes_document_title(self):
        self.assertEqual(select_vendor_address(HEAD_EN + ' TAX INVOICE / INVOICE',
                         '00000', []), (HEAD_EN, 'unchanged'))

    def test_head_office_keeps_only_one_office_both_languages(self):
        address, status = select_vendor_address(RAW, '00000', LINES)
        self.assertEqual((address, status), (HEAD + ' ' + HEAD_EN, 'selected'))
        self.assertNotIn('7/288', address)
        self.assertNotIn('TEL', address)
        self.assertNotIn('370 MOO', address)

    def test_actual_vendor_branch_is_independent_of_buyer_branch(self):
        self.assertEqual(select_vendor_address(RAW, '00002', LINES),
                         (BRANCH + ' ' + BRANCH_EN, 'selected'))

    def test_missing_or_unknown_branch_does_not_return_merged_address(self):
        for code in ['', '00001', '00009']:
            self.assertEqual(select_vendor_address(RAW, code, LINES), ('', 'unresolved'))

    def test_single_office_without_directory_is_unchanged(self):
        self.assertEqual(select_vendor_address(HEAD, '00000', [HEAD]), (HEAD, 'unchanged'))

    def test_separate_label_lines_and_multiline_address(self):
        lines = ['Head Office', HEAD, HEAD_EN, 'Branch No. 2', BRANCH, BRANCH_EN]
        self.assertEqual(select_vendor_address(RAW, '00002', lines),
                         (BRANCH + ' ' + BRANCH_EN, 'selected'))

    def test_no_branch_markers_with_two_addresses_requires_review(self):
        self.assertEqual(select_vendor_address(RAW, '00000', [HEAD, BRANCH]), ('', 'unresolved'))

    def test_conflicting_addresses_for_same_branch_require_review(self):
        lines = ['Head Office ' + HEAD, 'Head Office 777 Test Road 10570']
        self.assertEqual(select_vendor_address(HEAD + ' 777 Test Road 10570', '00000', lines),
                         ('', 'unresolved'))

    def test_address_of_different_vendor_does_not_replace_raw(self):
        self.assertEqual(select_vendor_address('999 Other Road 10570', '00000', LINES),
                         ('999 Other Road 10570', 'unchanged'))

    def test_opposite_office_only_is_not_used_as_selected_branch(self):
        self.assertEqual(select_vendor_address(BRANCH, '00000', LINES), ('', 'unresolved'))


if __name__ == '__main__':
    unittest.main()
