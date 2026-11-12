"""Offline regression checks for invoice address and Master matching."""
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from Matching_TextFile import (
    address_matches_master, load_master_supplier, supplier_matching_errors,
    append_emessage, highlight_excel_errors,
    address_matching_errors, validate_empty_branches,
)


class SupplierMatchingTests(unittest.TestCase):
    def setUp(self):
        self.rows = [{"SupplierName_Master_EN": "Vendor One", "SupplierName_Master_TH": ""}]

    def test_master_mismatch_even_when_ocr_matches_text(self):
        self.assertEqual(supplier_matching_errors("Other", self.rows, "Other"),
                         ["SupplierName is NOT MATCHING with Master Report"])
        self.assertEqual(supplier_matching_errors("Vendor One", [], "Vendor One"),
                         ["SupplierName is NOT MATCHING with Master Report"])

    def test_master_then_text_and_missing_names(self):
        self.assertEqual(supplier_matching_errors("Vendor One", self.rows, "Vendor One"), [])
        for text in ["Other", ""]:
            self.assertEqual(supplier_matching_errors("Vendor One", self.rows, text),
                             ["SupplierName is MATCHING with Master Report but NOT MATCHING with Text Report"])
        self.assertEqual(supplier_matching_errors("", self.rows, ""),
                         ["SupplierName is NOT MATCHING with Master Report"])

    def test_text_cannot_use_an_unmatched_master_row(self):
        rows = self.rows + [{"SupplierName_Master_EN": "Other", "SupplierName_Master_TH": ""}]
        self.assertEqual(supplier_matching_errors("Vendor One", rows, "Other"),
                         ["SupplierName is MATCHING with Master Report but NOT MATCHING with Text Report"])

    def test_name_match_does_not_require_separate_language_statuses(self):
        rows = [{"SupplierName_Master_EN": "Vendor One", "SupplierName_Master_TH": "บริษัททดสอบ"}]
        self.assertEqual(supplier_matching_errors("บริษัททดสอบ Other", rows, "บริษัททดสอบ"), [])

    def test_rerun_cleanup_and_highlight(self):
        wb = Workbook()
        ws = wb.active
        ws.append(["SupplierName", "Emessage"])
        ws.append(["Vendor One", "Original OCR warning | SupplierName OCR vs Master: TH N/A, EN NOT MATCHING"])
        message = "SupplierName is MATCHING with Master Report but NOT MATCHING with Text Report"
        append_emessage(ws['B2'], [message])
        append_emessage(ws['B2'], [message])
        self.assertEqual(ws['B2'].value, "Original OCR warning | " + message)
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'output.xlsx'
            wb.save(path)
            highlight_excel_errors(path)
            result = load_workbook(path)
            self.assertEqual(result.active['A2'].fill.patternType, 'solid')
            result.close()
        append_emessage(ws['B2'], [])
        self.assertEqual(ws['B2'].value, "Original OCR warning")
        wb.close()

    def test_both_english_errors_highlight_only_related_fields(self):
        messages = [
            "SupplierName is NOT MATCHING with Master Report",
            "SupplierName is MATCHING with Master Report but NOT MATCHING with Text Report",
        ]
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'highlight.xlsx'
            wb = Workbook()
            ws = wb.active
            ws.append(['SupplierName', 'Emessage', 'VendorAddress', 'Address'])
            for message in messages:
                ws.append(['Vendor', message, '123', '456'])
            ws.append(['Vendor', None, '123', '456'])
            ws.append(['Vendor', 'Supplier Address is NOT MATCHING | Name Address is NOT MATCHING', '123', '456'])
            wb.save(path)
            wb.close()
            highlight_excel_errors(path)
            result = load_workbook(path)
            ws = result.active
            for row in (2, 3):
                for column in ('A', 'B'):
                    self.assertEqual(ws[f'{column}{row}'].fill.fgColor.rgb, '00FFEB9C')
                    self.assertEqual(ws[f'{column}{row}'].font.color.rgb, '00000000')
                for column in ('C', 'D'):
                    self.assertIsNone(ws[f'{column}{row}'].fill.patternType)
            self.assertIsNone(ws['A4'].fill.patternType)
            self.assertIsNone(ws['B4'].fill.patternType)
            for column in ('B', 'C', 'D'):
                self.assertEqual(ws[f'{column}5'].fill.fgColor.rgb, '00FFEB9C')
            self.assertIsNone(ws['A5'].fill.patternType)
            result.close()


class AddressMatchingTests(unittest.TestCase):
    def test_name_address_ignores_commas_in_comparison(self):
        rows = [{'BranchCode': '00000', 'StreetNo': '1,234 Road', 'PostCode': '10,570'}]
        self.assertEqual(address_matching_errors('1234 Road 10,570', rows, '00000'), [])
        self.assertEqual(address_matching_errors('1234 Road 10,571', rows, '00000'),
                         ['Name Address is NOT MATCHING'])
        self.assertEqual(address_matching_errors('1234 Road 10,570', rows, '00000', supplier=True),
                         ['Supplier Address is NOT MATCHING'])
        rows = [{'BranchCode': '', 'StreetNo': '73, 73/2, 131 ม.7', 'PostCode': '74110'}]
        self.assertEqual(address_matching_errors('73,\t73/2, 131 ม. 7 74110,', rows, None), [])
        self.assertEqual(address_matching_errors('73,73/3,131 ม.7 74110', rows, None),
                         ['Name Address is NOT MATCHING'])

    def test_comma_separated_house_numbers_ignore_comma_spacing(self):
        master = {'StreetNo': '73, 73/2, 131 ม.7', 'PostCode': '74110'}
        address = '73,73/2,131 ม. 7 ถ.สวนหลวงร่วมใจ ต.สวนหลวง อ.กระทุ่มแบน จ.สมุทรสาคร 74110'
        self.assertTrue(address_matches_master(address, master))
        for street in ['73,73/2,131 ม.7', '73 , 73/2 , 131 ม.7', '73,\t73/2,\t131 ม.7']:
            with self.subTest(street=street):
                self.assertTrue(address_matches_master(address, dict(master, StreetNo=street)))
        self.assertFalse(address_matches_master(address.replace('73/2', '73/3'), master))
        self.assertFalse(address_matches_master(address.replace('74110', '74111'), master))
        self.assertFalse(address_matches_master('73 ม.7 74110', master))

    def test_empty_branch_validation_without_master_branch(self):
        record = {'Branch': '', 'TaxID': '123', 'VendorBranch': None, 'VendorTaxId': '123'}
        for master in ({}, {'123': [{'BranchCode': ''}]}):
            self.assertEqual(validate_empty_branches(record, master),
                             ['Branch is empty', 'VendorBranch is empty'])
        self.assertEqual(validate_empty_branches({'Branch': '00000', 'VendorBranch': 0}, {}), [])

    def test_branch_selection_and_missing_branch(self):
        rows = [
            {'BranchCode': '00000', 'StreetNo': '370 Road', 'PostCode': '10570'},
            {'BranchCode': '00001', 'StreetNo': '555 Road', 'PostCode': '25140'},
        ]
        for branch in ['00001', '1', 1, 1.0]:
            self.assertEqual(address_matching_errors('555 Road 25140', rows, branch), [])
        self.assertEqual(address_matching_errors('555 Road 25140', rows, '00000'),
                         ['Name Address is NOT MATCHING'])
        self.assertEqual(address_matching_errors('555 Road 25140', rows, '00002', True),
                         ['Supplier Address is NOT MATCHING'])
        self.assertEqual(address_matching_errors('555 Road 25140', rows, None),
                         [])
        self.assertEqual(address_matching_errors('555 Road 25140', rows, '', True),
                         [])
        rows = [{'BranchCode': '', 'StreetNo': '555', 'PostCode': '25140'}]
        self.assertEqual(address_matching_errors('555 Road 25140', rows, None, True), [])

    def test_master_continuation_rows_keep_both_branches(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'master.xlsx'
            wb = Workbook()
            ws = wb.active
            ws.append(['Header'] * 9)
            ws.append([None, None, None, '123', 'Vendor', 'บริษัท', '00000', '370 Road', 10570])
            ws.append([None] * 7 + ['District', None])
            ws.append([None] * 7 + ['Province', None])
            ws.append([None] * 6 + ['00001', '555 Road', 25140])
            ws.append([None] * 7 + ['Other District', None])
            ws.append(['Header'] * 9)
            ws.append([None] * 6 + ['00002', 'Unrelated', 99999])
            wb.save(path)
            wb.close()
            rows = load_master_supplier(path)['123']
            self.assertEqual(len(rows), 2)
            self.assertEqual(rows[0]['StreetNo'], '370 Road District Province')
            self.assertEqual(rows[1]['SupplierName_Master_EN'], 'Vendor')
            self.assertEqual(rows[1]['BranchCode'], '00001')
            self.assertEqual(address_matching_errors('555 Road 25140', rows, '00001'), [])

    def test_missing_branch_message_replaces_old_message_and_highlights(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'output.xlsx'
            wb = Workbook()
            ws = wb.active
            ws.append(['Address', 'Branch', 'VendorAddress', 'VendorBranch', 'Emessage'])
            ws.append(['370', None, '555', None, 'Branch is empty | VendorBranch is empty | Original warning'])
            messages = validate_empty_branches(
                {'Branch': None, 'TaxID': '123', 'VendorBranch': '', 'VendorTaxId': '123'},
                {'123': [{'BranchCode': '00001'}]},
            )
            self.assertEqual(messages, ['Branch is empty; Address matching skipped',
                                        'VendorBranch is empty; Supplier Address matching skipped'])
            append_emessage(ws['E2'], messages)
            append_emessage(ws['E2'], messages)
            self.assertEqual(ws['E2'].value, 'Original warning | ' + ' | '.join(messages))
            wb.save(path)
            wb.close()
            highlight_excel_errors(path)
            result = load_workbook(path)
            for cell in result.active[2]:
                self.assertEqual(cell.fill.fgColor.rgb, '00FFEB9C')
            result.close()

    def test_street_and_postcode_must_match_same_row(self):
        master = {"StreetNo": "123/4 Main Road", "PostCode": "10110"}
        for address, expected in [
            ("123/4 Road 10110", True),
            ("123/4 Road 10111", False),
            ("999 Road 10110", False),
            ("1234 Road 10110", False),
            ("123.4 Road 10110", False),
            (" \t123/4 Road 10 1\t10 \t", True),
            ("123/4 Road ABCDE", True),
            ("999 Road ABCDE", False),
            ("123/4 Road 1234", True),
        ]:
            with self.subTest(address=address):
                self.assertEqual(address_matches_master(address, master), expected)
        rows = [master, {"StreetNo": "999", "PostCode": "10111"}]
        self.assertFalse(any(address_matches_master("123/4 Road 10111", m) for m in rows))

    def test_missing_fields_only_compare_available_pairs(self):
        self.assertTrue(address_matches_master("123 Road 10110", {"StreetNo": "123", "PostCode": ""}))
        self.assertTrue(address_matches_master("", {"StreetNo": "123", "PostCode": "10110"}))
        self.assertFalse(address_matches_master("123 Road 10111", {"StreetNo": "", "PostCode": "10110"}))

    def test_master_column_nine_and_duplicate_tax_rows(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "master.xlsx"
            wb = Workbook()
            wb.active.append(["unused"] * 9)
            wb.active.append([None, None, None, "123", "Vendor", "", "0", "123 Road", 10110])
            wb.active.append([None, None, None, "123", "Vendor", "", "0", "999 Road", "00110"])
            wb.save(path)
            wb.close()
            rows = load_master_supplier(path)["123"]
            self.assertEqual([r["PostCode"] for r in rows], ["10110", "00110"])


if __name__ == "__main__":
    unittest.main()
