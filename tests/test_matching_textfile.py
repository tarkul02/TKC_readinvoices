"""Offline regression checks for invoice address and Master matching."""
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from Matching_TextFile import (
    address_matches_master, load_master_supplier, supplier_matching_errors,
    append_emessage, highlight_excel_errors,
)


class SupplierMatchingTests(unittest.TestCase):
    def setUp(self):
        self.rows = [{"SupplierName_Master_EN": "Vendor One", "SupplierName_Master_TH": ""}]

    def test_master_mismatch_even_when_ocr_matches_text(self):
        self.assertEqual(supplier_matching_errors("Other", self.rows, "Other"),
                         ["SupplierName is NOT MATCHING with Master file"])
        self.assertEqual(supplier_matching_errors("Vendor One", [], "Vendor One"),
                         ["SupplierName is NOT MATCHING with Master file"])

    def test_master_then_text_and_missing_names(self):
        self.assertEqual(supplier_matching_errors("Vendor One", self.rows, "Vendor One"), [])
        for text in ["Other", ""]:
            self.assertEqual(supplier_matching_errors("Vendor One", self.rows, text),
                             ["SupplierName is MATCHING with Master file but NOT MATCHING with Text file"])
        self.assertEqual(supplier_matching_errors("", self.rows, ""),
                         ["SupplierName is NOT MATCHING with Master file"])

    def test_text_cannot_use_an_unmatched_master_row(self):
        rows = self.rows + [{"SupplierName_Master_EN": "Other", "SupplierName_Master_TH": ""}]
        self.assertEqual(supplier_matching_errors("Vendor One", rows, "Other"),
                         ["SupplierName is MATCHING with Master file but NOT MATCHING with Text file"])

    def test_name_match_does_not_require_separate_language_statuses(self):
        rows = [{"SupplierName_Master_EN": "Vendor One", "SupplierName_Master_TH": "บริษัททดสอบ"}]
        self.assertEqual(supplier_matching_errors("บริษัททดสอบ Other", rows, "บริษัททดสอบ"), [])

    def test_rerun_cleanup_and_highlight(self):
        wb = Workbook()
        ws = wb.active
        ws.append(["SupplierName", "Emessage"])
        ws.append(["Vendor One", "Original OCR warning | SupplierName OCR vs Master: TH N/A, EN NOT MATCHING"])
        message = "SupplierName is MATCHING with Master file but NOT MATCHING with Text file"
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
            "SupplierName is NOT MATCHING with Master file",
            "SupplierName is MATCHING with Master file but NOT MATCHING with Text file",
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
