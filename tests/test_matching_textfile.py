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

    def test_rerun_removes_stale_highlights_and_reapplies_current_errors(self):
        from openpyxl.styles import PatternFill

        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "rerun.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.append(["VendorAddress", "Address", "Emessage", "Assignment"])
            ws.append(["1/23 Road 24130", "370 Road 10570",
                       "Supplier Address is NOT MATCHING", "invoice.pdf"])
            ws.append(["123", "456", None, "other.pdf"])
            ws['B3'].fill = PatternFill(fill_type="solid", fgColor="00FF00")
            ws['D2'].fill = PatternFill(fill_type="solid", fgColor="FFEB9C")
            wb.save(path)
            wb.close()
            highlight_excel_errors(path)
            wb = load_workbook(path)
            ws = wb.active
            self.assertEqual(ws['A2'].fill.fgColor.rgb, '00FFEB9C')
            append_emessage(ws['C2'], ["Name Address is NOT MATCHING"])
            wb.save(path)
            wb.close()
            highlight_excel_errors(path)
            wb = load_workbook(path)
            ws = wb.active
            self.assertIsNone(ws['A2'].fill.patternType)
            self.assertEqual(ws['B2'].fill.fgColor.rgb, '00FFEB9C')
            self.assertEqual(ws['C2'].fill.fgColor.rgb, '00FFEB9C')
            append_emessage(ws['C2'], [])
            wb.save(path)
            wb.close()
            highlight_excel_errors(path)
            wb = load_workbook(path)
            ws = wb.active
            self.assertIsNone(ws['C2'].value)
            for column in ('A', 'B', 'C'):
                self.assertIsNone(ws[f'{column}2'].fill.patternType)
            self.assertEqual(ws['B3'].fill.fgColor.rgb, '0000FF00')
            self.assertEqual(ws['D2'].fill.fgColor.rgb, '00FFEB9C')
            wb.close()

    def test_retired_company_address_error_is_removed_on_rerun(self):
        wb = Workbook()
        cell = wb.active['A1']
        cell.value = "Company address does not match | PO does not match Company branch"
        append_emessage(cell, ["Name Address is NOT MATCHING"])
        self.assertEqual(cell.value,
                         "PO does not match Company branch | Name Address is NOT MATCHING")
        append_emessage(cell, [])
        self.assertEqual(cell.value, "PO does not match Company branch")
        wb.close()

    def test_vendor_branch_text_mismatch_skips_empty_ocr_branch(self):
        from unittest.mock import patch
        from Matching_TextFile import validate_excel

        headers = ["TaxInvoiceNo", "InvoiceDate", "SupplierName", "VendorTaxId",
                   "VendorBranch", "VendorAddress", "Address", "TotalAmount",
                   "VATAmount", "AmountIncVat", "Emessage"]
        master = {"123": [{"SupplierName_Master_EN": "Vendor",
                           "SupplierName_Master_TH": "", "BranchCode": "",
                           "StreetNo": "37", "PostCode": "10570"}]}
        text = {"inv": [{"InvoiceDate": "2026-06-17", "SupplierName": "Vendor",
                          "VendorBranch": "00000", "TotalAmount": 100,
                          "VATAmount": 7, "AmountIncVat": 107}]}
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / "input.xlsx"
            output = Path(folder) / "output.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.append(headers)
            for branch in (None, "   ", "00001", "00000"):
                ws.append(["INV", "2026-06-17", "Vendor", "123", branch,
                           "37 Road 10570", "", 100, 7, 107,
                           "VendorBranch is NOT MATCHING"])
            wb.save(source)
            wb.close()
            with patch("Matching_TextFile.load_summary_h_records", return_value=text), \
                 patch("Matching_TextFile.load_master_supplier", return_value=master):
                validate_excel(source, source, source, output)
            highlight_excel_errors(output)
            wb = load_workbook(output)
            ws = wb.active
            self.assertEqual(ws['K2'].value, "VendorBranch is empty")
            self.assertEqual(ws['K3'].value, "VendorBranch is empty")
            self.assertEqual(ws['K4'].value, "VendorBranch is NOT MATCHING with Text Report")
            self.assertEqual(ws['E4'].fill.fgColor.rgb, '00FFEB9C')
            self.assertIsNone(ws['K5'].value)
            wb.close()


class InvoicePrefixTests(unittest.TestCase):
    def test_oth_numeric_prefix_fallback_for_uploaded_invoices(self):
        from Matching_TextFile import find_summary_invoice_records, normalize_invoice
        for number in ('2606200023', '2606200022', '2606200019', '2606190094'):
            records = [{'TaxInvoiceNo': number}]
            with self.subTest(number=number):
                self.assertEqual(find_summary_invoice_records(
                    normalize_invoice('OTH' + number), {number: records}), records)

    def test_exact_match_has_priority_and_other_numbers_are_unchanged(self):
        from Matching_TextFile import find_summary_invoice_records, normalize_invoice
        exact = [{'TaxInvoiceNo': 'OTH123'}]
        fallback = [{'TaxInvoiceNo': '123'}]
        records = {'oth123': exact, '123': fallback}
        self.assertEqual(find_summary_invoice_records(normalize_invoice('OTH123'), records), exact)
        self.assertEqual(find_summary_invoice_records('123', records), fallback)
        for number in ('abc123', 'xoth123', 'othabc123', 'oth', 'oth-123', ''):
            with self.subTest(number=number):
                self.assertEqual(find_summary_invoice_records(number, records), [])


class AddressMatchingTests(unittest.TestCase):
    def test_leading_ocr_quote_does_not_invent_missing_digits(self):
        rows = [{'BranchCode': '', 'StreetNo': '37 Road', 'PostCode': '10570'}]
        for quote in ['‘', '’', "'", '“', '”', '"']:
            with self.subTest(quote=quote):
                self.assertEqual(address_matching_errors(quote + '37 Road 10570', rows, None, True), [])
        rows[0]['StreetNo'] = '370 Road'
        self.assertEqual(address_matching_errors('‘37 Road 10570', rows, None, True),
                         ['Supplier Address is NOT MATCHING'])

    def test_trailing_house_comma_does_not_join_following_word(self):
        rows = [{'BranchCode': '00000', 'StreetNo': '370 หมู่ 17', 'PostCode': '10570'}]
        for address in ['370, MOO 17 10570', '370, หมู่ 17 10570', '370,\t MOO 17 10570']:
            with self.subTest(address=address):
                self.assertEqual(address_matching_errors(address, rows, '00000'), [])
        self.assertEqual(address_matching_errors('370, MOO 17 10571', rows, '00000'),
                         ['Name Address is NOT MATCHING'])

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

    def test_master_tokens_can_appear_anywhere_in_ocr(self):
        master = {"StreetNo": "37 Main Road", "PostCode": "10570 extra"}
        for supplier in (False, True):
            for address, expected in [
                ("Building A เลขที่37 Road 10570 Thailand", True),
                ("10570 Thailand Building A 37 Road", True),
                ("Building A 370 Road 10570 Thailand", False),
                ("Building A 137 Road 10570 Thailand", False),
                ("370 Road 10570 Thailand; 37 Road", True),
                ("Building A 37 Road 10571 Thailand", False),
                ("Building A 37 Road Thailand", True),
                ("Building A 37 Road 110570 Thailand", True),
                ("Building A 37 Road 110570 and 10571", False),
                ("Building A 37 Road 10571 and 10570 Thailand", True),
            ]:
                with self.subTest(supplier=supplier, address=address):
                    errors = address_matching_errors(
                        address, [dict(master, BranchCode="00001")], "1", supplier
                    )
                    self.assertEqual(not errors, expected)

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
