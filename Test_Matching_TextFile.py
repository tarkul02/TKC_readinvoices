
# -*- coding: utf-8 -*-
"""Validate OCR Excel against H records in Summary Invoice Report TXT.

This standalone test does not run OCR and does not read PDF files.
Only lines beginning with H are used. L and T records are ignored.
"""

import os
import re
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from openpyxl import load_workbook

# ============================================================
# Fixed paths: change these paths to match your computer
# ============================================================
EXCEL_PATH = r"C:\testTKC\OpenAI_Invoice_Processing_AP\INPUT\1100\ocr_output.xlsx"
TEXT_PATH = r"C:\testTKC\OpenAI_Invoice_Processing_AP\INPUT\1100\summary_invoice_report_20260620.txt"
OUTPUT_EXCEL_PATH = r"C:\testTKC\OpenAI_Invoice_Processing_AP\INPUT\1100\OutputExcel\ocr_output_matched.xlsx"

VALIDATION_PREFIX = "Summary validation:"


def clean(value):
    return str(value if value is not None else "").strip().lstrip("\ufeff")


def normalize_invoice_key(value):
    """Ignore spaces and case when matching invoice numbers."""
    return re.sub(r"\s+", "", clean(value)).casefold()


def normalize_text(value):
    """Normalize whitespace and case for text comparison."""
    return re.sub(r"\s+", " ", clean(value)).strip().casefold()


def normalize_date(value):
    """Normalize Excel dates and common date strings to DDMMYYYY where possible."""
    if value is None or clean(value) == "":
        return ""

    if isinstance(value, (datetime, date)):
        return value.strftime("%d%m%Y")

    text = clean(value)
    digits = re.sub(r"\D", "", text)

    if len(digits) == 8 and digits.isdigit():
        # Support YYYYMMDD and DDMMYYYY.
        if digits[:4] in ("19", "20"):
            try:
                return datetime.strptime(
                    digits, "%Y%m%d"
                ).strftime("%d%m%Y")
            except ValueError:
                pass

        try:
            return datetime.strptime(
                digits, "%d%m%Y"
            ).strftime("%d%m%Y")
        except ValueError:
            pass

    for fmt in (
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%Y-%m-%d",
        "%d/%m/%y",
        "%d-%m-%y",
    ):
        try:
            return datetime.strptime(text, fmt).strftime("%d%m%Y")
        except ValueError:
            continue

    return normalize_text(text)


def normalize_amount(value):
    """Normalize numeric values to two decimals; tolerate commas/currency symbols."""
    if value is None or clean(value) == "":
        return ""

    if isinstance(value, (int, float, Decimal)):
        try:
            return Decimal(str(value)).quantize(Decimal("0.01"))
        except InvalidOperation:
            return normalize_text(value)

    text = clean(value).replace(",", "")
    text = re.sub(r"[^\d.()\-]", "", text)

    if text.startswith("(") and text.endswith(")"):
        text = "-" + text[1:-1]

    try:
        return Decimal(text).quantize(Decimal("0.01"))
    except (InvalidOperation, ValueError):
        return normalize_text(value)


def extract_vendor_branch(supplier_field):
    """Extract the text inside the final parentheses."""
    matches = re.findall(r"\(([^()]*)\)", clean(supplier_field))
    return matches[-1].strip() if matches else ""


def supplier_name_without_branch(supplier_field):
    """Remove a trailing branch in parentheses from supplier name."""
    text = clean(supplier_field)
    text = re.sub(r"\s*\([^()]*\)\s*$", "", text)
    return re.sub(r"\s+", " ", text).strip()


def parse_summary_h_records(txt_path):
    """Read only H lines. Field positions are 1-based."""
    records = {}

    with open(
        txt_path,
        "r",
        encoding="utf-8-sig",
        errors="replace"
    ) as report_file:

        for line_no, raw_line in enumerate(report_file, 1):
            line = raw_line.rstrip("\r\n")

            # Ignore all lines not beginning with H.
            if not line.startswith("H"):
                continue

            fields = [clean(item) for item in line.split("|")]

            if not fields or not fields[0].startswith("H"):
                continue

            # Field 6: TaxInvoiceNo
            invoice_no = fields[5] if len(fields) >= 6 else ""
            key = normalize_invoice_key(invoice_no)

            if not key:
                continue

            # Field 20: SupplierName + VendorBranch
            supplier_field = fields[19] if len(fields) >= 20 else ""

            record = {
                "LineNo": line_no,
                "H": line,
                "TaxInvoiceNo": invoice_no,

                # Field 3
                "InvoiceDate": fields[2] if len(fields) >= 3 else "",

                # Field 20, separated into name and branch
                "SupplierName": supplier_name_without_branch(
                    supplier_field
                ),
                "VendorBranch": extract_vendor_branch(
                    supplier_field
                ),

                # Field 16
                "TotalAmount": fields[15] if len(fields) >= 16 else "",

                # Field 17
                "VATAmount": fields[16] if len(fields) >= 17 else "",

                # Field 14
                "AmountIncVat": fields[13] if len(fields) >= 14 else "",
            }

            records.setdefault(key, []).append(record)

    return records


def validate_value(excel_value, report_value, field_name):
    """Return a mismatch description, or an empty string if values match."""

    if field_name == "InvoiceDate":
        left = normalize_date(excel_value)
        right = normalize_date(report_value)

    elif field_name in (
        "TotalAmount",
        "VATAmount",
        "AmountIncVat",
    ):
        left = normalize_amount(excel_value)
        right = normalize_amount(report_value)

    elif field_name == "VendorBranch":
        left = re.sub(
            r"\s+", "", clean(excel_value)
        ).casefold()

        right = re.sub(
            r"\s+", "", clean(report_value)
        ).casefold()

    else:
        left = normalize_text(excel_value)
        right = normalize_text(report_value)

    if left == right:
        return ""

    return (
        f"{field_name} mismatch "
        f"(Excel='{clean(excel_value)}' / "
        f"TXT='{clean(report_value)}')"
    )


def validate_excel_against_summary(
    excel_path,
    text_path,
    output_path
):
    if not os.path.isfile(excel_path):
        raise FileNotFoundError(
            f"ไม่พบ Excel: {excel_path}"
        )

    if not os.path.isfile(text_path):
        raise FileNotFoundError(
            f"ไม่พบ Summary TXT: {text_path}"
        )

    output_dir = os.path.dirname(output_path)

    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    # Read H records from Summary TXT only.
    report_records = parse_summary_h_records(text_path)

    wb = load_workbook(excel_path)
    ws = wb.active

    # Map Excel header names to column numbers.
    column_map = {
        str(cell.value).strip(): cell.column
        for cell in ws[1]
        if cell.value is not None
    }

    required_columns = [
        "TaxInvoiceNo",
        "InvoiceDate",
        "SupplierName",
        "VendorBranch",
        "TotalAmount",
        "VATAmount",
        "AmountIncVat",
        "Emessage",
    ]

    missing_columns = [
        name
        for name in required_columns
        if name not in column_map
    ]

    if missing_columns:
        wb.close()
        raise RuntimeError(
            "ไม่พบคอลัมน์ใน Excel: "
            + ", ".join(missing_columns)
        )

    counts = {
        "MATCH": 0,
        "MISMATCH": 0,
        "NOT FOUND": 0,
        "DUPLICATE": 0,
        "EMPTY INVOICE": 0,
    }

    # Process Excel rows; row 1 is the header.
    for row_num in range(2, ws.max_row + 1):

        emessage_cell = ws.cell(
            row=row_num,
            column=column_map["Emessage"]
        )

        existing_message = clean(emessage_cell.value)

        # Preserve other messages and remove this validator's old result.
        other_messages = [
            part.strip()
            for part in existing_message.split("; ")
            if part.strip()
            and not part.strip().startswith(VALIDATION_PREFIX)
        ]

        invoice_no = ws.cell(
            row=row_num,
            column=column_map["TaxInvoiceNo"]
        ).value

        key = normalize_invoice_key(invoice_no)
        errors = []

        if not key:
            errors.append("TaxInvoiceNo is empty")
            counts["EMPTY INVOICE"] += 1

        else:
            candidates = report_records.get(key, [])

            if not candidates:
                errors.append(
                    "TaxInvoiceNo not found in Summary TXT "
                    f"('{clean(invoice_no)}')"
                )
                counts["NOT FOUND"] += 1

            else:
                # Use the first H record for field comparisons.
                record = candidates[0]

                if len(candidates) > 1:
                    errors.append(
                        "TaxInvoiceNo duplicated in Summary TXT "
                        f"({len(candidates)} H records)"
                    )
                    counts["DUPLICATE"] += 1

                # Compare only the requested fields.
                # L and T records are not used.
                fields_to_compare = (
                    "InvoiceDate",
                    "SupplierName",
                    "VendorBranch",
                    "TotalAmount",
                    "VATAmount",
                    "AmountIncVat",
                )

                for field_name in fields_to_compare:
                    excel_value = ws.cell(
                        row=row_num,
                        column=column_map[field_name]
                    ).value

                    mismatch = validate_value(
                        excel_value,
                        record[field_name],
                        field_name
                    )

                    if mismatch:
                        errors.append(mismatch)

                if errors and len(candidates) == 1:
                    counts["MISMATCH"] += 1

                elif not errors:
                    counts["MATCH"] += 1

        # Write validation status to Emessage.
        if errors:
            validation_message = (
                f"{VALIDATION_PREFIX} FAIL - "
                + " | ".join(errors)
            )
        else:
            validation_message = (
                f"{VALIDATION_PREFIX} PASS"
            )

        all_messages = other_messages + [validation_message]
        emessage_cell.value = "; ".join(all_messages)

    # Save to a separate output file.
    wb.save(output_path)
    wb.close()

    print("\n===== Summary H-line validation result =====")
    print(f"Excel input : {excel_path}")
    print(f"TXT input   : {text_path}")
    print(f"Excel output: {output_path}")

    total_h_records = sum(
        len(items) for items in report_records.values()
    )

    print(f"H records in TXT: {total_h_records}")

    for status, count in counts.items():
        print(f"{status:14}: {count}")


if __name__ == "__main__":
    validate_excel_against_summary(
        EXCEL_PATH,
        TEXT_PATH,
        OUTPUT_EXCEL_PATH
    )
