import os

import shutil

from invoice_email import send_invoice_email

import re

from decimal import Decimal, InvalidOperation

from datetime import datetime, date

from openpyxl import load_workbook

from openpyxl.styles import PatternFill, Font

# ============================================================

# FILE PATHS

# ============================================================

EXCEL_PATH = (

    # r"C:\TKC\TKC_readinvoices\INPUT\Process3\0000\OutputExcel\text\Excel\OCR\All_Invoices.xlsx"

    r"C:\TKC\TKC_readinvoices\INPUT\Process3\0000\OutputExcel\20261101_164630\All_Invoices.xlsx"

)

TEXT_PATH = (

    r"C:\TKC\TKC_readinvoices\INPUT\Process3\0000\OutputExcel\20261101_164630\summary invoice report_20260620.txt"

)

MASTER_PATH = (

    r"C:\TKC\TKC_readinvoices\MasterSuplier\Master Vendor Name (MAT)_TH_ENG.xlsx"

)

OUTPUT_EXCEL_PATH = (

    r"C:\TKC\TKC_readinvoices\INPUT\Process3\0000\OutputExcel\20261101_164630\ocr_output_matched_ุ6.xlsx"

)

# ============================================================

# NORMALIZE FUNCTIONS

# ============================================================

def clean(value):

    if value is None:

        return ""

    return str(value).strip().lstrip("\ufeff")

def normalize_invoice(value):

    return re.sub(r"\s+", "", clean(value)).casefold()

def normalize_tax_id(value):

    return re.sub(r"\D", "", clean(value))

def normalize_text(value):

    return re.sub(

        r"\s+", " ", clean(value)

    ).strip().casefold()

def normalize_branch(value):

    return re.sub(

        r"\s+", "", clean(value)

    ).casefold()

def normalize_date(value):

    #Excel format: dd/mm/yyyy

    #TXT format: ddmmyyyy

    if value is None or clean(value) == "":

        return ""

    if isinstance(value, (datetime, date)):

        return value.strftime("%d%m%Y")

    text = clean(value)

    for fmt in (

        "%d/%m/%Y",

        "%d-%m-%Y",

        "%Y-%m-%d",

        "%d%m%Y",

    ):

        try:

            return datetime.strptime(

                text, fmt

            ).strftime("%d%m%Y")

        except ValueError:

            pass

    return re.sub(r"\D", "", text)

def normalize_amount(value):

    if value is None or clean(value) == "":

        return None

    if isinstance(value, (int, float, Decimal)):

        try:

            return Decimal(str(value)).quantize(

                Decimal("0.01")

            )

        except InvalidOperation:

            return None

    text = clean(value).replace(",", "")

    text = re.sub(r"[^\d.()-]", "", text)

    if text.startswith("(") and text.endswith(")"):

        text = "-" + text[1:-1]

    try:

        return Decimal(text).quantize(

            Decimal("0.01")

        )

    except (InvalidOperation, ValueError):

        return None

# Name matching helpers: remove punctuation, separate Thai/English text.
def normalize_company_name(value):
    return re.sub(r"\s+", " ", clean(value).replace(".", "").replace(",", "")).strip().casefold()


def language_part(value, language):
    text = clean(value)
    if language == "th":
        text = re.sub(r"[^\u0e00-\u0e7f\s]", " ", text)
    else:
        text = re.sub(r"[^A-Za-z\s]", " ", text)
    return normalize_company_name(text)


def supplier_names_match(left, right):
    # Accept a matching Thai OR English component when both are present.
    if not clean(left) or not clean(right):
        return False
    if normalize_company_name(left) == normalize_company_name(right):
        return True
    return any(
        language_part(left, lang) == language_part(right, lang)
        for lang in ("th", "en")
        if language_part(left, lang) and language_part(right, lang)
    )


def supplier_language_status(ocr_name, reference_name, language):
    """Compare one language; missing language is distinct from a mismatch."""
    ocr_part = language_part(ocr_name, language)
    reference_part = language_part(reference_name, language)
    if not ocr_part and not reference_part:
        return "N/A"
    if not ocr_part:
        return "NOT FOUND IN OCR"
    if not reference_part:
        return "NOT FOUND IN REFERENCE"
    return "MATCHED" if ocr_part == reference_part else "NOT MATCHING"


def supplier_language_errors(ocr_name, reference_names, source):
    """Return one detailed error only when an available language fails."""
    statuses = []
    for reference_name in reference_names:
        th = supplier_language_status(ocr_name, reference_name, "th")
        en = supplier_language_status(ocr_name, reference_name, "en")
        statuses.append((th, en))
    if not statuses:
        return []
    # Multiple master rows for one tax ID: select the best matching row.
    def score(pair):
        return (sum(x == "MATCHED" for x in pair),
                -sum(x not in ("MATCHED", "N/A") for x in pair))
    th, en = max(statuses, key=score)
    if all(x in ("MATCHED", "N/A") for x in (th, en)):
        return []
    return [f"SupplierName OCR vs {source}: TH {th}, EN {en}"]


def first_address_word(value):
    parts = clean(value).split()
    return normalize_company_name(parts[0]) if parts else ""


def company_name_matches(name, master_en, master_th):
    if not clean(name):
        return False
    th = language_part(name, "th")
    en = language_part(name, "en")
    if th and supplier_names_match(th, master_th):
        return True
    # For English 'Name', accept THAI KOITO regardless of differing legal suffixes.
    if en and "thai koito" in en and "thai koito" in language_part(master_en, "en"):
        return True
    return bool(en and supplier_names_match(en, master_en))

REQUIRED_FIELDS = [

    "TaxInvoiceNo",

    "InvoiceDate",

    "SupplierName",

    "TotalAmount",

    "VATAmount",

    "AmountIncVat",

    "VendorTaxId",

    "VendorBranch",

    "PurchaseOrderNo",

    "TaxRemark",

    "TaxID",

]

# ============================================================

# SUMMARY TXT FUNCTIONS

# ============================================================

def load_summary_h_records(txt_path):

    """

    Read only H records.

    Ignore L and T records.

    Field positions:

        Field 3  = InvoiceDate

        Field 6  = TaxInvoiceNo

        Field 14 = AmountIncVat

        Field 16 = TotalAmount

        Field 17 = VATAmount

        Field 20 = SupplierName + VendorBranch

    """

    records = {}

    with open(

        txt_path,

        "r",

        encoding="utf-8-sig",

        errors="replace"

    ) as file:

        for line_number, line in enumerate(file, 1):

            line = line.rstrip("\r\n")

            # Only process H records.

            if not line.startswith("H"):

                continue

            fields = [

                clean(item)

                for item in line.split("|")

            ]

            if len(fields) < 20:

                continue

            invoice_no = fields[5]

            invoice_key = normalize_invoice(invoice_no)

            if not invoice_key:

                continue

            supplier_field = fields[19]

            # Extract branch from parentheses in Field 20.

            branch_matches = re.findall(

                r"\(([^()]*)\)",

                supplier_field

            )

            vendor_branch = (

                branch_matches[-1].strip()

                if branch_matches else ""

            )

            # Remove trailing branch from supplier name.

            supplier_name = re.sub(

                r"\s*\([^()]*\)\s*$",

                "",

                supplier_field

            ).strip()

            record = {

                "LineNumber": line_number,

                "InvoiceDate": fields[2],

                "TaxInvoiceNo": fields[5],

                "AmountIncVat": fields[13],

                "TotalAmount": fields[15],

                "VATAmount": fields[16],

                "SupplierName": supplier_name,

                "VendorBranch": vendor_branch,

            }

            records.setdefault(

                invoice_key, []

            ).append(record)

    return records

# ============================================================

# MASTER SUPPLIER FUNCTIONS

# ============================================================

def load_master_supplier(master_path):

    """

    Master Supplier columns:

        Column 4 = VendorTaxId

        Column 5 = SupplierName_Master_EN

        Column 6 = SupplierName_Master_TH

        Column 7 = BranchCode

        Column 8 = StreetNo

    """

    wb = load_workbook(

        master_path,

        read_only=True,

        data_only=True

    )

    ws = wb.active

    master_data = {}

    for row in ws.iter_rows(

        min_row=2,

        values_only=True

    ):

        vendor_tax_id = (

            row[3] if len(row) > 3 else None

        )

        supplier_name_en = (

            row[4] if len(row) > 4 else None

        )

        supplier_name_th = (

            row[5] if len(row) > 5 else None

        )

        BranchCode = (

            row[6] if len(row) > 6 else None

        )

        StreetNo = (

            row[7] if len(row) > 7 else None

        )

        tax_key = normalize_tax_id(vendor_tax_id)

        if not tax_key:

            continue

        master_data.setdefault(

            tax_key, []

        ).append({

            "SupplierName_Master_EN": clean(

                supplier_name_en

            ),

            "SupplierName_Master_TH": clean(

                supplier_name_th

            ),

            "BranchCode": clean(

                BranchCode

            ),

            "StreetNo": clean(

                StreetNo

            )

        })

    wb.close()

    return master_data

# ============================================================

# E-MESSAGE FUNCTION

# ============================================================

# Messages generated by this matching stage. Remove stale versions on reruns,
# but retain every original OCR / earlier validation message.
MATCHING_MESSAGES = {
    "TaxInvoiceNo is NOT FOUND in Summary Report",
    "VendorTaxId NOT FOUND in Master Supplier",
    "InvoiceDate is NOT MATCHING",
    "SupplierName is NOT MATCHING",
    "SupplierName is NOT MATCHING with Master file",
    "SupplierName is NOT MATCHING with Text file",
    "Address is NOT MATCHING",
    "Supplier Address is NOT MATCHING",
    "Name Address is NOT MATCHING",
    "Name is NOT MATCHING with Master file",
    "VendorBranch is NOT MATCHING",
    "TotalAmount is NOT MATCHING",
    "VATAmount is NOT MATCHING",
    "AmountIncVat is NOT MATCHING",
    "TaxRemark is empty",
    "TaxID is empty",
}


def append_emessage(cell, new_messages):
    """Keep pre-existing messages and append unique matching errors with ' | '."""
    existing = clean(cell.value)
    # Remove detailed supplier matching errors from previous runs before recomputing.
    existing = re.sub(
        r"(?:^|\s*\|\s*)SupplierName OCR vs (?:Master|Text file): "
        r"TH (?:MATCHED|NOT MATCHING|NOT FOUND IN OCR|NOT FOUND IN REFERENCE|N/A), "
        r"EN (?:MATCHED|NOT MATCHING|NOT FOUND IN OCR|NOT FOUND IN REFERENCE|N/A)",
        "", existing, flags=re.IGNORECASE,
    )
    # Older files may contain comma-separated matching messages; never split
    # arbitrary original OCR messages on commas.
    for msg in sorted(MATCHING_MESSAGES, key=len, reverse=True):
        existing = re.sub(r"(?<![A-Za-z])" + re.escape(msg) + r"(?![A-Za-z])", "", existing)
    existing = re.sub(r"\s*(?:\|\s*)+", " | ", existing).strip(" |,\t\r\n")
    parts = [part.strip() for part in existing.split(" | ") if part.strip()]
    for message in new_messages:
        if message and message not in parts:
            parts.append(message)
    cell.value = " | ".join(parts) if parts else None

# ==========================================================

# 🎨 Highlight Excel cells that have problems

# ==========================================================

def highlight_excel_errors(excel_path):

    """

    อ่าน Emessage ของแต่ละแถว แล้วทำสีช่องที่เกี่ยวข้องกับปัญหา

    - Emessage ที่มีข้อความ -> ทำสี

    - Required field ที่เป็น "<Field> is empty" -> ทำสี Field นั้น

    - Amount error -> ทำสี TotalAmount, VATAmount, AmountIncVat

    - Company address mismatch -> ทำสี CompanyAddress, CompanyBranch

    - PO/Company branch mismatch -> ทำสี PurchaseOrderNo, CompanyBranch

    - เอกสารสำเนา -> Emessage สีเหลือง (warning)

    """

    wb = load_workbook(excel_path)

    ws = wb.active

    error_fill = PatternFill(fill_type="solid", fgColor="FFEB9C")

    error_font = Font(color="000000")

    warning_fill = PatternFill(fill_type="solid", fgColor="FFEB9C")

    warning_font = Font(color="000000")

    column_map = {}

    for cell in ws[1]:

        if cell.value is not None:

            column_map[str(cell.value).strip()] = cell.column

    emessage_col = column_map.get("Emessage")

    if not emessage_col:

        print("⚠️ ไม่พบ Column Emessage → ข้ามการทำสี")

        wb.save(excel_path)

        return

    def mark_cell(row_number, field_name, warning=False):

        col_number = {"__SUPPLIER_ADDRESS_F__": 6, "__NAME_ADDRESS_O__": 15}.get(field_name, column_map.get(field_name))
        if col_number is None and field_name in ("TaxRemark", "TaxID"):
            col_number = next((col for header, col in column_map.items() if header.casefold() == field_name.casefold()), None)

        if not col_number:

            return

        cell = ws.cell(row=row_number, column=col_number)

        if warning:

            cell.fill = warning_fill

            cell.font = warning_font

        else:

            cell.fill = error_fill

            cell.font = error_font
    
    error_field_map = {
        "Company address does not match": ["CompanyAddress", "CompanyBranch"],
        "PO does not match Company branch": ["PurchaseOrderNo", "CompanyBranch"],
        "Duplicate Amount Found": ["TotalAmount", "VATAmount", "AmountIncVat"],
        "Amount format invalid": ["TotalAmount", "VATAmount", "AmountIncVat"],
        "VATAmount format invalid": ["TotalAmount", "VATAmount", "AmountIncVat"],
        "กรุณาเช็ค Amount ทั้ง 3 ช่อง": ["TotalAmount", "VATAmount", "AmountIncVat"],

        # Matching Summary Report
        "TaxInvoiceNo is NOT FOUND in Summary Report": ["TaxInvoiceNo"],
        "VendorTaxId NOT FOUND in Master Supplier": ["VendorTaxId"],
        "InvoiceDate is NOT MATCHING": ["InvoiceDate"],
        "SupplierName is NOT MATCHING": ["SupplierName"],
        "SupplierName is NOT MATCHING with Master file": ["SupplierName"],
        "SupplierName is NOT MATCHING with Text file": ["SupplierName"],
        "Address is NOT MATCHING": ["Address"],
        "Supplier Address is NOT MATCHING": ["__SUPPLIER_ADDRESS_F__"],
        "Name Address is NOT MATCHING": ["__NAME_ADDRESS_O__"],
        "Name is NOT MATCHING with Master file": ["Name"],
        "VendorBranch is NOT MATCHING": ["VendorBranch"],
        "TotalAmount is NOT MATCHING": ["TotalAmount"],
        "VATAmount is NOT MATCHING": ["VATAmount"],
        "AmountIncVat is NOT MATCHING": ["AmountIncVat"],
    }
    for row_number in range(2, ws.max_row + 1):

        emessage_cell = ws.cell(row=row_number, column=emessage_col)

        emessage = str(emessage_cell.value or "").strip()

        if not emessage:

            continue

        emessage_lower = emessage.casefold()

        copy_warning_only = (

            "เอกสารใบนี้เป็นสำเนา" in emessage

            and " | " not in emessage

        )

        mark_cell(row_number, "Emessage", warning=copy_warning_only)

        # Required Field Error เช่น VendorTaxId is empty

        for field in REQUIRED_FIELDS:

            required_error = f"{field} is empty".casefold()

            if required_error in emessage_lower:

                mark_cell(row_number, field)

        # Error เฉพาะประเภท

        for error_keyword, fields in error_field_map.items():

            message_parts = [x.strip().casefold() for x in emessage.split(" | ")]
            matched_error = (any(x.startswith(error_keyword.casefold()) for x in message_parts)
                             if error_keyword.startswith("SupplierName OCR vs ")
                             else error_keyword.casefold() in message_parts)
            if matched_error:

                for field in fields:

                    mark_cell(row_number, field)

        # Detailed SupplierName messages remain in Emessage, but highlight
        # only if neither Thai nor English matched for a comparison source.
        for message in emessage.split(" | "):
            if not message.startswith(("SupplierName OCR vs Master:", "SupplierName OCR vs Text file:")):
                continue
            status_match = re.fullmatch(
                r"SupplierName OCR vs (?:Master|Text file): TH (MATCHED|NOT MATCHING|NOT FOUND IN OCR|NOT FOUND IN REFERENCE|N/A), EN (MATCHED|NOT MATCHING|NOT FOUND IN OCR|NOT FOUND IN REFERENCE|N/A)",
                message.strip(), flags=re.IGNORECASE,
            )
            if status_match and all(status.upper() != "MATCHED" for status in status_match.groups()):
                mark_cell(row_number, "SupplierName")

    wb.save(excel_path)

    print("🎨 Highlight Error Cells completed")

# ============================================================

# VALIDATION FUNCTION

# ============================================================

def validate_excel(

    excel_path,

    txt_path,

    master_path,

    output_path

):

    # --------------------------------------------------------

    # Check input files

    # --------------------------------------------------------

    for path in (

        excel_path,

        txt_path,

        master_path

    ):

        if not os.path.isfile(path):

            raise FileNotFoundError(

                f"File not found: {path}"

            )

    # --------------------------------------------------------

    # Load TXT H records

    # --------------------------------------------------------

    summary_records = load_summary_h_records(

        txt_path

    )

    # --------------------------------------------------------

    # Load Master Supplier

    # --------------------------------------------------------

    master_records = load_master_supplier(

        master_path

    )

    # --------------------------------------------------------

    # Load OCR Excel

    # --------------------------------------------------------

    wb = load_workbook(excel_path)

    ws = wb.active

    # Map Excel headers.

    column_map = {

        clean(cell.value): cell.column

        for cell in ws[1]

        if cell.value is not None

    }

    required_columns = [

        "TaxInvoiceNo",

        "InvoiceDate",

        "SupplierName",

        "VendorTaxId",

        "VendorBranch",

        "TotalAmount",

        "VATAmount",

        "AmountIncVat",

        "Emessage",

    ]

    missing_columns = [

        col for col in required_columns

        if col not in column_map

    ]

    if missing_columns:

        wb.close()

        raise ValueError(

            "Missing Excel columns: "

            + ", ".join(missing_columns)

        )

    result_count = {

        "PASS": 0,

        "FAIL": 0,

        "TXT_NOT_FOUND": 0,

        "MASTER_NOT_FOUND": 0,

    }

    # ========================================================

    # PROCESS EACH EXCEL ROW

    # ========================================================

    for row_number in range(2, ws.max_row + 1):

        errors = []

        def get_excel_value(column_name):

            return ws.cell(

                row=row_number,

                column=column_map[column_name]

            ).value

        tax_invoice_no = get_excel_value("TaxInvoiceNo")

        if tax_invoice_no is None or clean(tax_invoice_no) == "":

            print(f"Row {row_number}: SKIPPED - Empty TaxInvoiceNo")

            continue

        invoice_key = normalize_invoice(

            tax_invoice_no

        )

        txt_record = None

        emessage_cell = ws.cell(

            row=row_number,

            column=column_map["Emessage"]

        )

        # TaxID is a required OCR field, independent of TXT matching.
        taxid_col = next((col for header, col in column_map.items()
                          if header.casefold() == "taxid"), None)
        missing_taxid = (taxid_col is not None and
                         not clean(ws.cell(row=row_number, column=taxid_col).value))
        if missing_taxid:
            errors.append("TaxID is empty")

        # ----------------------------------------------------

        # 1. Find TaxInvoiceNo in Summary TXT

        # ----------------------------------------------------

        tax_invoice_no = get_excel_value(

            "TaxInvoiceNo"

        )

        invoice_key = normalize_invoice(

            tax_invoice_no

        )

        txt_record = None

        if invoice_key:

            matched_records = summary_records.get(

                invoice_key, []

            )

            if matched_records:

                txt_record = matched_records[0]

        # ----------------------------------------------------

        # 2. If TaxInvoiceNo is NOT FOUND:

        #    Skip ALL remaining validation.

        # ----------------------------------------------------

        if txt_record is None:

            taxremark_col = next((col for header, col in column_map.items()
                                  if header.casefold() == "taxremark"), None)
            missing_taxremark = (taxremark_col is not None and
                                 not clean(ws.cell(row=row_number, column=taxremark_col).value))
            append_emessage(

                emessage_cell,

                ["TaxInvoiceNo is NOT FOUND in Summary Report"] + errors +
                (["TaxRemark is empty"] if missing_taxremark else [])

            )

            result_count["FAIL"] += 1

            result_count["TXT_NOT_FOUND"] += 1

            print(

                f"Row {row_number}: FAIL - ไม่พบ"

            )

            continue

        # ----------------------------------------------------

        # 3. Find VendorTaxId in Master Supplier

        # ----------------------------------------------------

        vendor_tax_id = get_excel_value(

            "VendorTaxId"

        )

        tax_id_key = normalize_tax_id(

            vendor_tax_id

        )

        SupplierName_Master_EN = ""

        SupplierName_Master_TH = ""

        master_found = False

        if tax_id_key:

            master_matches = master_records.get(

                tax_id_key, []

            )

            if master_matches:

                master_found = True

                SupplierName_Master_EN = (

                    master_matches[0][

                        "SupplierName_Master_EN"

                    ]

                )

                SupplierName_Master_TH = (

                    master_matches[0][

                        "SupplierName_Master_TH"

                    ]

                )

        if not master_found:

            errors.append(

                "VendorTaxId NOT FOUND in Master Supplier"

            )

            result_count["MASTER_NOT_FOUND"] += 1

        # ----------------------------------------------------

        # 4. Validate InvoiceDate

        #    Excel vs TXT Field 3

        # ----------------------------------------------------

        excel_date = normalize_date(

            get_excel_value("InvoiceDate")

        )

        txt_date = normalize_date(

            txt_record["InvoiceDate"]

        )

        if excel_date != txt_date:

            errors.append(

                "InvoiceDate is NOT MATCHING"

            )

        # ----------------------------------------------------
        # 5. SupplierName: OCR vs Master and OCR vs TXT
        #    Do not compare TXT SupplierName directly with Master.
        # ----------------------------------------------------
        excel_supplier = get_excel_value("SupplierName") if "SupplierName" in column_map else ""
        txt_supplier = txt_record["SupplierName"]
        if master_found:
            # Compare OCR to each Master row; never compare TXT directly to Master.
            master_names = []
            for master_row in master_matches:
                en_name = master_row["SupplierName_Master_EN"]
                th_name = master_row["SupplierName_Master_TH"]
                # Both languages are evaluated together for the same master row.
                master_names.append(f"{th_name} {en_name}")
            errors.extend(supplier_language_errors(excel_supplier, master_names, "Master"))
        errors.extend(supplier_language_errors(excel_supplier, [txt_supplier], "Text file"))

        # Supplier address is always Excel column F, matched by VendorTaxId.
        # Master column 8 is "Street"; compare the first whitespace token.
        supplier_address = ws.cell(row=row_number, column=6).value
        if master_found and not any(
            first_address_word(supplier_address)
            and first_address_word(supplier_address) == first_address_word(m["StreetNo"])
            for m in master_matches
        ):
            errors.append("Supplier Address is NOT MATCHING")

        # Company Name / Address: Excel Name + address in column O,
        # master keyed by TaxID (master column 4), independently of supplier.
        if taxid_col is not None:
            name_tax_id = ws.cell(row=row_number, column=taxid_col).value
            name_master_matches = master_records.get(normalize_tax_id(name_tax_id), [])
            if name_master_matches:
                if "Name" in column_map and not any(
                    company_name_matches(get_excel_value("Name"),
                                         m["SupplierName_Master_EN"],
                                         m["SupplierName_Master_TH"])
                    for m in name_master_matches
                ):
                    errors.append("Name is NOT MATCHING with Master file")
                name_address = ws.cell(row=row_number, column=15).value
                if not any(
                    first_address_word(name_address)
                    and first_address_word(name_address) == first_address_word(m["StreetNo"])
                    for m in name_master_matches
                ):
                    errors.append("Name Address is NOT MATCHING")

        # ----------------------------------------------------

        # 6. Validate VendorBranch

        #    Excel vs parentheses in TXT Field 20

        # ----------------------------------------------------

        excel_branch = normalize_branch(

            get_excel_value("VendorBranch")

        )

        txt_branch = normalize_branch(

            txt_record["VendorBranch"]

        )

        if excel_branch != txt_branch:

            errors.append(

                "VendorBranch is NOT MATCHING"

            )

        # ----------------------------------------------------

        # 7. Validate TotalAmount

        #    Excel vs TXT Field 16

        # ----------------------------------------------------

        excel_total = normalize_amount(

            get_excel_value("TotalAmount")

        )

        txt_total = normalize_amount(

            txt_record["TotalAmount"]

        )

        if excel_total != txt_total:

            errors.append(

                "TotalAmount is NOT MATCHING"

            )

        # ----------------------------------------------------

        # 8. Validate VATAmount

        #    Excel vs TXT Field 17

        # ----------------------------------------------------

        excel_vat = normalize_amount(

            get_excel_value("VATAmount")

        )

        txt_vat = normalize_amount(

            txt_record["VATAmount"]

        )

        if excel_vat != txt_vat:

            errors.append(

                "VATAmount NOT MATCHING"

            )

        # ----------------------------------------------------

        # 9. Validate AmountIncVat

        #    Excel vs TXT Field 14

        # ----------------------------------------------------

        excel_amount_inc_vat = normalize_amount(

            get_excel_value("AmountIncVat")

        )

        txt_amount_inc_vat = normalize_amount(

            txt_record["AmountIncVat"]

        )

        if excel_amount_inc_vat != txt_amount_inc_vat:

            errors.append(

                "AmountIncVat is NOT MATCHING"

            )

        # ----------------------------------------------------

        # Validate TaxRemark from OCR Excel (blank / whitespace only).
        # Keep this independent of the TXT invoice matching rules.
        taxremark_col = next(
            (col for header, col in column_map.items()
             if header.casefold() == "taxremark"),
            None,
        )
        if taxremark_col is not None:
            taxremark_value = ws.cell(row=row_number, column=taxremark_col).value
            if not clean(taxremark_value):
                errors.append("TaxRemark is empty")

        # 10. Append all errors to Emessage

        # ----------------------------------------------------

        append_emessage(

            emessage_cell,

            errors

        )

        if errors:

            result_count["FAIL"] += 1

            print(

                f"Row {row_number}: FAIL - "

                + ", ".join(errors)

            )

        else:

            result_count["PASS"] += 1

            print(

                f"Row {row_number}: PASS"

            )

    # ========================================================

    # SAVE OUTPUT

    # ========================================================

    output_dir = os.path.dirname(output_path)

    if output_dir:

        os.makedirs(

            output_dir,

            exist_ok=True

        )

    wb.save(output_path)

    wb.close()

    # ========================================================

    # PRINT SUMMARY

    # ========================================================

    print("\n========== VALIDATION RESULT ==========")

    print(f"PASS             : {result_count['PASS']}")

    print(f"FAIL             : {result_count['FAIL']}")

    print(f"TXT NOT FOUND    : {result_count['TXT_NOT_FOUND']}")

    print(f"MASTER NOT FOUND : {result_count['MASTER_NOT_FOUND']}")

    print(f"\nOutput: {output_path}")

# ============================================================

# MAIN

# ============================================================

if __name__ == "__main__":

    import sys

    # รับ path ทั้ง 4 ค่าเป็น parameter จาก 03_430000848; standalone ใช้ default ด้านบน

    if len(sys.argv) == 5:

        excel_path = sys.argv[1]

        text_path = sys.argv[2]

        master_path = sys.argv[3]

        output_excel_path = sys.argv[4]

    elif len(sys.argv) == 1:

        excel_path = EXCEL_PATH

        text_path = TEXT_PATH

        master_path = MASTER_PATH

        output_excel_path = OUTPUT_EXCEL_PATH

    else:

        raise SystemExit(

            "Usage: python Matching_TextFile.py <excel_path> <text_path> <master_path> <output_excel_path>"

        )

    validate_excel(

        excel_path,

        text_path,

        master_path,

        output_excel_path

    )

    # ทำสี Error บนไฟล์ผลลัพธ์ที่ได้รับจาก command line จริง

    # (ไม่ใช้ OUTPUT_EXCEL_PATH ซึ่งเป็นค่า default สำหรับ standalone)

    highlight_excel_errors(output_excel_path)

    # ย้าย TXT หลังจาก Matching และ Highlight สำเร็จแล้วเท่านั้น

    if not os.path.isfile(output_excel_path):

        raise FileNotFoundError(f"ไม่พบ Excel ผลลัพธ์: {output_excel_path}")

    destination_dir = os.path.dirname(os.path.abspath(output_excel_path))

    os.makedirs(destination_dir, exist_ok=True)

    txt_destination = os.path.join(destination_dir, os.path.basename(text_path))

    if os.path.normcase(os.path.abspath(text_path)) != os.path.normcase(os.path.abspath(txt_destination)):

        if not os.path.isfile(text_path):

            raise FileNotFoundError(f"ไม่พบ TXT ต้นฉบับสำหรับย้าย: {text_path}")

        if os.path.exists(txt_destination):

            raise FileExistsError(f"มี TXT ชื่อเดียวกันที่ปลายทางแล้ว ไม่เขียนทับ: {txt_destination}")

        shutil.move(text_path, txt_destination)

        print(f"📦 Moved TXT: {text_path} -> {txt_destination}")

    else:

        print(f"📄 TXT อยู่ในโฟลเดอร์ผลลัพธ์แล้ว: {text_path}")



    # ส่ง Excel ที่ผ่าน Validation และ Highlight แล้ว พร้อมสรุป Emessage
    send_invoice_email(output_excel_path)
