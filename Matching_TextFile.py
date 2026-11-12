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

    # r"C:\TKC\TKC_readinvoices\INPUT\Process3\00000\OutputExcel\text\Excel\OCR\All_Invoices.xlsx"

    r"C:\TKC\Document\INPUT\Process3\00000\OutputExcel\20261112_220831\All_Invoices.xlsx"

)

TEXT_PATH = (

    r"C:\TKC\Document\INPUT\Process3\00000\OutputExcel\20261112_220831\summary invoice report_20260620.txt"

)

MASTER_PATH = (

    r"C:\TKC\TKC_readinvoices\MasterSuplier\Master Vendor Name (MAT)_TH_ENG.xlsx"

)

OUTPUT_EXCEL_PATH = (

    r"C:\TKC\Document\INPUT\Process3\00000\OutputExcel\20261112_220831\ocr_output_matched_ุ6.xlsx"

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


def first_address_word(value):
    # Remove leading OCR quote artifacts without changing the house number.
    text = str(value or "").lstrip(" \t'\"‘’“”")
    # Join commas between house numbers only: a trailing comma must not join
    # the house number to the following street word (e.g. "370, MOO").
    text = re.sub(r"(?<=[0-9])[ \t]*,[ \t]*(?=[0-9])", ",", text)
    parts = text.split(" ", 1)
    return parts[0].replace("\t", "").casefold()


def supplier_matching_errors(ocr_name, master_rows, text_name):
    """Follow OCR -> matching Master row -> Text, with concise messages."""
    matched_names = []
    for row in master_rows:
        name = f'{row["SupplierName_Master_TH"]} {row["SupplierName_Master_EN"]}'
        if supplier_names_match(ocr_name, name):
            matched_names.append(name)
    if not matched_names:
        return ["SupplierName is NOT MATCHING with Master Report"]
    if not clean(text_name) or all(
        not supplier_names_match(name, text_name)
        for name in matched_names
    ):
        return ["SupplierName is MATCHING with Master Report but NOT MATCHING with Text Report"]
    return []


def normalize_address_comparison(value, ignore_commas=False):
    """Keep address comparison cleanup in one place for both OCR and Master."""
    return re.sub(r"[ \t,]" if ignore_commas else r"[ \t]", "", str(value or ""))


def address_matches_master(address, master_row, ignore_commas=False):
    """Compare available street/postcode pairs within the same Master row."""
    street = normalize_address_comparison(first_address_word(address), ignore_commas)
    master_street = normalize_address_comparison(first_address_word(master_row["StreetNo"]), ignore_commas)
    if street and master_street and street != master_street:
        return False
    compact_address = normalize_address_comparison(address, ignore_commas)
    postcode = compact_address[-5:]
    master_postcode = normalize_address_comparison(master_row.get("PostCode"), ignore_commas)
    if re.fullmatch(r"[0-9]{5}", postcode) and master_postcode:
        return postcode == master_postcode
    return True


def normalize_master_branch(value):
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    branch = normalize_branch(value)
    return branch.zfill(5) if branch.isascii() and branch.isdigit() else branch


def address_matching_errors(address, master_rows, branch, supplier=False):
    mismatch = "Supplier Address is NOT MATCHING" if supplier else "Name Address is NOT MATCHING"
    branch_rows = [r for r in master_rows if clean(r.get("BranchCode"))]
    if branch_rows:
        if not clean(branch):
            # Empty-field validation owns this message; do not duplicate it.
            return []
        master_rows = [r for r in branch_rows
                       if normalize_master_branch(r["BranchCode"]) == normalize_master_branch(branch)]
    if not any(address_matches_master(address, row, ignore_commas=not supplier) for row in master_rows):
        return [mismatch]
    return []


def validate_empty_branches(record, master_records):
    """Report empty branch fields once, adding skip context only when needed."""
    messages = []
    for field, tax_field, label in (
        ("Branch", "TaxID", "Address"),
        ("VendorBranch", "VendorTaxId", "Supplier Address"),
    ):
        if field not in record or clean(record[field]):
            continue
        rows = master_records.get(normalize_tax_id(record.get(tax_field)), [])
        message = f"{field} is empty"
        if any(clean(row.get("BranchCode")) for row in rows):
            message += f"; {label} matching skipped"
        messages.append(message)
    return messages


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

        Column 9 = PostCode

    """

    wb = load_workbook(

        master_path,

        read_only=True,

        data_only=True

    )

    ws = wb.active

    master_data = {}
    current_tax = None
    current_record = None
    for row in ws.iter_rows(min_row=2, values_only=True):
        values = list(row) + [None] * max(0, 9 - len(row))
        tax_id, name_en, name_th, branch, street, postcode = values[3:9]
        tax_key = normalize_tax_id(tax_id)
        # Repeated headers and separator rows end a continuation block.
        if clean(tax_id) and not tax_key or not any(clean(v) for v in values):
            current_tax = current_record = None
            continue
        if isinstance(postcode, float) and postcode.is_integer():
            postcode = int(postcode)
        if tax_key:
            current_tax = tax_key
        elif not current_tax or current_record is None:
            continue
        elif not clean(branch):
            # Address continuation lines belong to the current branch only.
            if clean(street):
                current_record["StreetNo"] = " ".join(
                    part for part in (current_record["StreetNo"], clean(street)) if part
                )
            if clean(postcode):
                current_record["PostCode"] = clean(postcode)
            continue
        if not tax_key:
            name_en = current_record["SupplierName_Master_EN"]
            name_th = current_record["SupplierName_Master_TH"]
        current_record = {
            "SupplierName_Master_EN": clean(name_en),
            "SupplierName_Master_TH": clean(name_th),
            "BranchCode": clean(branch),
            "StreetNo": clean(street),
            "PostCode": clean(postcode),
        }
        master_data.setdefault(current_tax, []).append(current_record)

    wb.close()

    return master_data

# ============================================================

# E-MESSAGE FUNCTION

# ============================================================

# Messages generated by this matching stage. Remove stale versions on reruns,
# but retain every original OCR / earlier validation message.
MATCHING_MESSAGES = {
    "TaxInvoiceNo is NOT FOUND in Summary Report",
    "VendorTaxId NOT FOUND in Master Report",
    "InvoiceDate is NOT MATCHING",
    "SupplierName is NOT MATCHING",
    "SupplierName is NOT MATCHING with Master Report",
    "SupplierName is NOT MATCHING with Text Report",
    "SupplierName ไม่ตรงกับ Master",
    "SupplierName ตรงกับ Master ไม่ตรงกับ Text",
    "SupplierName is MATCHING with Master Report but NOT MATCHING with Text Report",
    "Address is NOT MATCHING",
    "Supplier Address is NOT MATCHING",
    "Name Address is NOT MATCHING",
    "Name is NOT MATCHING with Master Report",
    "VendorBranch is NOT MATCHING",
    "TotalAmount is NOT MATCHING",
    "VATAmount is NOT MATCHING",
    "AmountIncVat is NOT MATCHING",
    "TaxRemark is empty",
    "TaxID is empty",
    "Address cannot be matched: Branch is empty",
    "Supplier Address cannot be matched: VendorBranch is empty",
    "Branch is empty; Address matching skipped",
    "VendorBranch is empty; Supplier Address matching skipped",
    "Branch is empty",
    "VendorBranch is empty",
}


def append_emessage(cell, new_messages):
    """Keep pre-existing messages and append unique matching errors with ' | '."""
    existing = clean(cell.value)
    # Discard legacy detailed supplier messages when recomputing this stage.
    existing = " | ".join(
        part for part in existing.split(" | ")
        if not part.strip().casefold().startswith("suppliername ocr vs ")
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

    - Company address mismatch -> ทำสี Address, Branch

    - PO/Company branch mismatch -> ทำสี PurchaseOrderNo, Branch

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

        col_number = column_map.get(field_name)
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
        "Company address does not match": ["Address", "Branch"],
        "PO does not match Company branch": ["PurchaseOrderNo", "Branch"],
        "Duplicate Amount Found": ["TotalAmount", "VATAmount", "AmountIncVat"],
        "Amount format invalid": ["TotalAmount", "VATAmount", "AmountIncVat"],
        "VATAmount format invalid": ["TotalAmount", "VATAmount", "AmountIncVat"],
        "กรุณาเช็ค Amount ทั้ง 3 ช่อง": ["TotalAmount", "VATAmount", "AmountIncVat"],

        # Matching Summary Report
        "TaxInvoiceNo is NOT FOUND in Summary Report": ["TaxInvoiceNo"],
        "VendorTaxId NOT FOUND in Master Report": ["VendorTaxId"],
        "InvoiceDate is NOT MATCHING": ["InvoiceDate"],
        "SupplierName is NOT MATCHING": ["SupplierName"],
        "SupplierName is NOT MATCHING with Master Report": ["SupplierName"],
        "SupplierName is NOT MATCHING with Text Report": ["SupplierName"],
        "SupplierName is MATCHING with Master Report but NOT MATCHING with Text Report": ["SupplierName"],
        "Address is NOT MATCHING": ["Address"],
        "Supplier Address is NOT MATCHING": ["VendorAddress"],
        "Name Address is NOT MATCHING": ["Address"],
        "Branch is empty; Address matching skipped": ["Address", "Branch"],
        "VendorBranch is empty; Supplier Address matching skipped": ["VendorAddress", "VendorBranch"],
        "Branch is empty": ["Branch"],
        "VendorBranch is empty": ["VendorBranch"],
        "Name is NOT MATCHING with Master Report": ["Name"],
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
            matched_error = error_keyword.casefold() in message_parts
            if matched_error:

                for field in fields:

                    mark_cell(row_number, field)

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

        "VendorAddress",

        "Address",

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

        # Own branch-empty messages here, before any TXT matching can skip a row.
        errors.extend(validate_empty_branches(
            {header: ws.cell(row=row_number, column=col).value
             for header, col in column_map.items()}, master_records,
        ))

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

                "VendorTaxId NOT FOUND in Master Report"

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
        # 5. SupplierName: OCR -> Master -> TXT.
        # ----------------------------------------------------
        excel_supplier = get_excel_value("SupplierName") if "SupplierName" in column_map else ""
        txt_supplier = txt_record["SupplierName"]
        errors.extend(supplier_matching_errors(
            excel_supplier, master_matches if master_found else [], txt_supplier
        ))

        # Read supplier address by header, matched by VendorTaxId.
        # Compare the first Street token and available five-digit postcode.
        if master_found:
            errors.extend(address_matching_errors(
                get_excel_value("VendorAddress"), master_matches,
                get_excel_value("VendorBranch"), supplier=True,
            ))

        # Company Name / Address: Excel Name + Address headers,
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
                    errors.append("Name is NOT MATCHING with Master Report")
                errors.extend(address_matching_errors(
                    get_excel_value("Address"), name_master_matches,
                    get_excel_value("Branch") if "Branch" in column_map else None,
                ))

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
