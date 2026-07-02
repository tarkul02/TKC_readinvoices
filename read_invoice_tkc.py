import os
import sys
import io
import re
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from pdf2image import convert_from_path
from PIL import Image
from azure.core.credentials import AzureKeyCredential
from azure.ai.documentintelligence import DocumentIntelligenceClient
from datetime import datetime, timedelta
from dateutil import parser
def parse_date_safe(date_str):
    if not date_str:
        return None

    if isinstance(date_str, datetime):
        return date_str

    for fmt in (
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%Y-%m-%d",
        "%d %b %Y",
        "%d %B %Y"
    ):
        try:
            return datetime.strptime(str(date_str).strip(), fmt)
        except:
            continue

    return None

# ====================================================
# 🔐 ตั้งค่า
# ====================================================
# 1. (Free) pwjdDocIntSEA
# Endoint: https://pwjddocintsea.cognitiveservices.azure.com/
# Key1: AatPPolGeyaXlUHxWfoZwp6qdzYaj62ofvEZJgqtJl43K9Cui6dIJQQJ99CBACqBBLyXJ3w3AAALACOGBvQb
# Key2: D6knMSHETmZgJ13dGIpIhnRpRu9QL8x3TLjT6oYpWkow3fWKP4iYJQQJ99CBACqBBLyXJ3w3AAALACOG5njz

# ถ้า 1. เต็ม
# 2. (Paid) pwjdDocIntSEAPaid
# Endpoint: https://pwjddocintseapaid.cognitiveservices.azure.com/
# Key1: DJInxVJPCWIjReFOSKXpiDPx0Y4guPKPQ6rgi6myTkYppDOyY8c6JQQJ99CFACqBBLyXJ3w3AAALACOGmDHF
# Key2: 7TrXNqo4uNazyc5am2QJfL9egP1INbTrQc91iRUZRYsMo4oVomA6JQQJ99CFACqBBLyXJ3w3AAALACOGoSE

endpoint = "https://pwjddocintseapaid.cognitiveservices.azure.com/"
key = "DJInxVJPCWIjReFOSKXpiDPx0Y4guPKPQ6rgi6myTkYppDOyY8c6JQQJ99CFACqBBLyXJ3w3AAALACOGmDHF"

# today_str = datetime.now().strftime("%Y%m%d_%H%M")
today_str = sys.argv[1]   # INPUT
branch_email = sys.argv[2]   # INPUT

# #test
# today_str = 20262406  # INPUT
# branch_email = 1100   # INPUT

mainpath = "C:/testTKC/OpenAI_Invoice_Processing_AP"

print(today_str)

input_folder = rf"{mainpath}\INPUT\{branch_email}"
temp_folder = rf"{mainpath}\TempSplit"
output_excel = rf"{mainpath}\INPUT\{branch_email}\OutputExcel\{today_str}\Excel\OCR\All_Invoices.xlsx"
output_excel_custom = rf"{mainpath}\INPUT\{branch_email}\OutputExcel\{today_str}\Excel\Custom"
output_excel_done = rf"{mainpath}\INPUT\{branch_email}\OutputExcel\{today_str}\Excel\Done"
dest_folder = rf"{mainpath}\INPUT\{branch_email}\OutputExcel\{today_str}"

os.makedirs(temp_folder, exist_ok=True)
os.makedirs(output_excel_custom, exist_ok=True)
os.makedirs(output_excel_done, exist_ok=True)
os.makedirs(os.path.dirname(output_excel), exist_ok=True)

# ====================================================
# 🚀 สร้าง Client
# ====================================================
client = DocumentIntelligenceClient(endpoint=endpoint, credential=AzureKeyCredential(key))

#load file PATTERN_FILE 
PATTERN_FILE = "patternsInvoice.json"

with open(PATTERN_FILE, "r", encoding="utf-8") as f:
    pattern_config = json.load(f)

INVOICE_PATTERNS = [p["regex"] for p in pattern_config.get("invoice_patterns", [])]

# ====================================================
# ฟังก์ชัน compress PDF หน้าเดียว
# ====================================================
def compress_pdf_page(input_pdf_path, output_pdf_path, max_width=2480, max_height=3508):
    pages = convert_from_path(input_pdf_path)
    writer = PdfWriter()
    for page in pages:
        page.thumbnail((max_width, max_height), Image.Resampling.LANCZOS)
        temp_bytes = io.BytesIO()
        page.save(temp_bytes, format="PDF")
        temp_bytes.seek(0)
        temp_reader = PdfReader(temp_bytes)
        writer.add_page(temp_reader.pages[0])
    with open(output_pdf_path, "wb") as f:
        writer.write(f)

def extract_supplier_name_from_pages(invoices):
    lines = []

    for page in invoices.pages:
        for line in page.lines:
            text = line.content.strip() if line.content else ""
            if text:
                lines.append(text)

    skip_keywords = [
        "TAX INVOICE",
        "INVOICE",
        "ORIGINAL",
        "PAGE",
        "BILL TO",
        "SHIP TO",
        "SELLER TAX ID",
        "COMPANY REGISTRATION",
        "TAX ID",
    ]

    for line in lines[:15]:
        upper = line.upper()

        if any(k in upper for k in skip_keywords):
            continue

        if re.search(r"(CO\.?\s*,?\s*LTD\.?|COMPANY\s+LIMITED|PUBLIC\s+COMPANY\s+LIMITED|บริษัท|จำกัด)", upper, re.IGNORECASE):
            return line.strip()

    return ""

from datetime import datetime
import re

def normalize_invoice_date(date_value):
    """
    รับวันที่หลายรูปแบบ แล้วคืนค่าเป็น DD/MM/YYYY
    รองรับ:
      20/06/69
      20/06/2569
      20/06/2026
      2026-06-20
      20-06-69
      20-06-2026
      20 Jun 2026
      20 June 2026
    """

    if not date_value:
        return ""

    if isinstance(date_value, datetime):
        return date_value.strftime("%d/%m/%Y")

    text = str(date_value).strip()

    # ------------------------
    # 20/06/69
    # ------------------------
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{2})$", text)
    if m:
        d, mth, y = map(int, m.groups())
        return f"{d:02d}/{mth:02d}/{2500+y-543:04d}"

    # ------------------------
    # 20-06-69
    # ------------------------
    m = re.match(r"^(\d{1,2})-(\d{1,2})-(\d{2})$", text)
    if m:
        d, mth, y = map(int, m.groups())
        return f"{d:02d}/{mth:02d}/{2500+y-543:04d}"

    formats = [

        "%d/%m/%Y",
        "%Y-%m-%d",
        "%d-%m-%Y",

        "%d %b %Y",
        "%d %B %Y",

        "%d %b %y",
        "%d %B %y",
    ]

    for fmt in formats:

        try:

            dt = datetime.strptime(text, fmt)

            year = dt.year

            # ถ้า OCR ได้ พ.ศ.
            if year > 2400:
                year -= 543

            dt = dt.replace(year=year)

            return dt.strftime("%d/%m/%Y")

        except:
            pass

    return text

def extract_vendor_branch(invoices, layout_result=None):

    lines = []

    for page in invoices.pages:
        for line in page.lines:
            if line.content:
                lines.append(line.content.strip())

    full_text = "\n".join(lines)

    # =====================================================
    # 1. ถ้ามี Checkbox ให้ใช้ก่อน
    # =====================================================
    if layout_result:

        for page in layout_result.pages:

            selected_marks = [
                m for m in (page.selection_marks or [])
                if m.state.name == "SELECTED"
            ]

            for mark in selected_marks:

                mark_x = mark.polygon[0]
                mark_y = mark.polygon[1]

                same_row_texts = []

                for line in page.lines:

                    text = line.content.strip()
                    line_x = line.polygon[0]
                    line_y = line.polygon[1]

                    # เอาเฉพาะข้อความที่อยู่แถวเดียวกับ checkbox
                    if abs(mark_y - line_y) <= 0.08:
                        same_row_texts.append(text)

                same_row_text = " ".join(same_row_texts)

                if re.search(r"สำนักงานใหญ่|Head\s*Office|HeadOffice", same_row_text, re.IGNORECASE):
                    return "00000"

                m = re.search(
                    r"สาขา(?:ที่|เลขที่)?\s*[:：]?\s*(\d{1,10})",
                    same_row_text,
                    re.IGNORECASE
                )

                if m:
                    return m.group(1).zfill(5)

    # =====================================================
    # 2. ใบกำกับภาษีออกโดย ...
    # =====================================================

    m = re.search(
        r"ออกโดย\s*[:：]?\s*(สำนักงานใหญ่|Head\s*Office|HeadOffice)",
        full_text,
        re.IGNORECASE,
    )

    if m:
        return "00000"

    m = re.search(
        r"ออกโดย\s*[:：]?\s*สาขา(?:ที่|เลขที่)?\s*(\d{1,10})",
        full_text,
        re.IGNORECASE,
    )

    if m:
        return m.group(1).zfill(5)

    # =====================================================
    # 3. อ่านเฉพาะ Vendor ก่อนถึง Customer
    # =====================================================

    vendor_lines = []

    stop_keywords = [
        r"^\s*BILL\s*TO",
        r"^\s*SHIP\s*TO",
        r"^\s*TO\s*:",
        r"^\s*Customer",
        r"^\s*Customer Name",
        r"^\s*Customer Code",
        r"^\s*รหัสลูกค้า",
        r"^\s*ชื่อลูกค้า",
        r"^\s*Delivery To",
    ]

    for line in lines:

        if any(re.search(p, line, re.IGNORECASE) for p in stop_keywords):
            break

        vendor_lines.append(line)

    vendor_text = "\n".join(vendor_lines)

    # =====================================================
    # 4. หาเลขสาขา Vendor
    # =====================================================

    patterns = [

        r"Branch\s*No\.?\s*(\d{1,10})",

        r"Branch\s*(?:No\.?|Number|Code|ID)?\s*[:：]?\s*(\d{1,10})",

        r"สาขา(?:ที่|เลขที่)?\s*[:：]?\s*(\d{1,10})",

    ]

    for pattern in patterns:

        m = re.search(pattern, vendor_text, re.IGNORECASE)

        if m:
            return m.group(1).zfill(5)

    # =====================================================
    # 5. Vendor เป็นสำนักงานใหญ่
    # =====================================================

    if re.search(
        r"(สำนักงานใหญ่|Head\s*Office|HeadOffice)",
        vendor_text,
        re.IGNORECASE,
    ):
        return "00000"

    return ""

def extract_checked_branch(layout_result):

    for page in layout_result.pages:

        selected_marks = [
            m for m in (page.selection_marks or [])
            if m.state.name == "SELECTED"
        ]

        if not selected_marks:
            continue

        for mark in selected_marks:

            mark_x = mark.polygon[0]
            mark_y = mark.polygon[1]

            best_branch = ""
            best_distance = 999999

            for line in page.lines:

                text = line.content.strip()

                match = re.search(
                    r"สาขา(?:ที่|เลขที่)?\s*[:：]?\s*(\d{1,10})",
                    text
                )

                if not match:
                    continue

                line_x = line.polygon[0]
                line_y = line.polygon[1]

                distance = (
                    abs(line_x - mark_x)
                    + abs(line_y - mark_y)
                )

                if distance < best_distance:
                    best_distance = distance
                    best_branch = match.group(1)

            if best_branch:
                return best_branch.zfill(5)

    return ""

#เช็ค vat เพิ่มในกรณีที่หาไม่เจอ
def extract_vat_from_pages(invoices):
    for page in invoices.pages:
        for line in page.lines:
            text = line.content.strip() if line.content else ""
            clean_text = text.replace(",", "")

            # ต้องมี VAT 7%
            if not re.search(r"VAT\s*7\s*%", clean_text, re.IGNORECASE):
                continue

            # เอาเฉพาะข้อความหลังคำว่า VAT 7%
            after_vat = re.split(
                r"VAT\s*7\s*%",
                clean_text,
                flags=re.IGNORECASE
            )[-1]

            nums = re.findall(r"\d+\.\d{2}", after_vat)

            if nums:
                # กรณีเจอ 235.20 ซ้ำ 2 ครั้ง ให้เอาตัวแรกหลัง VAT 7%
                return float(nums[0])

    return 0.0

def extract_tax_id_from_pages(invoices):

    full_text = ""

    for page in invoices.pages:
        for line in page.lines:
            if line.content:
                full_text += "\n" + line.content.strip()

    patterns = [

        r"เลขประจำตัวผู้เสียภาษี.*?([0-9OIl\s\-]{13,25})",

        r"TAX\s*ID.*?([0-9OIl\s\-]{13,25})",

        r"TIN.*?([0-9OIl\s\-]{13,25})",
    ]

    for pattern in patterns:

        m = re.search(pattern, full_text, re.IGNORECASE)

        if not m:
            continue

        tax = m.group(1)

        # OCR มักอ่านผิด
        tax = (
            tax.replace("O", "0")
               .replace("I", "1")
               .replace("l", "1")
        )

        tax = re.sub(r"\D", "", tax)

        if len(tax) >= 13:
            return tax[:13]

    # fallback
    ids = re.findall(r"0\d{12}", re.sub(r"\D", "", full_text))

    if ids:
        return ids[0]

    return ""

def extract_tax_invoice_no_from_layout(layout_result):

    lines = []

    for page in layout_result.pages:
        for line in page.lines:
            if line.content:
                lines.append(line.content.strip())

    full_text = " ".join(lines)
    full_text = re.sub(r"\s+", " ", full_text)

    patterns = [
        # Account No. CH008206 No. : C260100384
        r"Account\s*No\.?\s*[:：]?\s*[A-Za-z0-9\-\/]+\s*No\.?\s*[:：]?\s*([A-Za-z]\d{6,20})",

        # No. : C260100384
        r"\bNo\.?\s*[:：]?\s*([A-Za-z]\d{6,20})",
    ]

    for pattern in patterns:
        match = re.search(pattern, full_text, re.IGNORECASE)
        if match:
            return match.group(1).strip()

    return ""

def extract_po_from_text_and_tables(invoices):
        po_no_list = []

        # หาใน lines ทั้งหน้า
        for page in invoices.pages:
            for line in page.lines:
                text = line.content.strip() if line.content else ""
                matches = re.findall(r"(?:PO)?(410\d{7})", text)
                po_no_list.extend(matches)

        # หาใน tables
        if hasattr(invoices, "tables"):
            for table in invoices.tables:
                for cell in table.cells:
                    text = cell.content.strip() if cell.content else ""
                    matches = re.findall(r"(?:PO)?(410\d{7})", text)
                    po_no_list.extend(matches)

        # ลบค่าซ้ำ แต่คงลำดับเดิม
        po_unique = list(dict.fromkeys(po_no_list))

        return ",".join(po_unique)




def merge_invoice_row(existing, new):
    # เติมข้อมูลที่ว่าง
    for field in [
        "InvoiceDate",
        "PostingDate",
        "TaxInvoiceNo",
        "SupplierName",
        "Assignment",
        "VendorTaxId",
        "VendorBranch",
        "PurchaseOrderNo1",
        "PurchaseOrderNo",
    ]:
        if (not existing.get(field)) and new.get(field):
            existing[field] = new.get(field)

    # รวม PO ไม่ให้ซ้ำ
    for field in ["PurchaseOrderNo1", "PurchaseOrderNo"]:
        vals = []
        for v in [existing.get(field,""), new.get(field,"")]:
            if v:
                vals.extend([x.strip() for x in str(v).split(",") if x.strip()])
        if vals:
            existing[field] = ",".join(dict.fromkeys(vals))

    # ยอดเงิน เอาค่าที่ไม่เป็น 0
    for field in ["TotalAmount","VATAmount","AmountIncVat"]:
        try:
            if float(existing.get(field,0) or 0)==0 and float(new.get(field,0) or 0)>0:
                existing[field]=new.get(field)
        except:
            pass

    # รวมข้อความ Error
    if new.get("Emessage"):
        if existing.get("Emessage"):
            existing["Emessage"] += " | " + new["Emessage"]
        else:
            existing["Emessage"] = new["Emessage"]

    return existing


def to_float(value):
    try:
        return float(str(value).replace(",", "").strip())
    except:
        return 0.0

# ====================================================
# ฟังก์ชันแปลง invoice fields เป็น JSON ตาม prompt
# ====================================================
def extract_invoice_to_json(invoice, invoices):

    def get_string(field):
        if not field:
            return ""
        if hasattr(field, "value_string") and field.value_string:
            return field.value_string.strip()
        if hasattr(field, "value_date") and field.value_date:
            return field.value_date.strftime("%d/%m/%Y")
        if hasattr(field, "value_currency") and field.value_currency:
            return float(field.value_currency.amount)
        if hasattr(field, "value_number") and field.value_number is not None:
            return float(field.value_number)
        return ""

    row = {}
    supplier_name = get_string(invoice.fields.get("VendorAddressRecipient"))

    #ถ้าไม่เจอ supplier_nameให้เช็คเพิ่ม
    if not supplier_name:
        supplier_name = extract_supplier_name_from_pages(invoices)

    # invoice_date = get_string(invoice.fields.get("InvoiceDate"))
    invoice_date = normalize_invoice_date(
        get_string(invoice.fields.get("InvoiceDate"))
    )
    tax_invoice_no = get_string(invoice.fields.get("InvoiceId")).replace("\n", ",")
    tax_invoice_no_clean = tax_invoice_no.split(",")[-1].strip()
    purchase_order_no = get_string(invoice.fields.get("PurchaseOrder")).replace("\n", ",")
    total_amount = get_string(invoice.fields.get("SubTotal"))
    vat_amount = get_string(invoice.fields.get("TotalTax"))
    amount_inc_vat = get_string(invoice.fields.get("InvoiceTotal"))
    tax_code=""
    # service_start = get_string(invoice.fields.get("ServiceStartDate"))
    service_start = normalize_invoice_date(
        get_string(invoice.fields.get("ServiceStartDate"))
    )
    # due_date = get_string(invoice.fields.get("DueDate"))
    due_date = normalize_invoice_date(
        get_string(invoice.fields.get("DueDate"))
    )
    VendorTaxId = re.sub(
        r"\D",
        "",
        get_string(invoice.fields.get("VendorTaxId"))
    )
    CustomerTaxId = get_string(invoice.fields.get("CustomerTaxId"))

    def parse_date(date_str):
        if not date_str:
            return None
        # รองรับหลายรูปแบบ เช่น 2024-11-01 / 01/11/2024 / 1 Nov 2024
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d %b %Y", "%d %B %Y"):
            try:
                return datetime.strptime(date_str, fmt)
            except:
                pass
        return None

    def ymd_to_dmy(date_str):
        """แปลง YYYY-MM-DD → DD/MM/YYYY"""
        try:
            dt = datetime.strptime(date_str, "%Y-%m-%d")
            return dt.strftime("%d/%m/%Y")
        except:
            return date_str
    
    # แปลงเป็น supplier_name
    if supplier_name:
        supplier_name = supplier_name.replace(";", ",")

    # แปลงเป็น datetime
    invoice_dt = parse_date(invoice_date)
    service_dt = parse_date(service_start)
    due_dt = parse_date(due_date)

    current_year = datetime.now().year

    # ==========================
    # ✔ Logic ตามที่สั่ง
    # ==========================
    final_invoice_date = invoice_date  # ค่า default

    if invoice_dt:
        if invoice_dt.year != current_year:
            # หาอันที่น้อยสุดใน ServiceStartDate / DueDate
            candidates = [d for d in [service_dt, due_dt] if d is not None]

            if candidates:
                smallest_date = min(candidates)
                final_invoice_date = smallest_date.strftime("%Y-%m-%d")
        else:
            # ถ้าเป็นปีปัจจุบัน → แปลง format ให้เป็น dd/mm/yyyy
            final_invoice_date = invoice_dt.strftime("%Y-%m-%d")

    else:
        # ถ้า InvoiceDate parse ไม่ได้ → fallback หาอันที่น้อยสุดเหมือนกัน
        candidates = [d for d in [service_dt, due_dt] if d is not None]
        if candidates:
            smallest_date = min(candidates)
            final_invoice_date = smallest_date.strftime("%Y-%m-%d")

    # แปลง YYYY-MM-DD → DD/MM/YYYY ก่อนใช้งานจริง
    final_invoice_date = ymd_to_dmy(final_invoice_date)


    # Cleanup PO
    po_no = purchase_order_no
    if po_no:
        po_list = [p.strip() for p in po_no.split(",") if p.strip()]
        valid_po_list = [p for p in po_list if p.startswith("PO410") or p.startswith("410")]
        cleaned_list = [p[2:] if p.startswith("PO") else p for p in valid_po_list]
        po_no = ",".join(cleaned_list)

    return {
        "InvoiceDate": final_invoice_date,
        "PostingDate": final_invoice_date,
        "SupplierName": supplier_name,
        "TaxInvoiceNo": tax_invoice_no_clean,
        "TotalAmount": float(total_amount) if total_amount != "" else 0.0,
        "VATAmount": float(vat_amount) if vat_amount != "" else 0.0,
        "AmountIncVat": float(amount_inc_vat) if amount_inc_vat != "" else 0.0,
        "PurchaseOrderNo1": purchase_order_no,
        "PurchaseOrderNo": po_no,
        "VendorTaxId": VendorTaxId,
        "Emessage":"",
    }

# ====================================================
# อ่านไฟล์ PDF ทั้ง Folder
# ====================================================
pdf_list = [
    os.path.join(input_folder, f)
    for f in os.listdir(input_folder)
    if f.lower().endswith(".pdf")
]

if not pdf_list:
    print("❌ ไม่พบไฟล์ PDF ในโฟลเดอร์ InputPDF")
    exit()

all_data = []

for input_pdf in pdf_list:

    print(f"\n==============================")
    print(f"📁 Processing PDF File: {input_pdf}")
    print(f"==============================\n")

    # --------------------------------------------------
    # แยก PDF เป็นหน้าเดียว
    # --------------------------------------------------
    reader = PdfReader(input_pdf)
    pdf_files = []

    for i, page in enumerate(reader.pages):
        single_page_path = os.path.join(temp_folder, f"page_{i+1}.pdf")
        writer = PdfWriter()
        writer.add_page(page)
        with open(single_page_path, "wb") as f:
            writer.write(f)

        if os.path.getsize(single_page_path) > 50*1024*1024:
            print(f"⚠️ Page {i+1} too large → compressing...")
            compressed_path = os.path.join(temp_folder, f"page_{i+1}_compressed.pdf")
            compress_pdf_page(single_page_path, compressed_path)
            pdf_files.append(compressed_path)
        else:
            pdf_files.append(single_page_path)

    # --------------------------------------------------
    # วิเคราะห์แต่ละ PDF
    # --------------------------------------------------
    for pdf_path in pdf_files:
        print(f"\n📄 Processing Page: {pdf_path}")
        with open(pdf_path, "rb") as f:
            poller = client.begin_analyze_document(model_id="prebuilt-invoice", body=f)
        invoices = poller.result()

        with open(pdf_path, "rb") as f:
            poller_layout = client.begin_analyze_document(
                model_id="prebuilt-layout",
                body=f
            )

        layout_result = poller_layout.result()

        date_candidates = []
        today = datetime.today()
        one_year_ago = today - timedelta(days=365)
        one_year_future = today + timedelta(days=365)

        raw_page_text = ""
        raw_text = ""   
        for page in invoices.pages:
            page_text = " ".join([line.content for line in page.lines if line.content]).strip()
            raw_text += " " + page_text
            raw_page_text += " " + page_text.lower()

            # 🔎 Regex จับวันที่หลายรูปแบบ (ยืดหยุ่น)
            date_pattern = re.compile(
                r'\b('
                r'\d{4}-\d{2}-\d{2}'                  # 2026-02-23
                r'|\d{1,2}[\/-]\d{1,2}[\/-]\d{4}'     # 23/02/2026 , 23-02-2026
                r'|\d{1,2}[\s-][A-Za-z]{3,9}[\s-]\d{4}'  # 23 Feb 2026 , 23-Feb-2026
                r')\b'
            )

            matches = date_pattern.findall(raw_text)

            for date_str in matches:
                try:
                    parsed_date = parser.parse(date_str, dayfirst=True)

                    # ✅ กรองเฉพาะวันที่อยู่ในช่วง ±1 ปี
                    if one_year_ago <= parsed_date <= one_year_future:
                        date_candidates.append(parsed_date)

                except:
                    continue

            # 🎯 เลือกวันที่เก่าสุด
            invoiceDate_2 = ""

            if date_candidates:
                oldest_date = min(date_candidates)
                invoiceDate_2 = oldest_date.strftime("%d/%m/%Y")
                print(f"📅 Oldest valid date (±1 year): {invoiceDate_2}")
            else:
                print("⚠️ No valid date found within ±1 year range")

        if "good receipt" in raw_page_text or "goods receipt" in raw_page_text:
            print("⏭️ พบคำว่า 'Good Receipt' → ข้ามหน้านี้ทันที")
            continue

        for idx, invoice in enumerate(invoices.documents):
            print(f"\n-------- Invoice #{idx + 1} --------")
            invoice_data = extract_invoice_to_json(invoice, invoices)

            # -------------------------
            # หา vat อีกครั้ง
            # -------------------------
            if float(invoice_data.get("VATAmount", 0) or 0) == 0:
                vat_fallback = extract_vat_from_pages(invoices)

                if vat_fallback:
                    invoice_data["VATAmount"] = vat_fallback


            exclude_tax_ids = {"0105529030059"}

            if invoice_data.get("VendorTaxId") in exclude_tax_ids:
                invoice_data["VendorTaxId"] = ""

            if not invoice_data.get("VendorTaxId"):
                invoice_data["VendorTaxId"] = extract_tax_id_from_pages(invoices)

            vendor_branch = extract_vendor_branch(invoices, layout_result)

            invoice_data["VendorBranch"] = vendor_branch

            if not invoice_data.get("TaxInvoiceNo"):
                fallback_no = extract_tax_invoice_no_from_layout(layout_result)

                if fallback_no:
                    invoice_data["TaxInvoiceNo"] = fallback_no
                    print(f"✅ Fallback TaxInvoiceNo from layout: {fallback_no}")

            invoice_data["Assignment"] = os.path.basename(input_pdf)

            # =====================================
            # Normalize Amount
            # =====================================
            try:

                total_amount = float(invoice_data.get("TotalAmount", 0) or 0)
                vat_amount = float(invoice_data.get("VATAmount", 0) or 0)
                amount_inc_vat = float(invoice_data.get("AmountIncVat", 0) or 0)

                values = [total_amount, vat_amount, amount_inc_vat]

                # ถ้ามีค่าซ้ำกัน
                if len(set(values)) < 3:

                    current_msg = invoice_data.get("Emessage", "")

                    msg = (
                        f"Duplicate Amount Found "
                        f"(Total={total_amount}, "
                        f"VAT={vat_amount}, "
                        f"AmountIncVat={amount_inc_vat})"
                    )

                    invoice_data["Emessage"] = (
                        current_msg + " | " + msg
                        if current_msg else msg
                    )

                else:

                    values.sort()

                    invoice_data["VATAmount"] = values[0]
                    invoice_data["TotalAmount"] = values[1]
                    invoice_data["AmountIncVat"] = values[2]

            except Exception:

                current_msg = invoice_data.get("Emessage", "")

                invoice_data["Emessage"] = (
                    current_msg + " | Amount format invalid"
                    if current_msg
                    else "Amount format invalid"
                )

            # ===============================
            # 🔎 หา InvoiceId จาก OCR text (fallback)
            # ===============================
            if not invoice_data.get("TaxInvoiceNo"):
                invoice_no = ""

                for page in invoices.pages:
                    if not hasattr(page, "words"):
                        continue

                    for w in page.words:
                        text = w.content.strip() if w.content else ""

                        # A260100383, AR123456789, AB1234567890
                        if re.match(r"[A-Z]{1,3}\d{6,20}$", text):
                            invoice_no = text
                            break

                        # แบบใหม่: IV6808/0009
                        if re.match(r"[A-Z]{2}\d{4}/\d{4}$", text):
                            invoice_no = text
                            break

                        # แบบ: 20-213233
                        if re.match(r"\d{2}-\d{6}$", text):
                            invoice_no = text
                            break

                    if invoice_no:
                        break

                if invoice_no:
                    invoice_data["TaxInvoiceNo"] = invoice_no
                    print(f"✅ Fallback InvoiceId found from OCR: {invoice_no}")
                else:
                    print("⚠️ InvoiceId not found (even from OCR)")

            # -------------------------------
            # หา PO_temp และ merge จาก tables / text pages
            # -------------------------------
            invoice_no_list = []
            po_no_list = []

            # หา Invoice No จาก table ถ้า TaxInvoiceNo ยังว่าง
            if not invoice_data.get("TaxInvoiceNo") and hasattr(invoices, "tables"):
                invoice_no_list = []

                for table in invoices.tables:
                    for cell in table.cells:
                        text = cell.content.strip() if cell.content else ""

                        if re.search(r'IG\d{3,6}[/\-]\d{3,5}', text):
                            invoice_no_list.extend(
                                re.findall(r'IG\d{3,6}[/\-]\d{3,5}', text)
                            )

                if invoice_no_list:
                    invoice_data["TaxInvoiceNo"] = ",".join(list(dict.fromkeys(invoice_no_list)))


            # หา PO 410xxxxxxx จากทั้ง text และ table
            po_from_ocr = extract_po_from_text_and_tables(invoices)

            if po_from_ocr:
                invoice_data["PurchaseOrderNo"] = po_from_ocr

            # Merge PO จากข้อความหน้า PDF
            existing_po = invoice_data.get("PurchaseOrderNo", "")
            PO_temp = ""
            if existing_po and (existing_po.startswith("410") or existing_po.startswith("PO410")):
                PO_temp = existing_po
            else:
                po_pattern = re.compile(r'(?:PO410\d{7}|410\d{7})')
                PO_list = []
                for page in invoices.pages:
                    page_text = " ".join([line.content for line in page.lines if line.content]).strip()
                    matches = po_pattern.findall(page_text)
                    if matches:
                        PO_list.extend(matches)
                if PO_list:
                    PO_unique = list(dict.fromkeys(PO_list))
                    PO_temp = ",".join(PO_unique)
                    if not invoice_data.get("PurchaseOrderNo"):
                        invoice_data["PurchaseOrderNo"] = PO_temp

        

            tax_invoice_no = (invoice_data.get("TaxInvoiceNo") or "").strip()

            # ถ้าขึ้นต้นด้วย 410 → เคลียร์ค่า
            if tax_invoice_no.startswith("410"):
                tax_invoice_no = ""
            
            if tax_invoice_no.startswith("PO410"):
                tax_invoice_no = ""

            invoice_data["TaxInvoiceNo"] = tax_invoice_no

            invoice_date_main = parse_date_safe(invoice_data.get("InvoiceDate"))
            invoice_date_oldest = parse_date_safe(invoiceDate_2)

            final_date = None

            if invoice_date_main and invoice_date_oldest:
                # ✅ เอาวันที่ที่เก่ากว่า
                invoice_date_main = parse_date_safe(invoice_data.get("InvoiceDate"))

                if invoice_date_main:
                    final_date_str = invoice_date_main.strftime("%d/%m/%Y")
                    invoice_data["InvoiceDate"] = final_date_str
                    invoice_data["PostingDate"] = final_date_str

                elif invoiceDate_2:
                    invoice_date_oldest = parse_date_safe(invoiceDate_2)

                    if invoice_date_oldest:
                        final_date_str = invoice_date_oldest.strftime("%d/%m/%Y")
                        invoice_data["InvoiceDate"] = final_date_str
                        invoice_data["PostingDate"] = final_date_str

            elif invoice_date_oldest:
                final_date = invoice_date_oldest

            elif invoice_date_main:
                final_date = invoice_date_main

            # 🎯 แทนค่า InvoiceDate ด้วยวันที่ที่เก่ากว่า
            if final_date:
                final_date_str = final_date.strftime("%d/%m/%Y")

                invoice_data["InvoiceDate"] = final_date_str
                invoice_data["PostingDate"] = final_date_str

            merged = False

            current_tax = (invoice_data.get("TaxInvoiceNo") or "").strip()

            if current_tax:
                for row in all_data:
                    if (row.get("TaxInvoiceNo") or "").strip() == current_tax:
                        merge_invoice_row(row, invoice_data)
                        merged = True
                        break

            if not merged:
                all_data.append(invoice_data)
 
        # ===========================
        # 📌 Move PDF after processed
        # ===========================
        os.makedirs(dest_folder, exist_ok=True)

        pdf_name = os.path.basename(input_pdf)
        dest_path = os.path.join(dest_folder, pdf_name)

        # 👉 ถ้าไฟล์ชื่อเดียวกันมีอยู่แล้ว → ไม่ต้องย้าย
        if os.path.exists(dest_path):
            print(f"⚠️ File already exists, skip move → {dest_path}")
        else:
            try:
                os.rename(input_pdf, dest_path)
                print(f"📁 Moved processed PDF → {dest_path}")
            except Exception as e:
                print(f"⚠️ Move failed: {e}")

# ====================================================
# 💾 บันทึก Excel
# ====================================================
columns = [
    "InvoiceDate",
    "PostingDate",
    "TaxInvoiceNo",
    "SupplierName",
    "Assignment",
    "VendorTaxId",
    "VendorBranch",
    "TotalAmount",
    "VATAmount",
    "AmountIncVat",
    "PurchaseOrderNo",
    "Emessage",
]

required_fields = [
    "InvoiceDate",
    "PostingDate",
    "TaxInvoiceNo",
    "TotalAmount",
    "VATAmount",
    "SupplierName",
    "VendorTaxId",
    "VendorBranch",
    "PurchaseOrderNo",
]

for row in all_data:

    errors = []

    # เช็คค่าว่าง
    for field in required_fields:

        value = row.get(field)

        if value is None or str(value).strip() == "":
            errors.append(f"{field} is empty")

    # เช็ค VAT
    try:

        total_amount = float(row.get("TotalAmount", 0) or 0)
        vat_amount = float(row.get("VATAmount", 0) or 0)

        expected_vat = round(total_amount * 0.07, 2)

        # ยอมคลาดเคลื่อน 1 สตางค์
        if abs(vat_amount - expected_vat) > 0.01:
            errors.append(
                f"VATAmount ไม่ถูกต้อง (Expected {expected_vat:.2f})"
            )

    except Exception:
        errors.append("VATAmount format invalid")

    existing_msg = row.get("Emessage", "")
    new_error = " | ".join(errors)

    if existing_msg and new_error:
        row["Emessage"] = existing_msg + " | " + new_error
    elif existing_msg:
        row["Emessage"] = existing_msg
    else:
        row["Emessage"] = new_error

df = pd.DataFrame(all_data)

# เอาเฉพาะคอลัมน์ที่ต้องการ
df = df.reindex(columns=columns)

df.to_excel(output_excel, index=False)

print(f"\n✅ Done. All invoices saved to {output_excel}")
