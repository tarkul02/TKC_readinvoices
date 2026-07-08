import os
import sys
import io
import re
import json
import time
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from pdf2image import convert_from_path
from PIL import Image
from azure.core.credentials import AzureKeyCredential
from azure.ai.documentintelligence import DocumentIntelligenceClient
from azure.core.exceptions import ServiceRequestError, HttpResponseError
from datetime import datetime, timedelta
from dateutil import parser

# ====================================================
# 🔐 ตั้งค่า
# ====================================================
endpoint = "https://pwjddocintseapaid.cognitiveservices.azure.com/"
key = "DJInxVJPCWIjReFOSKXpiDPx0Y4guPKPQ6rgi6myTkYppDOyY8c6JQQJ99CFACqBBLyXJ3w3AAALACOGmDHF"

today_str = sys.argv[1]      # INPUT
branch_email = sys.argv[2]   # INPUT

mainpath = "C:/TOA/TKC_readinvoices"

input_folder = rf"{mainpath}\INPUT\{branch_email}"
temp_folder = rf"{mainpath}\TempSplit"
output_excel = rf"{mainpath}\INPUT\{branch_email}\OutputExcel\{today_str}\Excel\OCR\All_Invoices.xlsx"
dest_folder = rf"{mainpath}\INPUT\{branch_email}\OutputExcel\{today_str}"

os.makedirs(temp_folder, exist_ok=True)
os.makedirs(os.path.dirname(output_excel), exist_ok=True)
os.makedirs(dest_folder, exist_ok=True)

client = DocumentIntelligenceClient(
    endpoint=endpoint,
    credential=AzureKeyCredential(key)
)

# ====================================================
# 📌 Load patternsInvoice.json
# ====================================================
PATTERN_FILE = "patternsInvoice.json"

DEFAULT_PATTERN_CONFIG = {
    "invoice_patterns": [
        {"name": "Invoice AlphaNumeric", "regex": r"[A-Z]{1,3}\d{6,20}"},
        {"name": "Invoice Slash", "regex": r"[A-Z]{2}\d{4}/\d{4}"},
        {"name": "Invoice Dash", "regex": r"\d{2}-\d{6}"},
        {"name": "IG Invoice", "regex": r"IG\d{3,6}[/-]\d{3,5}"}
    ],
    "po_patterns": [
        {"name": "PO410", "regex": r"(?:PO)?(410\d{7})"},
        {"name": "PO140", "regex": r"(?:PO)?(140\d{7})"}
    ],
    "vendor_branch_patterns": [
        {"name": "Thai Branch", "regex": r"สาขา(?:ที่|เลขที่)?\s*[:：]?\s*(\d{1,10})"},
        {"name": "English Branch", "regex": r"Branch\s*(?:No\.?|Number|Code|ID)?\s*[:：]?\s*(\d{1,10})"}
    ],
    "taxid_patterns": [
        {"name": "Thai TaxID", "regex": r"เลขประจำตัวผู้เสียภาษี.*?([0-9OIl\s\-]{13,30})"},
        {"name": "English TaxID", "regex": r"TAX\s*ID.*?([0-9OIl\s\-]{13,30})"},
        {"name": "TIN", "regex": r"TIN.*?([0-9OIl\s\-]{13,30})"}
    ],
    "exclude_tax_ids": ["0105529030059"],
    "stop_keywords": [
        "BILL TO", "SHIP TO", "TO:", "Customer", "Customer Name",
        "Customer Code", "รหัสลูกค้า", "ชื่อลูกค้า", "Delivery To",
        "ส่งของที่", "ผู้ซื้อ"
    ],
    "supplier_skip_keywords": [
        "TAX INVOICE", "INVOICE", "ORIGINAL", "PAGE", "BILL TO", "SHIP TO",
        "SELLER TAX ID", "COMPANY REGISTRATION", "TAX ID", "TIN", "ชื่อผู้ซื้อ" 
    ],
    "supplier_name_patterns": [
        {"name": "Company Name", "regex": r"(CO\.?\s*,?\s*LTD\.?|COMPANY\s+LIMITED|PUBLIC\s+COMPANY\s+LIMITED|บริษัท|จำกัด)"}
    ]
}

try:
    with open(PATTERN_FILE, "r", encoding="utf-8") as f:
        pattern_config = json.load(f)
except FileNotFoundError:
    print(f"⚠️ ไม่พบ {PATTERN_FILE} → ใช้ Default Pattern ในโปรแกรม")
    pattern_config = DEFAULT_PATTERN_CONFIG

INVOICE_PATTERNS = [p["regex"] for p in pattern_config.get("invoice_patterns", DEFAULT_PATTERN_CONFIG["invoice_patterns"])]
PO_PATTERNS = [p["regex"] for p in pattern_config.get("po_patterns", DEFAULT_PATTERN_CONFIG["po_patterns"])]
VENDOR_BRANCH_PATTERNS = [p["regex"] for p in pattern_config.get("vendor_branch_patterns", DEFAULT_PATTERN_CONFIG["vendor_branch_patterns"])]
TAXID_PATTERNS = [p["regex"] for p in pattern_config.get("taxid_patterns", DEFAULT_PATTERN_CONFIG["taxid_patterns"])]
EXCLUDE_TAX_IDS = set(pattern_config.get("exclude_tax_ids", DEFAULT_PATTERN_CONFIG["exclude_tax_ids"]))
STOP_KEYWORDS = pattern_config.get("stop_keywords", DEFAULT_PATTERN_CONFIG["stop_keywords"])
SUPPLIER_SKIP_KEYWORDS = pattern_config.get("supplier_skip_keywords", DEFAULT_PATTERN_CONFIG["supplier_skip_keywords"])
SUPPLIER_NAME_PATTERNS = [p["regex"] for p in pattern_config.get("supplier_name_patterns", DEFAULT_PATTERN_CONFIG["supplier_name_patterns"])]

# ====================================================
# 📌 Date Utils
# ====================================================
def normalize_invoice_date(date_value):
    if not date_value:
        return ""

    current_year = datetime.now().year

    def fix_year(year):
        # พ.ศ. 4 หลัก เช่น 2569 -> 2026
        if year > 2400:
            year -= 543

        # Python/Azure ตีความปี 69 เป็น 1969
        elif 1969 <= year <= 1999:
            year += 57

        # Python ตีความปี 68 เป็น 2068 ให้ย้อนเป็น 2025
        if year > current_year + 1:
            year -= 43

        return year

    # กรณี Azure ส่งมาเป็น datetime/date
    if isinstance(date_value, datetime):
        year = fix_year(date_value.year)
        return date_value.replace(year=year).strftime("%d/%m/%Y")

    text = str(date_value).strip()
    text = re.sub(r"\s+", " ", text)
    text = text.replace(".", "")

    # กรณี dd/mm/yy หรือ dd-mm-yy เช่น 04/08/69
    m = re.match(r"^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2})$", text)
    if m:
        d, mth, yy = map(int, m.groups())

        # ถือว่า yy เป็น พ.ศ. 25xx เช่น 69 -> 2569 -> 2026
        year = (2500 + yy) - 543

        # กันปีอนาคตไกล เช่น 68 -> 2025
        if year > current_year + 1:
            year -= 43

        return f"{d:02d}/{mth:02d}/{year:04d}"

    formats = [
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%Y-%m-%d",
        "%d %b %Y",
        "%d %B %Y",
        "%d-%b-%Y",
        "%d-%B-%Y",
        "%d %b %y",
        "%d %B %y",
        "%d-%b-%y",
        "%d-%B-%y",
    ]

    for fmt in formats:
        try:
            dt = datetime.strptime(text, fmt)
            year = fix_year(dt.year)

            return dt.replace(year=year).strftime("%d/%m/%Y")

        except Exception:
            continue

    return text

#ใช้เฉพาะตอน InvoiceDate ว่างเท่านั้น
def extract_invoice_date_near_date_label(invoices):
    lines = get_all_lines(invoices)

    for line in lines:
        upper = line.upper()

        # ห้ามเอาวันครบกำหนด
        if "DUE DATE" in upper or "ครบกำหนด" in line:
            continue

        # เอาเฉพาะบรรทัดที่เป็นวันที่เอกสาร
        if "DATE" in upper or "วันที่" in line:
            m = re.search(r"\b\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}\b", line)
            if m:
                return normalize_invoice_date(m.group())

    return ""

def parse_date_safe(date_str):
    normalized = normalize_invoice_date(date_str)

    if not normalized:
        return None

    try:
        return datetime.strptime(normalized, "%d/%m/%Y")
    except Exception:
        return None

def extract_oldest_date_from_text(invoices):
    today = datetime.today()
    one_year_ago = today - timedelta(days=365)
    one_year_future = today + timedelta(days=365)

    raw_text = ""
    for page in invoices.pages:
        raw_text += " " + " ".join(
            [line.content for line in page.lines if line.content]
        )

    date_pattern = re.compile(
        r"\b("
        r"\d{4}-\d{2}-\d{2}"
        r"|\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}"
        r"|\d{1,2}[\s\-][A-Za-z]{3,9}[\s\-]\d{2,4}"
        r")\b"
    )

    candidates = []

    for date_str in date_pattern.findall(raw_text):
        normalized = normalize_invoice_date(date_str)
        parsed = parse_date_safe(normalized)

        if parsed and one_year_ago <= parsed <= one_year_future:
            candidates.append(parsed)

    if candidates:
        return min(candidates).strftime("%d/%m/%Y")

    return ""


# ====================================================
# 📌 Azure / PDF Utils
# ====================================================
def analyze_with_retry(client, model_id, pdf_path, max_retry=3, sleep_seconds=5):
    for attempt in range(1, max_retry + 1):
        try:
            with open(pdf_path, "rb") as f:
                poller = client.begin_analyze_document(
                    model_id=model_id,
                    body=f
                )
            return poller.result()

        except (ServiceRequestError, HttpResponseError) as e:
            print(f"⚠️ Azure error {model_id} attempt {attempt}/{max_retry}: {e}")

            if attempt == max_retry:
                print(f"❌ Skip page because Azure failed: {pdf_path}")
                return None

            time.sleep(sleep_seconds)


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


# ====================================================
# 📌 OCR Helper
# ====================================================
def get_field_value(field):
    if not field:
        return ""

    if hasattr(field, "value_string") and field.value_string:
        return field.value_string.strip()

    if hasattr(field, "value_date") and field.value_date:
        return normalize_invoice_date(field.value_date)

    if hasattr(field, "value_currency") and field.value_currency:
        return float(field.value_currency.amount)

    if hasattr(field, "value_number") and field.value_number is not None:
        return float(field.value_number)

    if hasattr(field, "content") and field.content:
        return str(field.content).strip()

    return ""


def normalize_number(value):
    try:
        return float(str(value).replace(",", "").strip())
    except Exception:
        return 0.0


def clean_tax_id(value):
    value = str(value or "")
    value = (
        value.replace("O", "0")
             .replace("o", "0")
             .replace("I", "1")
             .replace("l", "1")
    )
    return re.sub(r"\D", "", value)


def fix_tax_invoice_no_by_vendor(tax_invoice_no, vendor_tax_id):
    """
    แก้ OCR เฉพาะ Vendor ที่รู้ปัญหาแล้วเท่านั้น
    TPM Tax ID 0115539007424: OCR มักอ่าน No.S34716/69 เป็น No.834716/69
    """
    value = str(tax_invoice_no or "").strip()

    if vendor_tax_id == "0115539007424":
        # กรณีมี No. นำหน้า เช่น No.834716/69 -> No.S34716/69
        value = re.sub(
            r"(?i)\bNo\.?\s*8(?=\d{4,8}/\d{2,4})",
            "No.S",
            value
        )

        # กรณี Azure ตัดคำว่า No. ออก แล้วเหลือ 834716/69 -> S34716/69
        value = re.sub(
            r"^8(?=\d{4,8}/\d{2,4}$)",
            "S",
            value
        )

    return value


def get_all_lines(invoices):
    lines = []
    for page in invoices.pages:
        for line in page.lines:
            text = line.content.strip() if line.content else ""
            if text:
                lines.append(text)
    return lines


def get_all_text(invoices):
    return "\n".join(get_all_lines(invoices))


def append_msg(old_msg, new_msg):
    if old_msg and new_msg:
        return old_msg + " | " + new_msg
    return old_msg or new_msg


# ====================================================
# 📌 Pattern Matching Helpers
# ====================================================
def find_first_by_patterns(patterns, text, flags=re.IGNORECASE):
    for pattern in patterns:
        match = re.search(pattern, text, flags)
        if match:
            return match.group(1).strip() if match.lastindex else match.group(0).strip()
    return ""


def find_all_by_patterns(patterns, text, flags=re.IGNORECASE):
    results = []

    for pattern in patterns:
        matches = re.findall(pattern, text, flags)

        for m in matches:
            if isinstance(m, tuple):
                m = next((x for x in m if x), "")

            if m:
                results.append(str(m).strip())

    return results


# ====================================================
# 📌 Supplier / Vendor
# ====================================================
# def extract_supplier_name_from_pages(invoices):
#     lines = get_all_lines(invoices)

#     for line in lines[:20]:
#         upper = line.upper()

#         if any(k.upper() in upper for k in SUPPLIER_SKIP_KEYWORDS):
#             continue

#         for pattern in SUPPLIER_NAME_PATTERNS:
#             if re.search(pattern, upper, re.IGNORECASE):
#                 return line.strip().replace(";", ",")

#     return ""
def extract_supplier_name_from_pages(invoices):
    lines = get_all_lines(invoices)

    ssk_candidates = []

    for line in lines[:30]:
        text = line.strip()

        if "SSK" in text.upper() and re.search(r"PLASTIC|พลาสติก", text, re.IGNORECASE):
            ssk_candidates.append(text)

    if ssk_candidates:
        return max(ssk_candidates, key=len).replace(";", ",")

    full_text = "\n".join(lines).upper()

    # NIFCO
    if "NIFCO" in full_text or "นิฟโก้" in full_text:

        thai_name = ""
        eng_name = ""

        for line in lines:
            text = line.strip()

            if (
                "บริษัท" in text
                and "นิฟโก้" in text
                and "ได้รับ" not in text
                and "RECEIVED" not in text.upper()
                and "DIGITALLY" not in text.upper()
                and "THIS DOCUMENT" not in text.upper()
            ):
                thai_name = text
                break

        for line in lines:
            upper = line.upper()

            if (
                "UNION NIFCO" in upper
                and "DIGITALLY" not in upper
                and "THIS DOCUMENT" not in upper
                and "RECEIVED" not in upper
            ):
                eng_name = line.strip()

                # ถ้า LTD แยกไปอีกบรรทัด
                idx = lines.index(line)
                if idx + 1 < len(lines):
                    nxt = lines[idx + 1].strip().upper()

                    if nxt in ("LTD.", "LTD", "LIMITED"):
                        eng_name += " " + lines[idx + 1].strip()
                break

        if thai_name and eng_name:
            return f"{thai_name} {eng_name}".replace(";", ",")
        if thai_name:
            return thai_name.replace(";", ",")
        if eng_name:
            return eng_name.replace(";", ",")

    # Supplier ปกติ
    search_limit = 20

    for i, line in enumerate(lines[:search_limit]):
        upper = line.upper()

        if any(k.upper() in upper for k in SUPPLIER_SKIP_KEYWORDS):
            continue

        for pattern in SUPPLIER_NAME_PATTERNS:
            if re.search(pattern, upper, re.IGNORECASE):

                supplier = line.strip()

                # ถ้าบรรทัดก่อนหน้าเป็นภาษาอังกฤษสั้น ๆ เช่น SSK
                if i > 0:
                    prev = lines[i - 1].strip()

                    if re.fullmatch(r"[A-Z]{2,10}", prev):
                        supplier = prev + " " + supplier

                # ถ้าบรรทัดถัดไปเป็นชื่อภาษาไทย ให้ต่อเข้าไป
                if i + 1 < len(lines):
                    nxt = lines[i + 1].strip()

                    if "บริษัท" in nxt:
                        supplier += " " + nxt

                # return line.strip().replace(";", ",")
                return supplier.replace(";", ",")
    return ""

# def extract_supplier_name_from_pages(invoices):
#     lines = get_all_lines(invoices)

#     full_text = "\n".join(lines).upper()
#     isNifco = False
#     search_limit = 20

#     if "นิฟโก้" in full_text or "NIFCO" in full_text:
#         search_limit = 200
#         isNifco = True
    
#     for line in lines[:search_limit]:
#         upper = line.upper()

#         if any(k.upper() in upper for k in SUPPLIER_SKIP_KEYWORDS):
#             continue
#         if isNifco:
#             for pattern in SUPPLIER_NAME_PATTERNS_FOR_NIFCO:
#                 if re.search(pattern, upper, re.IGNORECASE):
#                     return line.strip().replace(";", ",")
#         else:
#             for pattern in SUPPLIER_NAME_PATTERNS:
#                 if re.search(pattern, upper, re.IGNORECASE):
#                     return line.strip().replace(";", ",")

#     return ""

def extract_vendor_branch(invoices, layout_result=None):
    lines = get_all_lines(invoices)
    full_text = "\n".join(lines)

    # SPECIAL CASE: T.KRUNGTHAI INDUSTRIES
    # บริษัทนี้ในหัวเอกสารมีรายการสาขา 00001/00002/00003 หลายบรรทัด
    # ห้ามดึงจากรายการสาขาด้านซ้าย ให้ดึงจากช่อง "สาขาที่" ข้างเลขที่ใบกำกับภาษีด้านขวาเท่านั้น
    if (
        "T.KRUNGTHAI INDUSTRIES" in full_text.upper()
        or "ที.กรุงไทยอุตสาหกรรม" in full_text
        or "กรุงไทยอุตสาหกรรม" in full_text
    ):
        branch_value = ""

        # 0.1) อ่านจาก layout โดยดูตำแหน่งด้านขวาของหน้า หรือบรรทัดที่มีเลขที่ใบกำกับภาษี
        if layout_result:
            for page in layout_result.pages:
                page_width = getattr(page, "width", 0) or 0

                for line in page.lines:
                    text = line.content.strip() if line.content else ""
                    if not text:
                        continue

                    m = re.search(r"สาขา\s*(?:ที่)?\s*[:：]?\s*(\d{1,10})", text, re.IGNORECASE)
                    if not m:
                        continue

                    x_min = 0
                    try:
                        xs = line.polygon[0::2]
                        x_min = min(xs) if xs else 0
                    except Exception:
                        x_min = 0

                    # เงื่อนไขหลัก: อยู่ด้านขวาของหน้า หรืออยู่บรรทัดเดียวกับคำว่า เลขที่
                    # เพื่อเลี่ยงรายการสาขาบริษัทด้านซ้ายบนเอกสาร
                    if (page_width and x_min >= page_width * 0.55) or re.search(r"เลขที่|No\.?", text, re.IGNORECASE):
                        branch_value = m.group(1)

                if branch_value:
                    return branch_value.zfill(5)

        # 0.2) fallback จาก OCR line: เลือกบรรทัดที่มีทั้ง เลขที่ และ สาขา
        for line in lines:
            if re.search(r"เลขที่|No\.?", line, re.IGNORECASE) and re.search(r"สาขา", line):
                m = re.search(r"สาขา\s*(?:ที่)?\s*[:：]?\s*(\d{1,10})", line, re.IGNORECASE)
                if m:
                    return m.group(1).zfill(5)

        # 0.3) fallback สุดท้าย: เอา occurrence ท้าย ๆ เพราะช่องสาขาใบกำกับภาษีมักอยู่หลังรายการสาขาด้านบน
        matches = re.findall(r"สาขา\s*(?:ที่)?\s*[:：]?\s*(\d{1,10})", full_text, re.IGNORECASE)
        if matches:
            return matches[-1].zfill(5)

    # 1) Checkbox selected: สำนักงานใหญ่ / สาขาที่
    if layout_result:
        for page in layout_result.pages:
            selected_marks = [
                m for m in (page.selection_marks or [])
                if m.state.name == "SELECTED"
            ]

            for mark in selected_marks:
                mark_y = mark.polygon[1]
                same_row_texts = []

                for line in page.lines:
                    text = line.content.strip() if line.content else ""
                    if not text:
                        continue

                    line_y = line.polygon[1]

                    if abs(mark_y - line_y) <= 0.08:
                        same_row_texts.append(text)

                same_row_text = " ".join(same_row_texts)

                if re.search(r"สำนักงานใหญ่|Head\s*Office|HeadOffice", same_row_text, re.IGNORECASE):
                    return "00000"

                value = find_first_by_patterns(VENDOR_BRANCH_PATTERNS, same_row_text)
                if value:
                    return value.zfill(5)

    # 2) ใบกำกับภาษีออกโดย : สำนักงานใหญ่ / สาขา
    if re.search(
        r"(?:ออกโดย|สาขาที่ออกใบกำกับภาษี)\s*[:：]?\s*(สำนักงานใหญ่|Head\s*Office|HeadOffice)",
        full_text,
        re.IGNORECASE,
    ):
        return "00000"

    m = re.search(
        r"(?:ออกโดย|สาขาที่ออกใบกำกับภาษี)\s*[:：]?\s*สาขา(?:ที่|เลขที่)?\s*(\d{1,10})",
        full_text,
        re.IGNORECASE,
    )
    if m:
        return m.group(1).zfill(5)

    # 3) Vendor zone only, stop before customer/TKC
    vendor_lines = []

    stop_patterns = [
        r"^\s*" + re.escape(k).replace(r"\ ", r"\s*")
        for k in STOP_KEYWORDS
    ]

    for line in lines:
        if any(re.search(p, line, re.IGNORECASE) for p in stop_patterns):
            break
        vendor_lines.append(line)

    vendor_text = "\n".join(vendor_lines)

    value = find_first_by_patterns(VENDOR_BRANCH_PATTERNS, vendor_text)
    if value:
        return value.zfill(5)

    if re.search(r"(สำนักงานใหญ่|Head\s*Office|HeadOffice)", vendor_text, re.IGNORECASE):
        return "00000"

    return ""


def extract_tax_id_from_pages(invoices):
    full_text = get_all_text(invoices)

    for pattern in TAXID_PATTERNS:
        for m in re.finditer(pattern, full_text, re.IGNORECASE | re.DOTALL):
            tax = clean_tax_id(m.group(1))

            if len(tax) >= 13:
                tax = tax[:13]

                if tax not in EXCLUDE_TAX_IDS:
                    return tax

    ids = re.findall(r"0\d{12}", clean_tax_id(full_text))

    for tax in ids:
        if tax not in EXCLUDE_TAX_IDS:
            return tax

    return ""

# def extract_vat_from_pages(invoices):
#     for page in invoices.pages:
#         for line in page.lines:
#             text = line.content.strip() if line.content else ""
#             clean_text = text.replace(",", "")

#             if not re.search(r"VAT\s*7\s*%", clean_text, re.IGNORECASE):
#                 continue

#             after_vat = re.split(
#                 r"VAT\s*7\s*%",
#                 clean_text,
#                 flags=re.IGNORECASE
#             )[-1]
#             nums = re.findall(r"\d+\.\d{2}", after_vat)

#             if nums:
#                 return float(nums[0])

#     return 0.0

def extract_vat_from_pages(invoices):

    full_text = "\n".join(
        line.content.strip()
        for page in invoices.pages
        for line in page.lines
        if line.content
    )

    full_text = full_text.replace(",", "")

    patterns = [
        r"SUB\s*TOTAL[\s\S]{0,150}?TOTAL\s*TAX[\s\S]{0,50}?VAT\s*7\s*%[\s\S]{0,50}?(\d+\.\d{2})",
        r"TOTAL\s+BEFORE\s+VAT[\s\S]{0,150}?VAT\s*7\s*%[\s\S]{0,50}?(\d+\.\d{2})",
        r"NET\s+AMOUNT[\s\S]{0,150}?VAT\s*7\s*%[\s\S]{0,50}?(\d+\.\d{2})",
        r"AMOUNT\s+BEFORE\s+VAT[\s\S]{0,150}?VAT\s*7\s*%[\s\S]{0,50}?(\d+\.\d{2})",
    ]

    for pattern in patterns:
        m = re.search(pattern, full_text, re.IGNORECASE)
        if m:
            return float(m.group(1))

    matches = re.findall(
        r"VAT\s*7\s*%[\s\S]{0,30}?(\d+\.\d{2})",
        full_text,
        re.IGNORECASE,
    )

    if matches:
        return float(matches[-1])

    return 0.0

def extract_po_from_text_and_tables(invoices):
    po_no_list = []

    for page in invoices.pages:
        for line in page.lines:
            text = line.content.strip() if line.content else ""
            po_no_list.extend(find_all_by_patterns(PO_PATTERNS, text))

    if hasattr(invoices, "tables"):
        for table in invoices.tables:
            for cell in table.cells:
                text = cell.content.strip() if cell.content else ""
                po_no_list.extend(find_all_by_patterns(PO_PATTERNS, text))

    cleaned = []

    for po in po_no_list:
        po = str(po).upper().replace("PO", "").strip()

        if po:
            cleaned.append(po)

    return ",".join(dict.fromkeys(cleaned))

def clean_po_from_field(purchase_order_no):
    po_no = str(purchase_order_no or "").replace("\n", ",")
    if not po_no:
        return ""

    po_no_list = []
    for p in [x.strip() for x in po_no.split(",") if x.strip()]:
        po_no_list.extend(find_all_by_patterns(PO_PATTERNS, p))

    cleaned = []
    for po in po_no_list:
        po = str(po).upper().replace("PO", "").strip()
        if po:
            cleaned.append(po)

    return ",".join(dict.fromkeys(cleaned))

def extract_tax_invoice_no_from_layout(layout_result, vendor_tax_id=""):
    lines = []

    for page in layout_result.pages:
        for line in page.lines:
            if line.content:
                lines.append(line.content.strip())

    full_text = " ".join(lines)
    full_text = re.sub(r"\s+", " ", full_text)

    # OCR Fix เฉพาะ VendorTaxId ที่กำหนด
    full_text = fix_tax_invoice_no_by_vendor(full_text, vendor_tax_id)

    value = find_first_by_patterns(INVOICE_PATTERNS, full_text)

    if value and not re.match(r"^(?:PO)?(?:410|140)\d{7}$", value, re.IGNORECASE):
        return value

    fallback_patterns = [
        r"Tax\s*Invoice\s*No\.?\s*[:：]?\s*([A-Za-z0-9\-\/]{5,30})",
        r"Invoice\s*No\.?\s*[:：]?\s*([A-Za-z0-9\-\/]{5,30})",
        r"เลขที่\s*[:：]?\s*([A-Za-z0-9\-\/]{5,30})",
        r"Account\s*No\.?\s*[:：]?\s*[A-Za-z0-9\-\/]+\s*No\.?\s*[:：]?\s*([A-Za-z]\d{6,20})",
        r"\bNo\.?\s*[:：]?\s*([A-Za-z]\d{6,20})",
    ]

    value = find_first_by_patterns(fallback_patterns, full_text)

    if "NIFCO" in full_text or "นิฟโก้" in full_text:
        value = "OTH".join(value)

    if value and not re.match(r"^(?:PO)?(?:410|140)\d{7}$", value, re.IGNORECASE):
        return value

    return ""

def find_invoice_no_from_words(invoices):
    for page in invoices.pages:
        if not hasattr(page, "words"):
            continue

        for w in page.words:
            text = w.content.strip() if w.content else ""

            if re.match(r"^(?:PO)?(?:410|140)\d{7}$", text, re.IGNORECASE):
                continue

            value = find_first_by_patterns(INVOICE_PATTERNS, text)

            if value:
                return value

    return ""

def normalize_amounts(row):
    try:
        check_invoice_No = row.get("TaxInvoiceNo","")
        if check_invoice_No:
            return row
        total_amount = normalize_number(row.get("TotalAmount", 0))
        vat_amount = normalize_number(row.get("VATAmount", 0))
        amount_inc_vat = normalize_number(row.get("AmountIncVat", 0))

        values = [total_amount, vat_amount, amount_inc_vat]

        if len(set(values)) < 3:
            msg = (
                f"Duplicate Amount Found "
                f"(Total={total_amount}, VAT={vat_amount}, AmountIncVat={amount_inc_vat})"
            )
            row["Emessage"] = append_msg(row.get("Emessage", ""), msg)
            return row

        values.sort()
        row["VATAmount"] = values[0]
        row["TotalAmount"] = values[1]
        row["AmountIncVat"] = values[2]

    except Exception:
        row["Emessage"] = append_msg(row.get("Emessage", ""), "Amount format invalid")

    return row

# 📌 Main Convert
def extract_invoice_to_json(invoice, invoices):

    supplier_name = get_field_value(invoice.fields.get("VendorAddressRecipient"))

    supplier_from_pages = extract_supplier_name_from_pages(invoices)

    if supplier_from_pages:
        if "SSK" in supplier_from_pages.upper() and "SSK" not in supplier_name.upper():
            supplier_name = supplier_from_pages
        elif not supplier_name or len(supplier_from_pages) > len(supplier_name):
            supplier_name = supplier_from_pages
    
    #fix บริฐัท SSK Plastic Co.,Ltd. Plastic Co.,Ltd. 
    supplier_name = re.sub(r"\s+", " ", supplier_name).strip()
    supplier_name = supplier_name.replace("SSK Plastic Co.,Ltd. Plastic Co.,Ltd.", "SSK Plastic Co.,Ltd. บริษัท เอส.เอส.เค พลาสติก จำกัด")

    supplier = (supplier_name or "").strip()

    has_company_th = "บริษัท" in supplier
    has_limited_th = "จำกัด" in supplier

    has_en = re.search(
        r"\b(COMPANY|LIMITED|LTD\.?)\b",
        supplier,
        re.IGNORECASE,
    )

    if (
        not supplier
        or (
            # มีภาษาไทย แต่ไม่ครบ
            (has_company_th or has_limited_th)
            and not (has_company_th and has_limited_th)
        )
        or (
            # ไม่มีภาษาไทยเลย และอังกฤษก็ไม่มี
            not has_company_th
            and not has_limited_th
            and not has_en
        )
    ):
        supplier_name = extract_supplier_name_from_pages(invoices)

    if supplier_name:
        supplier_name = supplier_name.replace(";", ",")

    # Normalize SSK Supplier Name
    supplier_name = supplier_name.strip()

    if supplier_name == "SSK Plastic Co.,Ltd.":
        supplier_name = "SSK Plastic Co.,Ltd. บริษัท เอส.เอส.เค พลาสติก จำกัด"

    invoice_date = normalize_invoice_date(
        get_field_value(invoice.fields.get("InvoiceDate"))
    )

    service_start = normalize_invoice_date(
        get_field_value(invoice.fields.get("ServiceStartDate"))
    )

    due_date = normalize_invoice_date(
        get_field_value(invoice.fields.get("DueDate"))
    )

    invoice_dt = parse_date_safe(invoice_date)
    service_dt = parse_date_safe(service_start)
    due_dt = parse_date_safe(due_date)
    current_year = datetime.now().year

    final_invoice_date = invoice_date

    if invoice_dt and invoice_dt.year != current_year:
        candidates = [d for d in [service_dt, due_dt] if d is not None]
        if candidates:
            final_invoice_date = min(candidates).strftime("%d/%m/%Y")
    elif not invoice_dt:
        candidates = [d for d in [service_dt, due_dt] if d is not None]
        if candidates:
            final_invoice_date = min(candidates).strftime("%d/%m/%Y")

    tax_invoice_no = str(
        get_field_value(invoice.fields.get("InvoiceId")) or ""
    ).replace("\n", ",")

    tax_invoice_no_clean = tax_invoice_no.split(",")[-1].strip()

    if re.match(r"^(?:PO)?(?:410|140)\d{7}$", tax_invoice_no_clean, re.IGNORECASE):
        tax_invoice_no_clean = ""

    purchase_order_no1 = str(
        get_field_value(invoice.fields.get("PurchaseOrder")) or ""
    ).replace("\n", ",")

    total_amount = get_field_value(invoice.fields.get("SubTotal"))
    vat_amount = get_field_value(invoice.fields.get("TotalTax"))
    amount_inc_vat = get_field_value(invoice.fields.get("InvoiceTotal"))

    vendor_tax_id = clean_tax_id(
        get_field_value(invoice.fields.get("VendorTaxId"))
    )

    if vendor_tax_id in EXCLUDE_TAX_IDS:
        vendor_tax_id = ""

    tax_invoice_no_clean = fix_tax_invoice_no_by_vendor(tax_invoice_no_clean, vendor_tax_id)

    return {
        "InvoiceDate": final_invoice_date,
        "PostingDate": final_invoice_date,
        "TaxInvoiceNo": tax_invoice_no_clean,
        "SupplierName": supplier_name,
        "Assignment": "",
        "VendorTaxId": vendor_tax_id,
        "VendorBranch": "",
        "TotalAmount": normalize_number(total_amount),
        "VATAmount": normalize_number(vat_amount),
        "AmountIncVat": normalize_number(amount_inc_vat),
        "PurchaseOrderNo": clean_po_from_field(purchase_order_no1),
        "Emessage": "",
    }


# def merge_invoice_row(existing, new):
#     for field in [
#         "InvoiceDate",
#         "PostingDate",
#         "TaxInvoiceNo",
#         "SupplierName",
#         "Assignment",
#         "VendorTaxId",
#         "VendorBranch",
#     ]:
#         if (not existing.get(field)) and new.get(field):
#             existing[field] = new.get(field)

#     vals = []
#     for v in [existing.get("PurchaseOrderNo", ""), new.get("PurchaseOrderNo", "")]:
#         if v:
#             vals.extend([x.strip() for x in str(v).split(",") if x.strip()])

#     if vals:
#         existing["PurchaseOrderNo"] = ",".join(dict.fromkeys(vals))

#     for field in ["TotalAmount", "VATAmount", "AmountIncVat"]:
#         try:
#             if normalize_number(existing.get(field, 0)) == 0 and normalize_number(new.get(field, 0)) > 0:
#                 existing[field] = new.get(field)
#         except Exception:
#             pass

#     if new.get("Emessage"):
#         existing["Emessage"] = append_msg(existing.get("Emessage", ""), new["Emessage"])

#     return existing

def merge_invoice_row(existing, new):

    for field in [
        "InvoiceDate",
        "PostingDate",
        "TaxInvoiceNo",
        "SupplierName",
        "Assignment",
        "VendorTaxId",
        "VendorBranch",
    ]:
        if (not existing.get(field)) and new.get(field):
            existing[field] = new.get(field)

    vals = []
    for v in [
        existing.get("PurchaseOrderNo", ""),
        new.get("PurchaseOrderNo", "")
    ]:
        if v:
            vals.extend(
                [x.strip() for x in str(v).split(",") if x.strip()]
            )

    if vals:
        existing["PurchaseOrderNo"] = ",".join(dict.fromkeys(vals))

    # TotalAmount และ AmountIncVat
    for field in ["TotalAmount", "AmountIncVat"]:
        try:
            if (
                normalize_number(existing.get(field, 0)) == 0
                and normalize_number(new.get(field, 0)) > 0
            ):
                existing[field] = new.get(field)
        except Exception:
            pass

    try:
        existing_vat = normalize_number(existing.get("VATAmount", 0))
        new_vat = normalize_number(new.get("VATAmount", 0))

        if new_vat > existing_vat:
            existing["VATAmount"] = new.get("VATAmount")
    except Exception:
        pass

    if new.get("Emessage"):
        existing["Emessage"] = append_msg(
            existing.get("Emessage", ""),
            new["Emessage"]
        )

    return existing

def add_or_merge_row(all_data, invoice_data):
    current_tax = (invoice_data.get("TaxInvoiceNo") or "").strip()

    if current_tax:
        for row in all_data:
            if (row.get("TaxInvoiceNo") or "").strip() == current_tax:
                merge_invoice_row(row, invoice_data)
                return

    current_assignment = invoice_data.get("Assignment", "")
    if not current_tax and all_data:
        last = all_data[-1]
        if last.get("Assignment") == current_assignment:
            merge_invoice_row(last, invoice_data)
            return

    all_data.append(invoice_data)

# 📌 Process PDF
pdf_list = [
    os.path.join(input_folder, f)
    for f in os.listdir(input_folder)
    if f.lower().endswith(".pdf")
]

if not pdf_list:
    print("❌ ไม่พบไฟล์ PDF ในโฟลเดอร์ Input")
    sys.exit()

all_data = []

for input_pdf in pdf_list:
    print("\n==============================")
    print(f"📁 Processing PDF File: {input_pdf}")
    print("==============================\n")

    pdf_files = []

    with open(input_pdf, "rb") as pdf_file:
        reader = PdfReader(pdf_file)

        for i, page in enumerate(reader.pages):
            single_page_path = os.path.join(temp_folder, f"page_{i + 1}.pdf")

            if os.path.exists(single_page_path):
                try:
                    os.remove(single_page_path)
                except PermissionError:
                    print(f"❌ Temp file locked, skip page: {single_page_path}")
                    continue

            writer = PdfWriter()
            writer.add_page(page)

            with open(single_page_path, "wb") as f:
                writer.write(f)

            if os.path.getsize(single_page_path) > 50 * 1024 * 1024:
                print(f"⚠️ Page {i + 1} too large → compressing...")
                compressed_path = os.path.join(temp_folder, f"page_{i + 1}_compressed.pdf")
                compress_pdf_page(single_page_path, compressed_path)
                pdf_files.append(compressed_path)
            else:
                pdf_files.append(single_page_path)

    for pdf_path in pdf_files:
        print(f"\n📄 Processing Page: {pdf_path}")

        invoices = analyze_with_retry(client, "prebuilt-invoice", pdf_path)
        if invoices is None:
            continue

        layout_result = analyze_with_retry(client, "prebuilt-layout", pdf_path)
        if layout_result is None:
            continue

        raw_page_text = get_all_text(invoices).lower()

        if "good receipt" in raw_page_text or "goods receipt" in raw_page_text:
            print("⏭️ พบคำว่า 'Good Receipt' → ข้ามหน้านี้ทันที")
            continue

        invoice_date_ocr = extract_oldest_date_from_text(invoices)

        for idx, invoice in enumerate(invoices.documents):

            invoice_data = extract_invoice_to_json(invoice, invoices)
            
            print("Azure InvoiceDate =", get_field_value(invoice.fields.get("InvoiceDate")))
            print("OCR InvoiceDate =", invoice_date_ocr)

            if not invoice_data.get("InvoiceDate") and invoice_date_ocr:
                invoice_data["InvoiceDate"] = invoice_date_ocr
                invoice_data["PostingDate"] = invoice_date_ocr

            if normalize_number(invoice_data.get("VATAmount", 0)) == 0:
                
                vat_fallback = extract_vat_from_pages(invoices)
                if vat_fallback:
                    invoice_data["VATAmount"] = vat_fallback

            if not invoice_data.get("VendorTaxId"):
                invoice_data["VendorTaxId"] = extract_tax_id_from_pages(invoices)

            # แก้ TaxInvoiceNo ที่ Azure prebuilt-invoice อ่านมาก่อนแล้ว เช่น No.834716/69
            invoice_data["TaxInvoiceNo"] = fix_tax_invoice_no_by_vendor(
                invoice_data.get("TaxInvoiceNo", ""),
                invoice_data.get("VendorTaxId", "")
            )

            invoice_data["VendorBranch"] = extract_vendor_branch(invoices, layout_result)

            # ==================================================
            # OCR Fix : TPM (Tax ID 0115539007424)
            # ==================================================
            if invoice_data.get("VendorTaxId") == "0115539007424":
                for page in layout_result.pages:
                    for line in page.lines:
                        if line.content:
                            line.content = re.sub(
                                r'(?i)\bNo\.\s*8(?=\d{4,8}/\d{2,4})',
                                "No.S",
                                line.content
                            )

            # ==================================================

            if not invoice_data.get("TaxInvoiceNo"):
                fallback_no = extract_tax_invoice_no_from_layout(layout_result, invoice_data.get("VendorTaxId", ""))
                if fallback_no:
                    invoice_data["TaxInvoiceNo"] = fallback_no
                    print(f"✅ Fallback TaxInvoiceNo from layout: {fallback_no}")

            if not invoice_data.get("TaxInvoiceNo"):
                fallback_no = find_invoice_no_from_words(invoices)
                if fallback_no:
                    invoice_data["TaxInvoiceNo"] = fix_tax_invoice_no_by_vendor(
                        fallback_no,
                        invoice_data.get("VendorTaxId", "")
                    )
                    print(f"✅ Fallback InvoiceId found from OCR: {invoice_data['TaxInvoiceNo']}")
                else:
                    print("⚠️ InvoiceId not found (even from OCR)")

            po_from_ocr = extract_po_from_text_and_tables(invoices)
            if po_from_ocr:
                invoice_data["PurchaseOrderNo"] = po_from_ocr

            tax_invoice_no = (invoice_data.get("TaxInvoiceNo") or "").strip()
            if re.match(r"^(?:PO)?(?:410|140)\d{7}$", tax_invoice_no, re.IGNORECASE):
                invoice_data["TaxInvoiceNo"] = ""

            # if not invoice_data.get("InvoiceDate") and invoice_date_ocr:
            #     invoice_data["InvoiceDate"] = invoice_date_ocr
            #     invoice_data["PostingDate"] = invoice_date_ocr

            invoice_data["InvoiceDate"] = normalize_invoice_date(invoice_data.get("InvoiceDate"))
            invoice_data["PostingDate"] = normalize_invoice_date(invoice_data.get("PostingDate"))
            invoice_data["Assignment"] = os.path.basename(input_pdf)

            normalize_amounts(invoice_data)
            add_or_merge_row(all_data, invoice_data)
        
        lines = get_all_lines(invoices)
        full_text = "\n".join(lines).upper()

        #Custom Nifco
        if "NIFCO" in full_text or "นิฟโก้" in full_text:
            invoice_data["TaxInvoiceNo"] = "OTH" + str(invoice_data.get("TaxInvoiceNo", ""))

    pdf_name = os.path.basename(input_pdf)
    dest_path = os.path.join(dest_folder, pdf_name)

    if os.path.exists(dest_path):
        print(f"⚠️ File already exists, skip move → {dest_path}")
    else:
        try:
            os.rename(input_pdf, dest_path)
            print(f"📁 Moved processed PDF → {dest_path}")
        except Exception as e:
            print(f"⚠️ Move failed: {e}")

# 💾 Save Excel
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

    for field in required_fields:
        value = row.get(field)

        if value is None or str(value).strip() == "":
            errors.append(f"{field} is empty")

    try:
        total_amount = normalize_number(row.get("TotalAmount", 0))
        vat_amount = normalize_number(row.get("VATAmount", 0))
        expected_vat = round(total_amount * 0.07, 2)

        if abs(vat_amount - expected_vat) > 0.01:
            errors.append(f"VATAmount ไม่ถูกต้อง (Expected {expected_vat:.2f})")

    except Exception:
        errors.append("VATAmount format invalid")

    row["Emessage"] = append_msg(row.get("Emessage", ""), " | ".join(errors))

df = pd.DataFrame(all_data)
df = df.reindex(columns=columns)
df.to_excel(output_excel, index=False)

print(f"\n✅ Done. All invoices saved to {output_excel}")