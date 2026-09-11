import os
import sys
import io
import re
import json
import time
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
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
from config import (
    AZURE_ENDPOINT,
    AZURE_KEY,
    get_paths
)


# ====================================================
# 📥 INPUT Parameter
# ====================================================

today_str = sys.argv[1]
branch_email = sys.argv[2]


# ====================================================
# 📁 PATH
# ====================================================

paths = get_paths(today_str, branch_email, "Process3")

input_folder = paths["input_folder"]
temp_folder = paths["temp_folder"]
output_excel = paths["output_excel"]
dest_folder = paths["dest_folder"]


# ====================================================
# 🔐 Azure
# ====================================================

endpoint = AZURE_ENDPOINT
key = AZURE_KEY
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
        {"name": "PO140", "regex": r"(?:PO)?(140\d{7})"},
        {"name": "P Invoice", "regex": r"(P\d{9})"},
        {"name": "PO41000", "regex": r"(41000\d{5})"},
        {"name": "PO43000", "regex": r"(43000\d{5})"},
        {"name": "PO44000", "regex": r"(44000\d{5})"},
        {"name": "PO45000", "regex": r"(45000\d{5})"},
        {"name": "PO46000", "regex": r"(46000\d{5})"},
        {"name": "E22000", "regex": r"(E22000\d{5})"},
        {"name": "P22000", "regex": r"(P22000\d{4})"},
        {"name": "T22000", "regex": r"(T22000\d{4})"},
        {"name": "K22000 PC LOCAL", "regex": r"(K22000\d{5})"}
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
        "BILL TO", "SHIP TO", "TO:", "Company", "Company Name",
        "Company Code", "รหัสลูกค้า", "ชื่อลูกค้า", "Delivery To",
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

# ====================================================
# 📌 PO Pattern Loader - รองรับทั้ง List แบบเดิม และ Dict แยก Branch
# ====================================================
RAW_PO_CONFIG = pattern_config.get("po_patterns", DEFAULT_PATTERN_CONFIG["po_patterns"])

# ใช้เก็บ config PO ที่แยกตาม Company Branch เช่น {"0000": [...], "1000": [...]}
PO_PATTERN_CONFIG = RAW_PO_CONFIG if isinstance(RAW_PO_CONFIG, dict) else {}

# PO_PATTERNS ใช้เป็น generic fallback และรองรับ clean_po_from_field() เดิม
PO_PATTERNS = []

if isinstance(RAW_PO_CONFIG, list):
    for p in RAW_PO_CONFIG:
        if isinstance(p, dict) and p.get("regex"):
            PO_PATTERNS.append(p["regex"])

elif isinstance(RAW_PO_CONFIG, dict):
    for branch_items in RAW_PO_CONFIG.values():
        if not isinstance(branch_items, list):
            continue
        for p in branch_items:
            if isinstance(p, dict) and p.get("regex"):
                PO_PATTERNS.append(p["regex"])

# ถ้า JSON ไม่มี PO ที่ใช้ได้ ให้ fallback กลับไป Default Pattern
if not PO_PATTERNS:
    PO_PATTERNS = [
        p["regex"]
        for p in DEFAULT_PATTERN_CONFIG["po_patterns"]
        if isinstance(p, dict) and p.get("regex")
    ]

VENDOR_BRANCH_PATTERNS = [p["regex"] for p in pattern_config.get("vendor_branch_patterns", DEFAULT_PATTERN_CONFIG["vendor_branch_patterns"])]
TAXID_PATTERNS = [p["regex"] for p in pattern_config.get("taxid_patterns", DEFAULT_PATTERN_CONFIG["taxid_patterns"])]
EXCLUDE_TAX_IDS = set(pattern_config.get("exclude_tax_ids", DEFAULT_PATTERN_CONFIG["exclude_tax_ids"]))
STOP_KEYWORDS = pattern_config.get("stop_keywords", DEFAULT_PATTERN_CONFIG["stop_keywords"])
SUPPLIER_SKIP_KEYWORDS = pattern_config.get("supplier_skip_keywords", DEFAULT_PATTERN_CONFIG["supplier_skip_keywords"])
SUPPLIER_NAME_PATTERNS = [p["regex"] for p in pattern_config.get("supplier_name_patterns", DEFAULT_PATTERN_CONFIG["supplier_name_patterns"])]

# ====================================================
# 📌 PO Branch Rules
# ====================================================
# 0000 = บางพลี สำนักงานใหญ่
# 1000 = ปราจีน สาขา 1
#
# หมายเหตุ:
# - ค้นหา PO ของ branch ที่ส่งเข้ามาก่อน
# - ถ้าไม่เจอ จึงค้นหา PO ของอีก branch
# - ถ้าเจอ PO ของอีก branch จะเก็บ PO ไว้ และเพิ่ม Emessage
#   "PO does not match Company branch"
PO_BRANCH_PATTERNS = {
    "0000": [
        r"(41000\d{5})",
        r"(45000\d{5})",
        r"(P22000\d{4})",
        r"(T22000\d{4})",
    ],
    "1000": [
        r"(43000\d{5})",
        r"(44000\d{5})",
        r"(E22000\d{5})",
        r"(K22000\d{5})",
        r"(46000\d{5})",
    ],
}


def get_po_patterns_by_branch(branch_code):
    """
    คืน PO regex ของ Company Branch ที่ระบุ

    Priority:
    1) patternsInvoice.json ถ้า po_patterns เป็น dict แยก branch
    2) PO_BRANCH_PATTERNS ในโปรแกรมเป็น fallback
    """
    branch_code = str(branch_code or "").strip().zfill(4)

    # ใช้ config จาก patternsInvoice.json ก่อน
    if isinstance(PO_PATTERN_CONFIG, dict):
        branch_items = PO_PATTERN_CONFIG.get(branch_code, [])
        patterns = [
            p["regex"]
            for p in branch_items
            if isinstance(p, dict) and p.get("regex")
        ]
        if patterns:
            return patterns

    # fallback กฎที่กำหนดไว้ในโปรแกรม
    return PO_BRANCH_PATTERNS.get(branch_code, [])


def get_other_branch(branch_code):
    branch_code = str(branch_code or "").strip().zfill(4)

    if branch_code == "0000":
        return "1000"
    if branch_code == "1000":
        return "0000"

    return ""

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

# ใช้หา "วันที่เอกสาร / Invoice Date" จาก OCR โดยตรง
# จุดสำคัญ: ห้ามนำ Due Date / PO Date / Delivery Date มาเป็น InvoiceDate
def extract_invoice_date_near_date_label(invoices, ocr_cache=None):
    lines = ocr_cache["lines"] if ocr_cache is not None else get_all_lines(invoices)

    # จำกัดช่วงด้านบนของเอกสารก่อน เพื่อลดโอกาสไปเจอวันที่รับสินค้า/วันที่เซ็นด้านล่าง
    search_lines = lines[:80]

    # label ที่ไม่ใช่ Invoice Date
    exclude_patterns = [
        r"\bDUE\s*DATE\b",
        r"\bPO\s*DATE\b",
        r"\bP\.O\.\s*DATE\b",
        r"\bDELIVERY\s*DATE\b",
        r"\bRECEIVE(?:D|R)?\s*DATE\b",
        r"\bPAYMENT\s*DATE\b",
        r"\bSERVICE\s*DATE\b",
        r"วันครบกำหนด",
        r"กำหนดชำระ",
        r"วันที่ส่ง",
        r"วันที่รับ",
    ]

    # Priority 1: บรรทัดที่ระบุ Invoice Date / วันที่-Date ชัดเจน
    preferred_patterns = [
        r"\bINVOICE\s*DATE\b",
        r"วันที่\s*/\s*DATE",
        r"\bDATE\s*[:：]",
        r"วันที่\s*[:：]",
    ]

    date_pattern = re.compile(
        r"\b(?:"
        r"\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}"
        r"|\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}"
        r")\b"
    )

    def is_excluded(line):
        return any(
            re.search(pattern, line, re.IGNORECASE)
            for pattern in exclude_patterns
        )

    # รอบแรก: เอาเฉพาะ label ที่ชัดเจนที่สุด
    for line in search_lines:
        text = re.sub(r"\s+", " ", str(line or "")).strip()

        if not text or is_excluded(text):
            continue

        if any(re.search(pattern, text, re.IGNORECASE) for pattern in preferred_patterns):
            m = date_pattern.search(text)
            if m:
                return normalize_invoice_date(m.group(0))

    # รอบสอง: fallback สำหรับ template ที่เขียนเพียง DATE / วันที่
    for line in search_lines:
        text = re.sub(r"\s+", " ", str(line or "")).strip()
        upper = text.upper()

        if not text or is_excluded(text):
            continue

        # หลีกเลี่ยงวันที่ที่เกี่ยวกับ PO แม้ OCR จะแยกคำไม่สวย
        if re.search(r"\bP\.?\s*O\.?\b", upper):
            continue

        if "DATE" in upper or "วันที่" in text:
            m = date_pattern.search(text)
            if m:
                return normalize_invoice_date(m.group(0))

    return ""

def parse_date_safe(date_str):
    normalized = normalize_invoice_date(date_str)

    if not normalized:
        return None

    try:
        return datetime.strptime(normalized, "%d/%m/%Y")
    except Exception:
        return None

def extract_oldest_date_from_text(invoices, ocr_cache=None):
    today = datetime.today()
    one_year_ago = today - timedelta(days=365)
    one_year_future = today + timedelta(days=365)

    if ocr_cache is not None:
        # เดิม function รวมทุก line ด้วยช่องว่าง จึงแปลง newline ใน cache เป็นช่องว่าง
        raw_text = " " + ocr_cache["full_text"].replace("\n", " ")
    else:
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


def get_address_value(field):
    """
    อ่าน Address จาก Azure Document Intelligence

    Priority:
    1. field.content -> เก็บ Address ตามข้อความจริงบนเอกสาร
    2. value_address -> fallback กรณี content ไม่มี

    ช่วยป้องกัน:
    - แขวง/เขตหาย
    - Address ไม่ครบ
    - house_number / street_address ซ้ำกัน
    """

    if not field:
        return ""

    # ==========================================================
    # 1. ใช้ OCR content ก่อน
    # ==========================================================
    try:
        if hasattr(field, "content") and field.content:

            content = str(field.content).strip()

            # รวม newline เป็นช่องว่าง
            content = re.sub(r"\s+", " ", content)

            # ทำความสะอาด separator
            content = content.strip(" ,;:-")

            if content:
                return content

    except Exception:
        pass

    # ==========================================================
    # 2. Fallback -> Azure structured address
    # ==========================================================
    try:
        if hasattr(field, "value_address") and field.value_address:

            addr = field.value_address

            parts = []

            # เรียงจากละเอียด -> กว้าง
            attributes = [
                "house_number",
                "road",
                "street_address",
                "unit",
                "city_district",
                "city",
                "state",
                "postal_code",
                "country_region",
            ]

            for attr in attributes:

                value = getattr(addr, attr, None)

                if not value:
                    continue

                value = str(value).strip()

                if not value:
                    continue

                # ----------------------------------------------
                # ป้องกันค่าซ้ำ
                #
                # ตัวอย่าง:
                # house_number = 607
                # street_address = 607 ถนนอโศก
                #
                # ไม่ต้องเก็บ 607 ซ้ำ
                # ----------------------------------------------
                duplicate = False

                for existing in parts:

                    existing_lower = existing.casefold()
                    value_lower = value.casefold()

                    if (
                        value_lower == existing_lower
                        or value_lower in existing_lower
                        or existing_lower in value_lower
                    ):
                        # ถ้า value ใหม่ละเอียดกว่า ให้แทนค่าเก่า
                        if len(value) > len(existing):

                            try:
                                index = parts.index(existing)
                                parts[index] = value
                            except Exception:
                                pass

                        duplicate = True
                        break

                if not duplicate:
                    parts.append(value)

            if parts:
                return ", ".join(parts)

    except Exception:
        pass

    return ""


def clean_company_address_start(address):
    """
    Clean CompanyAddress:
    1. ตัดข้อความขยะก่อนเลขที่บ้าน
    2. ลบ OCR คำว่า Address ที่ติดหน้า Samutprakarn/Samutprakan
    3. ตัดข้อความหลัง postcode 10570 หรือ 25140
    4. ถ้าหาเลขที่บ้านไม่ได้ ให้เก็บข้อความเดิมไว้
    """

    if not address:
        return ""

    address = re.sub(r"\s+", " ", str(address)).strip(" ,;:-")

    # ----------------------------------------------------------
    # ลบ OCR noise:
    # AddressSamutprakarn -> Samutprakarn
    # Address Samutprakarn -> Samutprakarn
    # ----------------------------------------------------------
    address = re.sub(
        r"(?i)\bAddress\s*(?=Samut\s*prak(?:arn|an))",
        "",
        address
    )

    address = re.sub(r"\s+", " ", address).strip(" ,;:-")

    # ----------------------------------------------------------
    # หาเลขที่บ้าน
    # เช่น 370 / 555 / 99/9 / 123/456
    # ----------------------------------------------------------
    matches = list(
        re.finditer(
            r"\b\d{1,4}(?:/\d{1,5})?\b",
            address
        )
    )

    cleaned = address

    for match in matches:
        value = match.group(0)

        # ป้องกันไม่ให้ postcode ถูกมองเป็นเลขที่บ้าน
        if value in ("10570", "25140"):
            continue

        cleaned = address[match.start():]
        break

    # ----------------------------------------------------------
    # ตัดทุกอย่างหลัง postcode
    # รองรับ:
    # 10570
    # , 10570,
    # 25140 THAILAND
    # ----------------------------------------------------------
    postcode_match = re.search(
        r"\b(?:10570|25140)\b",
        cleaned
    )

    if postcode_match:
        cleaned = cleaned[:postcode_match.end()]

    cleaned = re.sub(r"\s+", " ", cleaned)
    cleaned = cleaned.strip(" ,;:-.")

    return cleaned

def validate_Company_branch(Company_address, input_branch):
    """
    ตรวจสอบ Company Branch จากที่อยู่ เทียบกับค่าที่ส่งเข้ามาตอนรันโปรแกรม

    ตัวอย่าง:
    python script.py testfile 0000

    ถ้า CompanyAddress มีเลข 370
        expected = 0000

    ถ้าไม่มี 370
        expected = 1000

    return:
        Company_branch
        is_correct
        expected_branch
    """

    Company_address = str(Company_address or "").strip()
    input_branch = str(input_branch or "").strip().zfill(4)

    # กำหนด branch ที่ควรจะเป็นจาก Address
    if re.search(r"\b(?:370|10570)\b", Company_address):
        expected_branch = "0000"
    else:
        expected_branch = "1000"

    is_correct = input_branch == expected_branch

    return input_branch, is_correct, expected_branch

def extract_vendor_address_from_pages(invoices, ocr_cache=None):
    """
    OCR fallback สำหรับดึง Address ของ Vendor จากข้อความบน Invoice
    """
    lines = ocr_cache["lines"] if ocr_cache is not None else get_all_lines(invoices)
    address_lines = []
    found_address = False

    stop_keywords = [
        "TAX ID", "TAXID", "VAT ID", "VATID",
        "TEL", "FAX", "ACCOUNT", "SHIP TO", "BILL TO",
        "ATTN", "DATE", "PAGE", "INVOICE",
        "ORIGINAL TAX INVOICE", "TERM OF PAYMENT", "REMARK"
    ]

    for line in lines:
        text = line.strip()
        if not text:
            continue

        # Address อยู่บรรทัดเดียวกับ label
        m = re.match(r"^(?:Address|ที่อยู่)\s*[:：]?\s*(.+)$", text, re.IGNORECASE)
        if m:
            found_address = True
            value = m.group(1).strip()
            if value:
                address_lines.append(value)
            continue

        # Address เป็น label เดี่ยว
        if re.fullmatch(r"(?:Address|ที่อยู่)\s*[:：]?", text, re.IGNORECASE):
            found_address = True
            continue

        if not found_address:
            continue

        upper = text.upper()

        if any(k in upper for k in stop_keywords):
            break

        address_lines.append(text)

        if len(address_lines) >= 4:
            break

    cleaned = []
    for line in address_lines:
        line = re.sub(r"\s+", " ", line).strip()
        if line and line not in cleaned:
            cleaned.append(line)

    return " ".join(cleaned).replace(";", ",").strip()


def extract_Company_address_from_pages(invoices, ocr_cache=None):
    """
    OCR fallback สำหรับดึง Address ของ Company / Account To / Bill To
    โดยจะเริ่มอ่านหลังพบโซนลูกค้า และหยุดก่อน Ship To / Tax ID / Tel / ส่วนถัดไป
    """
    lines = ocr_cache["lines"] if ocr_cache is not None else get_all_lines(invoices)
    address_lines = []
    found_Company_zone = False
    found_address = False

    Company_zone_patterns = [
        r"\bACCOUNT\s+TO\b",
        r"\bBILL\s+TO\b",
        r"ชื่อ.*ลูกค้า",
        r"ชื่อลูกค้า",
        r"ชื่อผู้ซื้อ",
        r"ผู้ซื้อ",
    ]

    stop_patterns = [
        r"\bSHIP\s+TO\b",
        r"\bDELIVERY\s+TO\b",
        r"\bATTN\b",
        r"\bTAX\s*ID\b",
        r"\bTAXID\b",
        r"\bVAT\s*ID\b",
        r"\bTEL\b",
        r"\bFAX\b",
        r"TERM\s+OF\s+PAYMENT",
        r"REMARK",
        r"ORIGINAL\s+TAX\s+INVOICE",
    ]

    for line in lines:
        text = re.sub(r"\s+", " ", line.strip())
        if not text:
            continue

        upper = text.upper()

        if not found_Company_zone:
            if any(re.search(p, text, re.IGNORECASE) for p in Company_zone_patterns):
                found_Company_zone = True

                # บางใบมี Address อยู่ในบรรทัด Account To เดียวกัน
                m_inline = re.search(
                    r"(?:Address|ที่อยู่)\s*[:：]?\s*(.+)$",
                    text,
                    re.IGNORECASE,
                )
                if m_inline:
                    value = m_inline.group(1).strip()
                    if value:
                        address_lines.append(value)
                        found_address = True
            continue

        # เมื่อเข้าพื้นที่ลูกค้าแล้ว ให้เริ่มจาก label Address/ที่อยู่
        m = re.match(r"^(?:Address|ที่อยู่)\s*[:：]?\s*(.*)$", text, re.IGNORECASE)
        if m:
            found_address = True
            value = m.group(1).strip()
            if value:
                address_lines.append(value)
            continue

        if not found_address:
            continue

        # หยุดเมื่อเจอส่วนถัดไป
        if any(re.search(p, text, re.IGNORECASE) for p in stop_patterns):
            break

        address_lines.append(text)

        # ปกติ address 1-3 บรรทัดก็เพียงพอ
        if len(address_lines) >= 4:
            break

    cleaned = []
    for line in address_lines:
        line = re.sub(r"\s+", " ", line).strip(" ,;:-")
        if line and line not in cleaned:
            cleaned.append(line)

    return " ".join(cleaned).replace(";", ",").strip()


def clean_tax_id(value):
    value = str(value or "")
    value = (
        value.replace("O", "0")
             .replace("o", "0")
             .replace("I", "1")
             .replace("l", "1")
    )
    return re.sub(r"\D", "", value)


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


def build_ocr_cache(invoices):
    """
    สร้าง OCR cache หนึ่งครั้งต่อหน้า หลัง Azure prebuilt-invoice เสร็จ

    จุดประสงค์:
    - ลดการวน invoices.pages / invoices.tables ซ้ำหลายรอบ
    - ไม่เปลี่ยนข้อมูลที่ใช้ในการ extract จึงคง logic/accuracy เดิม
    - Azure ยังคงเรียก prebuilt-invoice และ prebuilt-layout เหมือนเดิม
    """
    lines = get_all_lines(invoices)
    full_text = "\n".join(lines)

    table_texts = []
    if hasattr(invoices, "tables"):
        for table in invoices.tables:
            for cell in table.cells:
                text = cell.content.strip() if cell.content else ""
                if text:
                    table_texts.append(text)

    table_text = "\n".join(table_texts)

    # รักษาลำดับเดิมของ PO search: OCR Lines -> OCR Tables
    if full_text and table_text:
        po_base_text = full_text + "\n" + table_text
    else:
        po_base_text = full_text or table_text

    return {
        "lines": lines,
        "full_text": full_text,
        "full_text_upper": full_text.upper(),
        "full_text_lower": full_text.lower(),
        "table_texts": table_texts,
        "table_text": table_text,
        "po_base_text": po_base_text,
    }


def normalize_vendor_specific_invoice_no(invoice_data, full_text=""):
    """
    Final vendor-specific normalization for TaxInvoiceNo.

    TPM rule:
    - VendorTaxId 0115539007424, or TPM / THAI PRESS AND MACHINERY in vendor text
    - OCR often reads the printed letter S after "No." as digit 8
    - Example: No.834716/69 -> No.S34716/69

    This function is intentionally called at the LAST step before add_or_merge_row(),
    so later extraction/fallback logic cannot overwrite the corrected value.
    """
    if not isinstance(invoice_data, dict):
        return invoice_data

    invoice_no = str(invoice_data.get("TaxInvoiceNo", "") or "").strip()
    if not invoice_no:
        return invoice_data

    vendor_tax_id = clean_tax_id(invoice_data.get("VendorTaxId", ""))
    supplier_name = str(invoice_data.get("SupplierName", "") or "")
    detect_text = f"{supplier_name}\n{full_text or ''}".upper()

    is_tpm = (
        vendor_tax_id == "0115539007424"
        or "THAI PRESS AND MACHINERY" in detect_text
        or "ไทยเพรส แอนด์ แมชชีนเนอรี่" in detect_text
        or re.search(r"(?:^|\s)TPM(?:\s|$)", detect_text) is not None
    )

    if not is_tpm:
        return invoice_data

    original = invoice_no

    # Normalize OCR separators/spaces but preserve the invoice number itself.
    # Accepted examples:
    #   No.834716/69
    #   No 834716/69
    #   No.: 834716/69
    #   NO. 834716/69
    m = re.fullmatch(
        r"(?i)NO\.?\s*[:：]?\s*8\s*(\d{4,10}\s*/\s*\d{2,4})",
        invoice_no,
    )

    if m:
        suffix = re.sub(r"\s+", "", m.group(1))
        invoice_no = f"No.S{suffix}"
        invoice_data["TaxInvoiceNo"] = invoice_no
        print(f"✅ TPM FINAL NORMALIZE: {original} -> {invoice_no}")
    else:
        # If Azure already read S correctly, only normalize the prefix formatting.
        m = re.fullmatch(
            r"(?i)NO\.?\s*[:：]?\s*S\s*(\d{4,10}\s*/\s*\d{2,4})",
            invoice_no,
        )
        if m:
            suffix = re.sub(r"\s+", "", m.group(1))
            invoice_no = f"S{suffix}"
            invoice_data["TaxInvoiceNo"] = invoice_no
            print(f"✅ TPM InvoiceNo already S: {original} -> {invoice_no}")
        else:
            print(
                f"⚠️ TPM detected but InvoiceNo format not matched: {repr(original)}"
            )

    return invoice_data


def append_msg(old_msg, new_msg):
    """
    รวม Emessage โดยไม่ให้ข้อความซ้ำกัน

    ตัวอย่าง:
    old = "Company address does not match | PO does not match Company branch"
    new = "PO does not match Company branch"

    result:
    "Company address does not match | PO does not match Company branch"
    """

    old_msg = str(old_msg or "").strip()
    new_msg = str(new_msg or "").strip()

    messages = []

    # รวมทั้งข้อความเดิมและใหม่
    for source in [old_msg, new_msg]:

        if not source:
            continue

        # แยกด้วย |
        for msg in source.split("|"):

            msg = re.sub(r"\s+", " ", msg).strip()

            if not msg:
                continue

            # ตรวจซ้ำแบบไม่สนตัวพิมพ์ใหญ่/เล็ก
            if not any(
                msg.casefold() == existing.casefold()
                for existing in messages
            ):
                messages.append(msg)

    return " | ".join(messages)

def extract_tax_remark(invoices, ocr_cache=None):

    full_text = (
        ocr_cache["full_text"]
        if ocr_cache is not None
        else get_all_text(invoices)
    )

    normalized_text = re.sub(r"[.\s]", "", full_text).upper()

    tax_patterns = [
        "รายการที่ไม่สามารถหักภาษี ณ ที่จ่าย",
        "รายการที่ไม่สามารถหักภาษีณ.ที่จ่ายได้",
        "ลูกค้าจึงไม่มีหน้าที่ หัก ภาษี ณ ที่จ่าย",
        "จะต้องถูกหักภาษี ณ ที่จ่าย",
        "ไม่สามารถหักภาษี ณ ที่จ่ายได้",
        "ไม่สามารถหัก ณ ที่จ่ายได้",
        "ไม่สามารถหัก ณ. ที่จ่าย",
        "ไม่สามารถหักภาษี ณ ที่จ่าย",
        "ไม่สามารถหัก ภาษี ณ ที่จ่าย",
        "ห้ามหักภาษี ณ ที่จ่าย",
        "NO DEDUCT WITH HOLDING TAX",
        "NO WITH HOLDING TAX",
        "(NO.WHT)",
        "ไม่หัก ณ ที่จ่าย",
        "ไม่หักภาษี ณ ที่จ่าย",
        "ให้หักภาษี ณ ที่จ่าย",
        "กรุณาอย่าหักภาษีหัก ณ ที่จ่าย",
        "กรุณาอย่าหักภาษี ณ ที่จ่าย",
        "กรุณาอย่าหัก ณ ที่จ่าย",
        "หัก ณ ที่จ่ายทั้งหมด",
        "ไม่ต้องหักภาษี ณ. ที่จ่าย",
        "ไม่ต้องหักณที่จ่าย",
        "สามารถหักภาษี ณ ที่จ่าย ได้",
        "สามารถหักภาษีณ.ที่จ่าย",
        "หักภาษี ณ ที่จ่ายไม่ได้",
        "หักภาษี ณ ที่จ่ายได้",
        "หัก ณ. ที่จ่าย",
        "หัก ณ ที่จ่าย",
        "หัก ภาษี ณ ที่จ่าย",
        "หักภาษี ณ ที่จ่าย",
    ]

    tax_patterns = sorted(
        tax_patterns,
        key=lambda x: len(re.sub(r"[.\s]", "", x)),
        reverse=True
    )

    for pattern in tax_patterns:

        normalized_pattern = re.sub(
            r"[.\s]",
            "",
            pattern
        ).upper()

        # เช็คแบบ normalized ก่อน
        if normalized_pattern in normalized_text:

            regex_parts = []

            for char in pattern:

                # รองรับ whitespace ทุกชนิด
                if char.isspace() or char == ".":
                    regex_parts.append(r"[.\s]*")
                else:
                    regex_parts.append(re.escape(char))

            flexible_pattern = "".join(regex_parts)

            match = re.search(
                flexible_pattern,
                full_text,
                re.IGNORECASE
            )

            if match:

                remark = match.group(0)
                return remark

    return ""

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


def remove_supplier_logo_prefix(name):
    """
    ถ้ามีข้อความใด ๆ อยู่ก่อนคำว่า 'บริษัท'
    ให้ตัดข้อความด้านหน้าทั้งหมดออก

    ตัวอย่าง:
    CHO บริษัท บิโก้ จำกัด
    -> บริษัท บิโก้ จำกัด

    KAT บริษัท ที.กรุงไทยอุตสาหกรรม จำกัด
    -> บริษัท ที.กรุงไทยอุตสาหกรรม จำกัด

    TPM บริษัท ไทยเพรส แอนด์ แมชชีนเนอรี่ โปรดักส์ จำกัด
    -> บริษัท ไทยเพรส แอนด์ แมชชีนเนอรี่ โปรดักส์ จำกัด

    บริษัท ไทย อะคิบะ จำกัด
    -> บริษัท ไทย อะคิบะ จำกัด
    """

    if not name:
        return ""

    name = re.sub(r"\s+", " ", str(name)).strip(" ,;:-")

    # หา "บริษัท" ตัวแรก
    match = re.search(r"บริษัท", name)

    if match and match.start() > 0:
        original = name

        # เก็บตั้งแต่คำว่า "บริษัท" เป็นต้นไป
        name = name[match.start():].strip()

        print(
            f"🧹 Supplier prefix removed: "
            f"{original} -> {name}"
        )

    return name



def clean_supplier_legal_ending(name):
    """
    จัด SupplierName ให้เหลือเฉพาะชื่อบริษัท

    - เก็บชื่อไทยจนถึง จำกัด / จำกัด (มหาชน)
    - ตัด (สำนักงานใหญ่), (สาขา...)
    - ตัด (HEAD OFFICE), (BRANCH...)
    - ถ้ามีชื่ออังกฤษ ให้เก็บจนถึง legal ending
    """

    if not name:
        return ""

    name = re.sub(r"\s+", " ", str(name)).strip(" ,;:-")

    # ตัดข้อมูล Branch / Head Office
    name = re.sub(
        r"\s*\(\s*(?:"
        r"สำนักงานใหญ่"
        r"|HEAD\s*OFFICE"
        r"|สาขา(?:ที่)?\s*[^)]*"
        r"|BRANCH(?:\s+NO\.?)?\s*[^)]*"
        r")\s*\)",
        "",
        name,
        flags=re.IGNORECASE,
    )

    name = re.sub(r"\s+", " ", name).strip(" ,;:-")

    # ใช้เฉพาะกรณีที่มีชื่อบริษัทภาษาไทย
    thai_start = re.search(r"บริษัท", name)

    if not thai_start:
        return name

    # หา ending ของชื่อบริษัทไทย
    thai_legal_pattern = re.compile(
        r"""
        บริษัท
        .*?
        จำกัด
        (?:
            \s*\(\s*มหาชน\s*\)
        )?
        """,
        re.IGNORECASE | re.VERBOSE,
    )

    thai_match = thai_legal_pattern.search(name, thai_start.start())

    if not thai_match:
        return name

    thai_name = thai_match.group(0).strip(" ,;:-")
    tail = name[thai_match.end():].strip(" ,;:-")

    if not tail:
        return thai_name

    english_suffix = re.compile(
        r"""
        (?:
            PUBLIC\s+COMPANY\s+LIMITED
            |
            COMPANY\s+LIMITED
            |
            CO\.?\s*,?\s*LTD\.?
            |
            LTD\.?
            |
            LIMITED
        )
        """,
        re.IGNORECASE | re.VERBOSE,
    )

    matches = list(english_suffix.finditer(tail))

    if not matches:
        return thai_name

    candidates = []
    segment_start = 0

    for m in matches:
        segment_end = m.end()

        candidate = tail[segment_start:segment_end].strip(" ,;:-")

        candidate = re.sub(
            r"^[|/\\\-–—:;,.\s]+",
            "",
            candidate
        ).strip()

        if candidate:
            candidates.append(candidate)

        segment_start = segment_end

    if not candidates:
        return thai_name

    def candidate_score(value):
        alnum_len = len(re.sub(r"[^A-Za-z0-9]", "", value))
        word_count = len(re.findall(r"[A-Za-z0-9]+", value))

        return (alnum_len, word_count, len(value))

    best_english = max(candidates, key=candidate_score)

    return re.sub(
        r"\s+",
        " ",
        f"{thai_name} {best_english}"
    ).strip()


def clean_supplier_name(name):
    """
    ทำความสะอาด Supplier Name แบบ Generic
    - ตัด Supplier Logo / Prefix ที่กำหนดไว้
    - ตัด ISO / Certification / Website / Contact / Tax ID / Address
    - แก้ OCR กรณีคำว่า 'บริษัท' ด้านหน้าหายเป็น 'ษัท'
    - ลบ phrase ซ้ำติดกัน
    - เก็บชื่อไทย + อังกฤษของบริษัทไว้
    """

    if not name:
        return ""

    name = str(name).replace("\n", " ")
    name = re.sub(r"\s+", " ", name).strip(" ,;:-")

    if not name:
        return ""

    # 1) ลบ Logo / Prefix ก่อน
    name = remove_supplier_logo_prefix(name)

    # 2) OCR บางใบอ่านคำว่า "บริษัท" ด้านหน้าหายเป็น "ษัท"
    # แก้เฉพาะต้น SupplierName เพื่อลดผลกระทบกับข้อความส่วนอื่น
    name = re.sub(r"^\s*ษัท\s+", "บริษัท ", name)

    stop_patterns = [
        # Certification / Quality
        r"\bISO\s*\d{3,6}(?:\s*:\s*\d{4})?\b",
        r"\bIATF\s*\d+\b",
        r"\bCERTIFIED\b",
        r"\bCERTIFICATION\b",
        r"\bQUALITY\s+(?:SYSTEM|MANAGEMENT)\b",
        r"\bQUALITY\s+STANDARD\b",
        r"ได้รับมาตรฐาน",
        r"มาตรฐานระบบคุณภาพ",
        r"ระบบบริหารคุณภาพ",
        r"ระบบคุณภาพ",
        r"การรับรองมาตรฐาน",

        # Website / Internet
        r"\bhttps?://",
        r"\bwww\.",
        r"\bwebsite\b",

        # Contact
        r"\bTEL(?:EPHONE)?\.?\s*[:：]",
        r"\bPHONE\s*[:：]",
        r"\bFAX\.?\s*[:：]",
        r"\bE-?MAIL\s*[:：]",

        # Tax / document information
        r"\bTAX\s*ID\b",
        r"\bVAT\s*ID\b",
        r"เลขประจำตัวผู้เสียภาษี",
        r"\bTAX\s+INVOICE\b",
        r"\bINVOICE\b",

        # Address
        r"\bADDRESS\s*[:：]",
        r"ที่อยู่\s*[:：]",
    ]

    positions = []
    for pattern in stop_patterns:
        m = re.search(pattern, name, re.IGNORECASE)
        if m:
            positions.append(m.start())

    if positions:
        name = name[:min(positions)].strip(" ,;:-")

    name = re.sub(r"\s+", " ", name).strip(" ,;:-")

    if not name:
        return ""

    # 3) ลบ phrase ที่ซ้ำติดกันแบบ Generic
    words = name.split()
    result = []
    i = 0

    while i < len(words):
        duplicated = False
        max_block = min((len(words) - i) // 2, 20)

        for size in range(max_block, 1, -1):
            block1 = words[i:i + size]
            block2 = words[i + size:i + (size * 2)]

            if " ".join(block1).casefold() == " ".join(block2).casefold():
                result.extend(block1)
                i += size * 2
                duplicated = True
                break

        if not duplicated:
            result.append(words[i])
            i += 1

    name = " ".join(result)
    name = re.sub(r"\s+", " ", name).strip(" ,;:-")

    # Final guard หลัง merge / cleanup
    name = remove_supplier_logo_prefix(name)

    # เก็บเฉพาะชื่อบริษัทที่ลงท้ายด้วยรูปแบบนิติบุคคลที่สมบูรณ์
    # และตัด English fragment/ชื่อซ้ำที่ไม่สมบูรณ์ออก
    name = clean_supplier_legal_ending(name)

    return name


def extract_supplier_name_from_pages(invoices, ocr_cache=None):
    lines = ocr_cache["lines"] if ocr_cache is not None else get_all_lines(invoices)

    ssk_candidates = []

    for line in lines[:30]:
        text = line.strip()

        if "SSK" in text.upper() and re.search(r"PLASTIC|พลาสติก", text, re.IGNORECASE):
            ssk_candidates.append(clean_supplier_name(text))

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
                thai_name = clean_supplier_name(text)
                break

        for i, line in enumerate(lines):
            upper = line.upper()

            if (
                "UNION NIFCO" in upper
                and "DIGITALLY" not in upper
                and "THIS DOCUMENT" not in upper
                and "RECEIVED" not in upper
            ):
                eng_name = clean_supplier_name(line.strip())

                if i + 1 < len(lines):
                    nxt = lines[i + 1].strip().upper()
                    if nxt in ("LTD.", "LTD", "LIMITED"):
                        eng_name = clean_supplier_name(
                            eng_name + " " + lines[i + 1].strip()
                        )
                break

        if thai_name and eng_name:
            return clean_supplier_name(
                f"{thai_name} {eng_name}"
            ).replace(";", ",")

        if thai_name:
            return thai_name.replace(";", ",")

        if eng_name:
            return eng_name.replace(";", ",")

    # Supplier ปกติ
    search_limit = 20

    for i, line in enumerate(lines[:search_limit]):
        text = line.strip()
        upper = text.upper()

        if any(k.upper() in upper for k in SUPPLIER_SKIP_KEYWORDS):
            continue

        for pattern in SUPPLIER_NAME_PATTERNS:
            if re.search(pattern, upper, re.IGNORECASE):

                supplier = clean_supplier_name(text)

                # ไม่เติมข้อความจากบรรทัดก่อนหน้าเข้ามาใน SupplierName
                # เพราะกฎใหม่กำหนดว่า ถ้ามีข้อความก่อนคำว่า "บริษัท"
                # ให้ตัดทิ้งทั้งหมด เพื่อกัน Logo/Prefix ทุกชนิดโดยอัตโนมัติ

                # บรรทัดถัดไป:
                # ต่อเฉพาะเมื่อดูเหมือนเป็นชื่อบริษัทอีกภาษาและไม่ใช่ข้อความซ้ำ
                if i + 1 < len(lines):

                    nxt = lines[i + 1].strip()
                    cleaned_next = clean_supplier_name(nxt)

                    current_has_thai = bool(
                        re.search(r"[\u0E00-\u0E7F]", supplier)
                    )
                    next_has_thai = bool(
                        re.search(r"[\u0E00-\u0E7F]", cleaned_next)
                    )
                    current_has_english = bool(
                        re.search(r"[A-Za-z]", supplier)
                    )
                    next_has_english = bool(
                        re.search(r"[A-Za-z]", cleaned_next)
                    )

                    next_is_company = bool(re.search(
                        r"(บริษัท|จำกัด|CO\.?\s*,?\s*LTD\.?|COMPANY\s+LIMITED|PUBLIC\s+COMPANY)",
                        cleaned_next,
                        re.IGNORECASE,
                    ))

                    different_language = (
                        (current_has_english and next_has_thai)
                        or (current_has_thai and next_has_english)
                    )

                    not_duplicate = (
                        cleaned_next
                        and cleaned_next.casefold() not in supplier.casefold()
                        and supplier.casefold() not in cleaned_next.casefold()
                    )

                    if next_is_company and different_language and not_duplicate:
                        supplier += " " + cleaned_next

                return clean_supplier_name(supplier).replace(";", ",")

    return ""


# def extract_supplier_name_from_pages(invoices, ocr_cache):
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

def extract_vendor_branch(invoices, layout_result=None, ocr_cache=None):
    """
    SINGLE-MODEL VERSION
    ใช้ผลจาก prebuilt-invoice เพียงตัวเดียว ทั้ง OCR text, line position/polygon
    และ selection marks (ถ้ามีใน analyze result) โดยไม่ต้องยิง prebuilt-layout เพิ่ม

    layout_result ถูกเก็บไว้ใน signature เพื่อไม่ให้ส่วนอื่นที่อาจเรียก function นี้พัง
    แต่ใน flow ใหม่ไม่จำเป็นต้องส่งค่าเข้ามา
    """
    lines = ocr_cache["lines"] if ocr_cache is not None else get_all_lines(invoices)
    full_text = ocr_cache["full_text"] if ocr_cache is not None else "\n".join(lines)

    # ใช้ pages จาก prebuilt-invoice เป็นแหล่งข้อมูลตำแหน่งหลัก
    pages = getattr(invoices, "pages", None) or []

    # SPECIAL CASE: T.KRUNGTHAI INDUSTRIES
    # บริษัทนี้ในหัวเอกสารมีรายการสาขา 00001/00002/00003 หลายบรรทัด
    # เลือกช่อง "สาขาที่" ด้านขวาของหน้า หรือบรรทัดเดียวกับ เลขที่/No.
    if (
        "T.KRUNGTHAI INDUSTRIES" in full_text.upper()
        or "ที.กรุงไทยอุตสาหกรรม" in full_text
        or "กรุงไทยอุตสาหกรรม" in full_text
    ):
        branch_value = ""

        # 0.1) ใช้ line polygon จาก prebuilt-invoice โดยตรง
        for page in pages:
            page_width = getattr(page, "width", 0) or 0

            for line in (getattr(page, "lines", None) or []):
                line_text = line.content.strip() if getattr(line, "content", None) else ""
                if not line_text:
                    continue

                m = re.search(r"สาขา\s*(?:ที่)?\s*[:：]?\s*(\d{1,10})", line_text, re.IGNORECASE)
                if not m:
                    continue

                x_min = 0
                try:
                    polygon = getattr(line, "polygon", None) or []
                    # SDK รุ่นใหม่ polygon มักเป็น list ของ Point(x,y)
                    if polygon and hasattr(polygon[0], "x"):
                        xs = [p.x for p in polygon]
                    else:
                        # รองรับรูปแบบเลข flat list เดิม
                        xs = polygon[0::2] if polygon else []
                    x_min = min(xs) if xs else 0
                except Exception:
                    x_min = 0

                if (page_width and x_min >= page_width * 0.55) or re.search(r"เลขที่|No\.?", line_text, re.IGNORECASE):
                    branch_value = m.group(1)

            if branch_value:
                return branch_value.zfill(5)

        # 0.2) fallback จาก OCR line: เลือกบรรทัดที่มีทั้ง เลขที่/No. และ สาขา
        for line in lines:
            if re.search(r"เลขที่|No\.?", line, re.IGNORECASE) and re.search(r"สาขา", line):
                m = re.search(r"สาขา\s*(?:ที่)?\s*[:：]?\s*(\d{1,10})", line, re.IGNORECASE)
                if m:
                    return m.group(1).zfill(5)

        # 0.3) fallback สุดท้าย เลือก occurrence ท้ายสุด
        matches = re.findall(r"สาขา\s*(?:ที่)?\s*[:：]?\s*(\d{1,10})", full_text, re.IGNORECASE)
        if matches:
            return matches[-1].zfill(5)

    # 1) Selection mark จาก prebuilt-invoice ถ้ามี
    # ไม่ยิง layout เพิ่ม แต่ยังใช้ checkbox ได้เมื่อ model ส่ง selection_marks มา
    for page in pages:
        selected_marks = []
        for mark in (getattr(page, "selection_marks", None) or []):
            state = getattr(mark, "state", None)
            state_name = getattr(state, "name", str(state or ""))
            if str(state_name).upper().endswith("SELECTED"):
                selected_marks.append(mark)

        for mark in selected_marks:
            try:
                mark_polygon = getattr(mark, "polygon", None) or []
                if mark_polygon and hasattr(mark_polygon[0], "y"):
                    mark_y = min(p.y for p in mark_polygon)
                else:
                    mark_y = mark_polygon[1] if len(mark_polygon) > 1 else 0
            except Exception:
                mark_y = 0

            same_row_texts = []

            for line in (getattr(page, "lines", None) or []):
                line_text = line.content.strip() if getattr(line, "content", None) else ""
                if not line_text:
                    continue

                try:
                    line_polygon = getattr(line, "polygon", None) or []
                    if line_polygon and hasattr(line_polygon[0], "y"):
                        line_y = min(p.y for p in line_polygon)
                    else:
                        line_y = line_polygon[1] if len(line_polygon) > 1 else 0
                except Exception:
                    line_y = 0

                if abs(mark_y - line_y) <= 0.08:
                    same_row_texts.append(line_text)

            same_row_text = " ".join(same_row_texts)

            if re.search(r"สำนักงานใหญ่|Head\s*Office|HeadOffice", same_row_text, re.IGNORECASE):
                return "00000"

            value = find_first_by_patterns(VENDOR_BRANCH_PATTERNS, same_row_text)
            if value:
                return value.zfill(5)

    # 2) ข้อความระบุสาขาที่ออกใบกำกับภาษีโดยตรง
    if re.search(
        r'(?:ออกโดย|สาขาที่ออกใบกำกับภาษี)\s*[:：]?\s*["“”]?\s*(สำนักงานใหญ่|Head\s*Office|HeadOffice)',
        full_text,
        re.IGNORECASE,
    ):
        return "00000"

    m = re.search(
        r'(?:ออกโดย|สาขาที่ออกใบกำกับภาษี)\s*[:：]?\s*["“”]?\s*สาขา(?:ที่|เลขที่)?\s*(\d{1,10})',
        full_text,
        re.IGNORECASE,
    )
    if m:
        return m.group(1).zfill(5)

    # 3) Vendor zone only, stop before Company/TKC
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


def extract_tax_id_from_pages(invoices, ocr_cache=None):
    full_text = ocr_cache["full_text"] if ocr_cache is not None else get_all_text(invoices)

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

def extract_vat_from_pages(invoices, ocr_cache=None):

    if ocr_cache is not None:
        full_text = ocr_cache["full_text"]
    else:
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

def clean_po_list(po_list):
    """Normalize, uppercase และตัดค่าซ้ำของ PO"""
    cleaned = []

    for po in po_list:
        po = str(po or "").upper().strip()

        # เอาคำว่า PO ที่เป็น prefix ออกเท่านั้น
        po = re.sub(r"^PO\s*", "", po, flags=re.IGNORECASE).strip()

        if po:
            cleaned.append(po)

    return ",".join(dict.fromkeys(cleaned))


def extract_po_from_text_and_tables(invoices, branch_code, extra_text="", ocr_cache=None):
    """
    ค้นหา PO โดยอิง Company Branch ที่ส่งมาตอน Run

    ลำดับ:
    1) หา PO ของ branch ปัจจุบันก่อน
    2) ถ้าไม่เจอ -> หา PO ของอีก branch
       ถ้าเจอ จะคืน po_branch_mismatch = True
    3) ถ้ายังไม่เจอ -> fallback ไปใช้ PO_PATTERNS เดิม
       เพื่อรองรับ PO รูปแบบทั่วไปที่ไม่อยู่ในกฎ branch

    return:
        (po_value, po_branch_mismatch)
    """

    branch_code = str(branch_code or "").strip().zfill(4)

    if ocr_cache is not None:
        full_text = ocr_cache["po_base_text"]
        if extra_text:
            full_text = full_text + ("\n" if full_text else "") + str(extra_text)
    else:
        texts = []

        # OCR Lines
        for page in invoices.pages:
            for line in page.lines:
                if line.content:
                    texts.append(line.content.strip())

        # OCR Tables
        if hasattr(invoices, "tables"):
            for table in invoices.tables:
                for cell in table.cells:
                    if cell.content:
                        texts.append(cell.content.strip())

        # PO ที่ Azure prebuilt-invoice ดึงมาแล้ว
        if extra_text:
            texts.append(str(extra_text))

        full_text = "\n".join(texts)

    # --------------------------------------------------
    # 1) หา PO ของ branch ปัจจุบันก่อน
    # --------------------------------------------------
    current_patterns = get_po_patterns_by_branch(branch_code)
    current_po_list = find_all_by_patterns(current_patterns, full_text)

    if current_po_list:
        po_value = clean_po_list(current_po_list)
        print(f"✅ PO matched branch {branch_code}: {po_value}")
        return po_value, False

    # --------------------------------------------------
    # 2) ไม่เจอ -> หา PO ของอีก branch
    # --------------------------------------------------
    other_branch = get_other_branch(branch_code)

    if other_branch:
        other_patterns = get_po_patterns_by_branch(other_branch)
        other_po_list = find_all_by_patterns(other_patterns, full_text)

        if other_po_list:
            po_value = clean_po_list(other_po_list)
            print(
                f"⚠️ PO found in other branch "
                f"(Input Branch={branch_code}, PO Branch={other_branch}): {po_value}"
            )
            return po_value, True

    # --------------------------------------------------
    # 3) Fallback PO Patterns เดิม
    #    เช่น PO140 หรือรูปแบบอื่นที่ไม่ได้ผูกกับ branch
    # --------------------------------------------------
    po_no_list = find_all_by_patterns(PO_PATTERNS, full_text)

    if po_no_list:
        po_value = clean_po_list(po_no_list)
        print(f"✅ PO found by generic pattern: {po_value}")
        return po_value, False

    return "", False


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

def extract_tax_invoice_no_from_pages(invoices, ocr_cache=None):
    """Fallback InvoiceNo จากผล prebuilt-invoice เพียง model เดียว"""
    if ocr_cache is not None:
        full_text = ocr_cache.get("full_text", "") or ""
    else:
        all_lines = []
        for page in (getattr(invoices, "pages", None) or []):
            for line in (getattr(page, "lines", None) or []):
                if getattr(line, "content", None):
                    all_lines.append(line.content.strip())
        full_text = " ".join(all_lines)

    full_text = re.sub(r"\s+", " ", full_text)

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

    # เดิมใช้ "OTH".join(value) ซึ่งจะแทรก OTH ระหว่างทุกตัวอักษร
    # รักษาเจตนาเดิมโดยเติม OTH เฉพาะเมื่อ NIFCO และยังไม่มี prefix OTH
    if value and ("NIFCO" in full_text.upper() or "นิฟโก้" in full_text):
        if not str(value).upper().startswith("OTH"):
            value = "OTH" + str(value)

    if value and not re.match(r"^(?:PO)?(?:410|140)\d{7}$", str(value), re.IGNORECASE):
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

def check_copy_document(full_text):
    """
    ตรวจว่าเอกสารเป็นสำเนาหรือไม่

    ตรวจคำว่า:
    - COPY
    - สำเนา

    return True = เป็นเอกสารสำเนา
    """

    if not full_text:
        return False

    text = str(full_text)

    has_copy = (
        "สำเนา" in text
        or re.search(r"\bCOPY\b", text, re.IGNORECASE)
    )

    has_original = (
        "ต้นฉบับ" in text
        or re.search(r"\bORIGINAL\b", text, re.IGNORECASE)
    )

    return bool(has_copy and not has_original)

# 📌 Main Convert
def extract_invoice_to_json(invoice, invoices, ocr_cache=None):

    supplier_name = get_field_value(invoice.fields.get("VendorAddressRecipient"))

    # ==========================================================
    # Address
    # ==========================================================
    # 1) อ่านจาก Azure VendorAddress ก่อน
    address = get_address_value(
        invoice.fields.get("VendorAddress")
    )

    # 2) ถ้า Azure ไม่มี Vendor Address ให้ fallback ไป OCR
    if not address:
        address = extract_vendor_address_from_pages(invoices, ocr_cache)

    # ==========================================================
    # Customer / Company Address
    # ==========================================================
    # prebuilt-invoice ใช้ CustomerAddress สำหรับที่อยู่ผู้ซื้อ/ลูกค้า
    # เก็บลง CompanyAddress ต่อไป เพื่อไม่กระทบ logic / Excel เดิม
    Company_address = get_address_value(
        invoice.fields.get("CustomerAddress")
    )

    # บาง template Azure อาจ map ที่อยู่ผู้ซื้อไป BillingAddress
    if not Company_address:
        Company_address = get_address_value(
            invoice.fields.get("BillingAddress")
        )

    # รองรับกรณี SDK/model บางเวอร์ชันคืน key เดิมที่โปรแกรมเคยใช้งาน
    if not Company_address:
        Company_address = get_address_value(
            invoice.fields.get("CompanyAddress")
        )

    # ถ้า Azure ไม่มี field ให้ fallback ไป OCR จาก Account To / Bill To
    if not Company_address:
        Company_address = extract_Company_address_from_pages(invoices, ocr_cache)

    # ==========================================================
    # Clean CompanyAddress
    # ตัดข้อความก่อนเลขที่ตั้ง เช่น
    # ชื่อลูกค้า 370... -> 370...
    # HEADOFFICE:370... -> 370...
    # ==========================================================
    Company_address = clean_company_address_start(Company_address)

    supplier_from_pages = extract_supplier_name_from_pages(invoices, ocr_cache)

    if supplier_from_pages:
        if "SSK" in supplier_from_pages.upper() and "SSK" not in supplier_name.upper():
            supplier_name = supplier_from_pages
        elif not supplier_name or len(supplier_from_pages) > len(supplier_name):
            supplier_name = supplier_from_pages
    
    # ทำความสะอาด Supplier Name แบบ Generic
    supplier_name = clean_supplier_name(supplier_name)

    # Fix SSK เดิม
    supplier_name = supplier_name.replace(
        "SSK Plastic Co.,Ltd. Plastic Co.,Ltd.",
        "SSK Plastic Co.,Ltd. บริษัท เอส.เอส.เค พลาสติก จำกัด"
    )

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
        supplier_name = extract_supplier_name_from_pages(invoices, ocr_cache)

    if supplier_name:
        supplier_name = supplier_name.replace(";", ",")

    # Clean อีกครั้งหลัง fallback/merge SupplierName เสร็จ
    supplier_name = clean_supplier_name(supplier_name)

    # Normalize SSK Supplier Name
    supplier_name = supplier_name.strip()

    if supplier_name == "SSK Plastic Co.,Ltd.":
        supplier_name = "SSK Plastic Co.,Ltd. บริษัท เอส.เอส.เค พลาสติก จำกัด"

    # Final SupplierName cleanup หลัง extraction/fallback/merge ทุกขั้นตอน
    supplier_name = clean_supplier_name(supplier_name)

    # ==========================================================
    # Invoice Date
    # ==========================================================
    # Priority:
    # 1) วันที่ที่อยู่ใกล้ label Invoice Date / วันที่-Date บนเอกสาร
    # 2) Azure InvoiceDate
    # 3) ServiceStartDate
    #
    # IMPORTANT:
    # - ห้ามใช้ DueDate เป็น InvoiceDate
    # - ถ้า Azure อ่าน InvoiceDate ไปตรงกับ DueDate แต่ OCR พบวันที่เอกสารจริง
    #   ให้ใช้วันที่จาก OCR label แทน
    #
    # ตัวอย่าง:
    #   วันที่/Date     = 20/06/2026
    #   DUE DATE       = 20/07/2026
    #   ผลที่ต้องได้   = 20/06/2026
    # ==========================================================

    invoice_date = normalize_invoice_date(
        get_field_value(invoice.fields.get("InvoiceDate"))
    )

    service_start = normalize_invoice_date(
        get_field_value(invoice.fields.get("ServiceStartDate"))
    )

    due_date = normalize_invoice_date(
        get_field_value(invoice.fields.get("DueDate"))
    )

    # อ่านวันที่จาก label บนหน้าเอกสารโดยตรง
    label_invoice_date = extract_invoice_date_near_date_label(
        invoices,
        ocr_cache
    )

    invoice_dt = parse_date_safe(invoice_date)
    service_dt = parse_date_safe(service_start)
    due_dt = parse_date_safe(due_date)
    label_dt = parse_date_safe(label_invoice_date)
    current_year = datetime.now().year

    final_invoice_date = invoice_date

    # ----------------------------------------------------------
    # 1) ถ้า OCR พบ Invoice Date ที่ label ชัดเจน ให้ใช้แก้กรณี Azure map ผิด
    # ----------------------------------------------------------
    if label_dt:

        # Azure ไม่มี InvoiceDate
        if not invoice_dt:
            final_invoice_date = label_dt.strftime("%d/%m/%Y")

        # Azure เอา DueDate มาใส่เป็น InvoiceDate
        elif (
            due_dt
            and invoice_dt.date() == due_dt.date()
            and label_dt.date() != due_dt.date()
        ):
            print(
                f"📅 InvoiceDate corrected from DueDate: "
                f"{invoice_date} -> {label_invoice_date}"
            )
            final_invoice_date = label_dt.strftime("%d/%m/%Y")

        # Azure อ่านปีผิด แต่วันที่ที่ label อยู่ในปีปัจจุบัน
        elif (
            invoice_dt.year != current_year
            and label_dt.year == current_year
        ):
            print(
                f"📅 InvoiceDate corrected from label: "
                f"{invoice_date} -> {label_invoice_date}"
            )
            final_invoice_date = label_dt.strftime("%d/%m/%Y")

    # ----------------------------------------------------------
    # 2) ถ้ายังไม่มี InvoiceDate จริง ๆ ใช้ ServiceStartDate เป็น fallback
    #    แต่ไม่ใช้ DueDate
    # ----------------------------------------------------------
    final_dt = parse_date_safe(final_invoice_date)

    if not final_dt and service_dt:
        final_invoice_date = service_dt.strftime("%d/%m/%Y")

    # Safety guard:
    # ถ้าค่าสุดท้ายยังเท่ากับ DueDate และมีวันที่จาก label ที่ต่างกัน
    # ให้ยืนยันใช้วันที่จาก label
    final_dt = parse_date_safe(final_invoice_date)

    if (
        final_dt
        and due_dt
        and label_dt
        and final_dt.date() == due_dt.date()
        and label_dt.date() != due_dt.date()
    ):
        print(
            f"📅 Safety Date Fix: "
            f"{final_invoice_date} -> {label_invoice_date}"
        )
        final_invoice_date = label_dt.strftime("%d/%m/%Y")

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

    return {
        "InvoiceDate": final_invoice_date,
        "PostingDate": final_invoice_date,
        "TaxInvoiceNo": tax_invoice_no_clean,
        "SupplierName": supplier_name,
        "Address": address,
        "CompanyAddress": Company_address,
        "CompanyName": "",
        "CompanyTaxID": "",
        "CompanyBranch": "",
        "Assignment": "",
        "VendorTaxId": vendor_tax_id,
        "VendorBranch": "",
        "TotalAmount": normalize_number(total_amount),
        "VATAmount": normalize_number(vat_amount),
        "AmountIncVat": normalize_number(amount_inc_vat),
        "PurchaseOrderNo": clean_po_from_field(purchase_order_no1),
        "TaxRemark": "",
        "Emessage": "",
    }


# ==========================================================
# Build Excel Row
# ==========================================================

def build_excel_row(invoice):

    return {


        "TaxInvoiceNo": invoice.get("TaxInvoiceNo", ""),
        "InvoiceDate": invoice.get("InvoiceDate", ""),
        "SupplierName": invoice.get("SupplierName", ""),
        "VendorBranch": invoice.get("VendorBranch", ""),  
        "VendorTaxId": invoice.get("VendorTaxId", ""), 
        "Address": invoice.get("Address", ""),
        "TotalAmount": invoice.get("TotalAmount", ""),
        "VATAmount": invoice.get("VATAmount", ""),
        "AmountIncVat": invoice.get("AmountIncVat", ""),
        "PurchaseOrderNo": invoice.get("PurchaseOrderNo", ""),
        "TaxRemark": invoice.get("TaxRemark", ""),
        "CustomerName": invoice.get("CompanyName", ""),
        "CustomerTaxID": invoice.get("CompanyTaxID", ""),
        "CustomerBranch": invoice.get("CompanyBranch", ""),
        "CustomerAddress": invoice.get("CompanyAddress", ""),
        "Assignment": invoice.get("Assignment", ""),
        "Emessage": invoice.get("Emessage", "")
    }

def merge_invoice_row(existing, new):

    for field in [
        "InvoiceDate",
        "PostingDate",
        "TaxInvoiceNo",
        "SupplierName",
        "Assignment",
        "VendorTaxId",
        "VendorBranch",
        "Address",
        "CompanyAddress",
        "TaxRemark",
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

        # ==================================================
        # SINGLE MODEL MODE
        # ยิง Azure เพียง prebuilt-invoice ครั้งเดียวต่อหน้า
        # ใช้ invoices.pages / invoices.tables / OCR cache แทน prebuilt-layout
        # ==================================================
        ocr_cache = build_ocr_cache(invoices)

        raw_page_text = ocr_cache["full_text_lower"]

        if "good receipt" in raw_page_text or "goods receipt" in raw_page_text:
            print("⏭️ พบคำว่า 'Good Receipt' → ข้ามหน้านี้ทันที")
            continue

        # OCR fallback: ให้วันที่ใกล้ label Invoice Date มาก่อน
        # และค่อยใช้ oldest date เมื่อหา label ไม่เจอ
        invoice_date_ocr = extract_invoice_date_near_date_label(invoices, ocr_cache)
        if not invoice_date_ocr:
            invoice_date_ocr = extract_oldest_date_from_text(invoices, ocr_cache)

        for idx, invoice in enumerate(invoices.documents):

            invoice_data = extract_invoice_to_json(invoice, invoices, ocr_cache)

            lines = ocr_cache["lines"]
            full_text = ocr_cache["full_text_upper"]

            # ==================================================
            # Check COPY / สำเนา
            # ==================================================

            if check_copy_document(ocr_cache["full_text"]):
                invoice_data["Emessage"] = append_msg(
                    invoice_data.get("Emessage", ""),
                    "เอกสารใบนี้เป็นสำเนา"
                )

                print("⚠️ เอกสารใบนี้เป็นสำเนา")

            # Address fallback อีกชั้นก่อนทำขั้นตอนต่อไป
            if not invoice_data.get("Address"):
                invoice_data["Address"] = extract_vendor_address_from_pages(invoices, ocr_cache)

            if not invoice_data.get("CompanyAddress"):
                invoice_data["CompanyAddress"] = clean_company_address_start(
                    extract_Company_address_from_pages(invoices, ocr_cache)
                )

            # ==================================================
            # Company Fix จาก CompanyAddress
            # ถ้าที่อยู่มีเลข 370 → กำหนด Company เป็น XXX1
            # ==================================================
            Company_address = str(
                invoice_data.get("CompanyAddress", "") or ""
            ).strip()

            # ==================================================
            # Company Fix + Validate Branch จากค่าที่ส่งมาตอน Run
            # ==================================================

            Company_address = str(
                invoice_data.get("CompanyAddress", "") or ""
            ).strip()

            invoice_data["CompanyName"] = "THAI KOITO COMPANY LIMITED"
            invoice_data["CompanyTaxID"] = "0105529030059"

            Company_branch, branch_correct, expected_branch = validate_Company_branch(
                Company_address,
                branch_email
            )

            invoice_data["CompanyBranch"] = Company_branch

            if branch_correct:
                print(
                    f"✅ Company Branch ถูกต้อง "
                    f"(Input={Company_branch}, Expected={expected_branch})"
                )
            else:
                print(
                    f"❌ ที่อยู่ไม่สอดคลองกัน "
                    f"(Input={Company_branch}, Expected={expected_branch})"
                )

                invoice_data["Emessage"] = append_msg(
                    invoice_data.get("Emessage", ""),
                    "Company address does not match"
                )

            if not invoice_data.get("InvoiceDate") and invoice_date_ocr:
                invoice_data["InvoiceDate"] = invoice_date_ocr
                invoice_data["PostingDate"] = invoice_date_ocr

            if normalize_number(invoice_data.get("VATAmount", 0)) == 0:
                
                vat_fallback = extract_vat_from_pages(invoices, ocr_cache)
                if vat_fallback:
                    invoice_data["VATAmount"] = vat_fallback

            if not invoice_data.get("VendorTaxId"):
                invoice_data["VendorTaxId"] = extract_tax_id_from_pages(invoices, ocr_cache)
            
            # ==================================================
            # TPM Final InvoiceNo Fix
            # ==================================================
            if invoice_data.get("VendorTaxId") == "0115539007424":

                tax_invoice_no = str(
                    invoice_data.get("TaxInvoiceNo", "") or ""
                ).strip()

                if tax_invoice_no:

                    fixed_invoice_no = re.sub(
                        r'(?i)\bNo\.?\s*8(?=\d{4,8}/\d{2,4}\b)',
                        "No.S",
                        tax_invoice_no
                    )

                    if fixed_invoice_no != tax_invoice_no:

                        print(
                            f"✅ TPM Final InvoiceNo Fix: "
                            f"{tax_invoice_no} -> {fixed_invoice_no}"
                        )

                        invoice_data["TaxInvoiceNo"] = fixed_invoice_no


            # ==================================================
            # VENDOR / TPM DETECTION (SINGLE MODEL)
            # ==================================================
            is_union_plastic = "UNION PLASTIC" in full_text
            vendor_tax_id = clean_tax_id(
                invoice_data.get("VendorTaxId", "")
            )

            vendor_tax_id = clean_tax_id(
                invoice_data.get("VendorTaxId", "")
            )

            # ตรวจ Tax ID จาก OCR แต่ละบรรทัดด้วย
            tpm_tax_id_found_in_ocr = any(
                "0115539007424" in clean_tax_id(line)
                for line in ocr_cache["lines"]
            )

            is_tpm = (
                vendor_tax_id == "0115539007424"
                or tpm_tax_id_found_in_ocr
                or "THAI PRESS AND MACHINERY" in full_text.upper()
                or "ไทยเพรส แอนด์ แมชชีนเนอรี่" in full_text
            )

            print(
                f"🔎 TPM DETECT: "
                f"VendorTaxId={vendor_tax_id} | "
                f"TaxIdInOCR={tpm_tax_id_found_in_ocr} | "
                f"is_tpm={is_tpm}"
            )

            # ==================================================
            # VendorBranch - SINGLE MODEL
            # ใช้ OCR + line position/polygon + selection_marks (ถ้ามี)
            # จาก prebuilt-invoice โดยตรง
            # UNION PLASTIC จะใช้ OCR/vendor-zone fallback เช่นเดิม โดยไม่ต้องมี layout
            # ==================================================
            invoice_data["VendorBranch"] = extract_vendor_branch(
                invoices,
                layout_result=None,
                ocr_cache=ocr_cache
            )

            # ==================================================
            # TaxInvoiceNo fallback - SINGLE MODEL
            # ใช้ OCR text ที่ได้จาก prebuilt-invoice แทน prebuilt-layout
            # ==================================================
            if not invoice_data.get("TaxInvoiceNo"):
                fallback_no = extract_tax_invoice_no_from_pages(invoices, ocr_cache)

                if fallback_no:
                    invoice_data["TaxInvoiceNo"] = fallback_no
                    print(f"✅ Fallback TaxInvoiceNo from prebuilt-invoice OCR: {fallback_no}")

            if not invoice_data.get("TaxInvoiceNo"):
                fallback_no = find_invoice_no_from_words(invoices)
                if fallback_no:
                    invoice_data["TaxInvoiceNo"] = fallback_no
                    print(f"✅ Fallback InvoiceId found from OCR: {fallback_no}")

                    if is_tpm:

                        tax_invoice_no = str(
                            invoice_data.get("TaxInvoiceNo", "") or ""
                        ).strip()

                        fixed_invoice_no = re.sub(
                            r'(?i)No\.?\s*8(?=\d{4,8}/\d{2,4}\b)',
                            "No.S",
                            tax_invoice_no
                        )

                        invoice_data["TaxInvoiceNo"] = fixed_invoice_no

                        print(
                            f"✅ TPM FINAL: "
                            f"{tax_invoice_no} -> {fixed_invoice_no}"
                        )

                else:
                    print("⚠️ InvoiceId not found (even from OCR)")

            # ==================================================
            # TPM FINAL TaxInvoiceNo Fix
            # No.8xxxxx/xx -> No.Sxxxxx/xx
            # ==================================================
            if is_tpm:

                tax_invoice_no = str(
                    invoice_data.get("TaxInvoiceNo", "") or ""
                ).strip()

                print(
                    f"🔎 TPM CHECK: "
                    f"VendorTaxId={invoice_data.get('VendorTaxId')} | "
                    f"TaxInvoiceNo={tax_invoice_no}"
                )

                if tax_invoice_no:

                    fixed_invoice_no = re.sub(
                        r'(?i)No\.?\s*8(?=\d{4,8}/\d{2,4}\b)',
                        "No.S",
                        tax_invoice_no
                    )

                    if fixed_invoice_no != tax_invoice_no:
                        print(
                            f"✅ TPM InvoiceNo Fix: "
                            f"{tax_invoice_no} -> {fixed_invoice_no}"
                        )

                        invoice_data["TaxInvoiceNo"] = fixed_invoice_no

            # ==================================================
            # PO Validation by Company Branch
            # 1) หา PO ของ branch ที่ส่งมาตอน Run ก่อน
            # 2) ถ้าไม่เจอ -> หาอีก branch
            # 3) ถ้าเจออีก branch -> เก็บ PO และเพิ่ม Emessage
            # ==================================================
            po_from_ocr, po_branch_mismatch = extract_po_from_text_and_tables(
                invoices,
                branch_email,
                invoice_data.get("PurchaseOrderNo", ""),
                ocr_cache
            )

            if po_from_ocr:
                invoice_data["PurchaseOrderNo"] = po_from_ocr

            if po_branch_mismatch:
                invoice_data["Emessage"] = append_msg(
                    invoice_data.get("Emessage", ""),
                    "PO does not match Company branch"
                )

            tax_invoice_no = (invoice_data.get("TaxInvoiceNo") or "").strip()
            if re.match(r"^(?:PO)?(?:410|140)\d{7}$", tax_invoice_no, re.IGNORECASE):
                invoice_data["TaxInvoiceNo"] = ""

            # if not invoice_data.get("InvoiceDate") and invoice_date_ocr:
            #     invoice_data["InvoiceDate"] = invoice_date_ocr
            #     invoice_data["PostingDate"] = invoice_date_ocr

            #Custom Nifco
            if "NIFCO" in full_text or "นิฟโก้" in full_text:
                invoice_data["TaxInvoiceNo"] = "OTH" + str(invoice_data.get("TaxInvoiceNo", ""))

            taxRemark = extract_tax_remark(invoices, ocr_cache)
            invoice_data["TaxRemark"] = taxRemark

            invoice_data["InvoiceDate"] = normalize_invoice_date(invoice_data.get("InvoiceDate"))
            invoice_data["PostingDate"] = normalize_invoice_date(invoice_data.get("PostingDate"))
            invoice_data["Assignment"] = os.path.basename(input_pdf)

            # ==========================================================
            # FINAL Vendor-specific TaxInvoiceNo normalization
            # Run here LAST so no fallback can overwrite the corrected value
            # ==========================================================
            normalize_vendor_specific_invoice_no(invoice_data, full_text)

            print(
                f"📌 FINAL BEFORE SAVE: "
                f"VendorTaxId={invoice_data.get('VendorTaxId')} | "
                f"TaxInvoiceNo={invoice_data.get('TaxInvoiceNo')}"
            )

            normalize_amounts(invoice_data)
            add_or_merge_row(all_data, invoice_data)
        
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

# ==========================================================
# Excel Columns
# ==========================================================

EXCEL_COLUMNS = [
    "TaxInvoiceNo",
    "InvoiceDate",
    "SupplierName",
    "Address",
    "Assignment",
    "VendorTaxId",
    "VendorBranch",
    "TotalAmount",
    "VATAmount",
    "AmountIncVat",
    "CustomerName",
    "CustomerTaxID",
    "CustomerBranch",
    "CustomerAddress",
    "PurchaseOrderNo",
    "TaxRemark",
    "Emessage",
]
REQUIRED_FIELDS = [
    "TaxInvoiceNo",
    "InvoiceDate",
    "TotalAmount",
    "VATAmount",
    "SupplierName",
    "VendorTaxId",
    "VendorBranch",
    "PurchaseOrderNo",
]

for row in all_data:

    errors = []

    for field in REQUIRED_FIELDS:

        value = row.get(field)

        if value is None or str(value).strip() == "":
            errors.append(f"{field} is empty")

    try:
        total_amount = normalize_number(row.get("TotalAmount", 0))
        vat_amount = normalize_number(row.get("VATAmount", 0))

        expected_vat = round(total_amount * 0.07, 2)

        if abs(vat_amount - expected_vat) > 0.01:
            errors.append(
                "กรุณาเช็ค Amount ทั้ง 3 ช่อง"
            )

    except Exception:
        errors.append("VATAmount format invalid")

    row["Emessage"] = append_msg(
        row.get("Emessage", ""),
        " | ".join(errors)
    )

excel_rows = [build_excel_row(row) for row in all_data]

df = pd.DataFrame(excel_rows)
df = df.reindex(columns=EXCEL_COLUMNS)

df.to_excel(output_excel, index=False)

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
        col_number = column_map.get(field_name)
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
        "Company address does not match": [
            "CompanyAddress",
            "CompanyBranch",
        ],
        "PO does not match Company branch": [
            "PurchaseOrderNo",
            "CompanyBranch",
        ],
        "Duplicate Amount Found": [
            "TotalAmount",
            "VATAmount",
            "AmountIncVat",
        ],
        "Amount format invalid": [
            "TotalAmount",
            "VATAmount",
            "AmountIncVat",
        ],
        "VATAmount format invalid": [
            "TotalAmount",
            "VATAmount",
            "AmountIncVat",
        ],
        "กรุณาเช็ค Amount ทั้ง 3 ช่อง": [
            "TotalAmount",
            "VATAmount",
            "AmountIncVat",
        ],
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
            if error_keyword.casefold() in emessage_lower:
                for field in fields:
                    mark_cell(row_number, field)

    wb.save(excel_path)
    print("🎨 Highlight Error Cells completed")


highlight_excel_errors(output_excel)

print(f"\n✅ Done. All invoices saved to {output_excel}")