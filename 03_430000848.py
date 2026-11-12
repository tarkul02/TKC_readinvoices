import os
import sys
import shutil
import subprocess
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
from po_patterns import load_po_patterns, is_po_value

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

if len(sys.argv) < 2:
    print("❌ Usage: python 03_430000848.py <branch_email>")
    sys.exit(2)

# สร้างวันที่และเวลาอัตโนมัติ
today_str = datetime.now().strftime("%Y%m%d_%H%M%S")

# รับเฉพาะ Branch จาก Parameter
branch_email = sys.argv[1].strip()
if branch_email not in ("00000", "00001"):
    print("❌ Company Branch: ใช้ 00000 (บางพลี) หรือ 00001 (ปราจีน) เท่านั้น")
    sys.exit(2)

print(f"📅 OCR Folder: {today_str}")
print(f"🏢 Branch: {branch_email}")


# ====================================================
# 📁 PATH
# ====================================================

paths = get_paths(today_str, branch_email, "Process3")

# Existing OCR paths (keep these as the source of truth when paths overlap)
input_folder = paths["input_folder"]
temp_folder = paths["temp_folder"]
output_excel = paths["output_excel"]
dest_folder = paths["dest_folder"]
pdf_folder = os.path.join(dest_folder, "PDF")
REPORT_TXT_FOLDER = paths["report_txt_folder"]

# Matching_TextFile configuration
MATCHING_TEXT_FOLDER = paths["input_folder"]
MATCHING_MASTER_PATH = paths["master_supplier_path"]
MATCHING_OUTPUT_ROOT = os.path.join(
    MATCHING_TEXT_FOLDER,
    "OutputExcel"
)
MATCHING_TIMESTAMP = today_str
MATCHING_OUTPUT_FOLDER = dest_folder

# Function Matching_TextFile.py will use these paths to read the summary invoice report and master supplier file.
MATCHING_SCRIPT = paths["matching_script_path"]

def _report_clean(value):
    return str(value or "").strip().lstrip("\ufeff")


# def convert_report_ht_to_excel(report_folder):
#     """แปลง H/T เป็น Excel หากมี Report; หากไม่มีให้ข้ามและทำงานต่อ."""
#     if not os.path.isdir(report_folder):
#         print(
#             f"⚠️ ไม่พบโฟลเดอร์ Summary Invoice Report: {report_folder}"
#         )
#         print("⏭️ ข้ามขั้นตอน Summary Invoice Report และทำ OCR ต่อ")
#         return None

#     # 2. ตรวจสอบว่ามีไฟล์ TXT หรือไม่
#     txt_files = [
#         f for f in os.listdir(report_folder)
#         if f.lower().endswith(".txt")
#     ]

#     if not txt_files:
#         print("⚠️ ไม่พบไฟล์ Report TXT — ข้ามขั้นตอน Report")
#         return None

#     # 3. โค้ดเดิมสำหรับอ่านและแปลงไฟล์ H/T
#     # ให้วางโค้ดเดิมต่อจากตรงนี้

#     candidates = []
#     for filename in os.listdir(report_folder):
#         full_path = os.path.join(report_folder, filename)
#         if not os.path.isfile(full_path) or not filename.lower().endswith(".txt"):
#             continue
#         normalized_name = os.path.splitext(filename)[0].casefold().replace(" ", "_")
#         if normalized_name.startswith("summary") or (
#             "invoice" in normalized_name and "report" in normalized_name
#         ):
#             candidates.append(full_path)

#     if not candidates:
#         raise RuntimeError(
#             f"ไม่พบไฟล์ Summary Invoice Report .txt ใน {report_folder}; หยุดก่อนเริ่ม OCR"
#         )
#     if len(candidates) != 1:
#         names = ", ".join(sorted(os.path.basename(p) for p in candidates))
#         raise RuntimeError(
#             f"พบไฟล์ Summary Invoice Report มากกว่า 1 ไฟล์: {names}; หยุดก่อนเริ่ม OCR"
#         )

#     txt_path = candidates[0]
#     headers, totals, errors, order = {}, {}, [], []
#     try:
#         with open(txt_path, "r", encoding="utf-8-sig", errors="replace") as report_file:
#             lines = report_file.read().splitlines()
#     except OSError as exc:
#         raise RuntimeError(f"อ่านไฟล์ TXT ไม่สำเร็จ: {exc}; หยุดก่อนเริ่ม OCR") from exc

#     for line_no, line in enumerate(lines, 1):
#         if not line.strip():
#             continue
#         fields = [_report_clean(item) for item in line.split("|")]
#         record_type = fields[0][:1].upper() if fields and fields[0] else ""
#         if record_type not in ("H", "T"):
#             # L (line item) และ record ประเภทอื่นไม่ต้องส่งออก
#             continue

#         record_id = fields[0][1:].strip()
#         if not record_id or not record_id.isdigit():
#             errors.append(f"บรรทัด {line_no}: Record ID ของ {record_type} ไม่ถูกต้อง")
#             continue

#         if record_type == "H":
#             if record_id in headers:
#                 errors.append(f"บรรทัด {line_no}: พบ H Record ID ซ้ำ {record_id}")
#                 continue
#             headers[record_id] = fields
#             order.append(record_id)
#             if len(fields) < 20:
#                 errors.append(
#                     f"บรรทัด {line_no}: H Record ID {record_id} มี {len(fields)} fields; ต้องมีอย่างน้อย 20"
#                 )
#         else:
#             if record_id in totals:
#                 errors.append(f"บรรทัด {line_no}: พบ T Record ID ซ้ำ {record_id}")
#                 continue
#             totals[record_id] = fields
#             if len(fields) < 4:
#                 errors.append(
#                     f"บรรทัด {line_no}: T Record ID {record_id} มี {len(fields)} fields; ต้องมีอย่างน้อย 4"
#                 )

#     if not headers:
#         errors.append("ไม่พบ H record ในไฟล์รายงาน")
#     for record_id in headers.keys() - totals.keys():
#         errors.append(f"ไม่พบ T record ที่จับคู่กับ H Record ID {record_id}")
#     for record_id in totals.keys() - headers.keys():
#         errors.append(f"T Record ID {record_id} ไม่มี H record ที่จับคู่กัน")

#     if errors:
#         detail = "\n - ".join(errors[:30])
#         if len(errors) > 30:
#             detail += f"\n - และมีปัญหาเพิ่มเติมอีก {len(errors) - 30} รายการ"
#         raise RuntimeError("ตรวจสอบ Summary Invoice Report ไม่ผ่าน; หยุดก่อนเริ่ม OCR:\n - " + detail)

#     rows = []
#     for record_id in order:
#         h, t = headers[record_id], totals[record_id]
#         rows.append({
#             "ReportRecordID": record_id,
#             "CompanyCode": h[1], "DocumentDate": h[2], "PostingDate": h[3],
#             "DocumentType": h[4], "InvoiceNo": h[5], "ReferenceCode": h[6],
#             "Currency": h[7], "HeaderField09": h[8], "HeaderField10": h[9],
#             "HeaderField11": h[10], "HeaderDate": h[11],
#             "AmountIncVat_H": h[12], "AmountIncVat_H_Repeat": h[13],
#             "TaxCode": h[14], "TotalAmount_T": t[1], "VATAmount_T": t[3],
#             "HeaderField18": h[17], "HeaderField19": h[18],
#             "SupplierName": h[19], "TotalAmount_T_Repeat": t[2],
#             "VATAmount_T_Repeat": t[3],
#         })

#     # เก็บทั้ง TXT และ Excel ไว้ในโฟลเดอร์ Report เดียวกัน
#     try:
#         os.makedirs(REPORT_ARCHIVE_FOLDER, exist_ok=True)
#     except OSError as exc:
#         raise RuntimeError(
#             f"สร้างโฟลเดอร์ Report ไม่สำเร็จ: {exc}; หยุดก่อนเริ่ม OCR"
#         ) from exc

#     report_basename = os.path.splitext(os.path.basename(txt_path))[0]
#     output_path = os.path.join(REPORT_ARCHIVE_FOLDER, report_basename + ".xlsx")
#     archived_txt_path = os.path.join(REPORT_ARCHIVE_FOLDER, os.path.basename(txt_path))
#     if os.path.exists(output_path):
#         raise RuntimeError(
#             f"ไฟล์ Excel ปลายทางมีอยู่แล้ว จึงไม่เขียนทับ: {output_path}; หยุดก่อนเริ่ม OCR"
#         )
#     if os.path.exists(archived_txt_path):
#         raise RuntimeError(
#             f"ไฟล์ TXT ปลายทางมีอยู่แล้ว จึงไม่เขียนทับ: {archived_txt_path}; หยุดก่อนเริ่ม OCR"
#         )

#     try:
#         with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
#             df = pd.DataFrame(rows, columns=REPORT_HT_COLUMNS)
#             df.to_excel(writer, index=False, sheet_name="Report_HT")
#             worksheet = writer.sheets["Report_HT"]
#             worksheet.freeze_panes = "A2"
#             worksheet.auto_filter.ref = worksheet.dimensions
#             for column_cells in worksheet.columns:
#                 max_width = max(
#                     (len(str(cell.value)) if cell.value is not None else 0)
#                     for cell in column_cells
#                 )
#                 worksheet.column_dimensions[column_cells[0].column_letter].width = min(
#                     max(max_width + 2, 12), 40
#                 )
#     except Exception as exc:
#         raise RuntimeError(
#             f"สร้าง Excel จาก {os.path.basename(txt_path)} ไม่สำเร็จ: {exc}; หยุดก่อนเริ่ม OCR"
#         ) from exc

#     print("\n================ Summary Invoice Report ================")
#     print(f"TXT ต้นฉบับ: {txt_path}")
#     print(f"Excel ที่สร้าง: {output_path}")
#     print(f"จำนวน Invoice (H/T): {len(rows)}")

#     # ย้ายไฟล์ TXT ต้นฉบับไปไว้คู่กับ Excel ในโฟลเดอร์ Report
#     # เนื้อหา TXT ไม่ถูกแก้ไข และตรวจสอบชื่อซ้ำก่อนย้ายเพื่อป้องกันการเขียนทับ
#     try:
#         shutil.move(txt_path, archived_txt_path)
#         print(f"ย้าย TXT ไปเก็บแล้ว: {archived_txt_path}")
#         print(f"เก็บ Excel ไว้ที่: {output_path}")
#     except Exception as exc:
#         raise RuntimeError(
#             f"สร้าง/ย้ายไฟล์ TXT ไปยังโฟลเดอร์ Report ไม่สำเร็จ: {exc}; หยุดก่อนเริ่ม OCR"
#         ) from exc

#     print("========================================================\n")
#     return output_path


# ต้องแปลงและตรวจสอบรายงานให้สำเร็จก่อนสร้าง Azure client/เริ่ม OCR
# try:
#     report_excel_path = convert_report_ht_to_excel(REPORT_TXT_FOLDER)
# except Exception as report_error:
#     print(f"❌ Summary Invoice Report Error: {report_error}")
#     sys.exit(1)

# ====================================================
# 🔐 Azure
# ====================================================

endpoint = AZURE_ENDPOINT
key = AZURE_KEY
os.makedirs(temp_folder, exist_ok=True)
os.makedirs(os.path.dirname(output_excel), exist_ok=True)
os.makedirs(dest_folder, exist_ok=True)
os.makedirs(pdf_folder, exist_ok=True)

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
    "vendor_branch_patterns": [
        {"name": "Thai Branch", "regex": r"สาขา(?:ที่|เลขที่)?\s*[:：]?\s*(\d{1,10})"},
        {"name": "English Branch", "regex": r"Branch\s*(?:No\.?|Number|Code|ID)?\s*[:：]?\s*(\d{1,10})"}
    ],
    # Vendor-only: คำที่ระบุสาขาผู้ออกเอกสารโดยตรง
    "vendor_issuer_patterns": [
        {"name": "Thai Tax Invoice Issued By", "regex": r"ใบกำกับภาษี\s*ออกโดย"},
        {"name": "Thai Document Issued By", "regex": r"เอกสาร\s*ออกโดย"},
        {"name": "Thai Issued By", "regex": r"ออกโดย\s*[:：]?"},
        {"name": "English Issued By", "regex": r"\bISSUED\s+BY\b"},
        {"name": "English Invoice Issued By", "regex": r"\bINVOICE\s+ISSUED\s+BY\b"}
    ],
    # Customer-only: ใช้เฉพาะ CompanyBranch_invoice ห้ามใช้กับ VendorBranch
    "customer_branch_patterns": [
        {"name": "Thai Customer Branch", "regex": r"สาขา(?:ที่|เลขที่)?\s*[:：#-]?\s*(\d{1,10})"},
        {"name": "English Customer Branch", "regex": r"\bBRANCH\s*(?:NO\.?|NUMBER|CODE|ID)?\s*[:：#-]?\s*(\d{1,10})\b"}
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

# PO rules come only from patternsPO.json next to the shared loader.
# Invalid/missing configuration stops startup instead of using stale defaults.
PO_PATTERN_CONFIG = load_po_patterns()
PO_PATTERNS = [pattern for patterns in PO_PATTERN_CONFIG.values() for pattern in patterns]

VENDOR_BRANCH_PATTERNS = [p["regex"] for p in pattern_config.get("vendor_branch_patterns", DEFAULT_PATTERN_CONFIG["vendor_branch_patterns"])]
VENDOR_ISSUER_PATTERNS = [p["regex"] for p in pattern_config.get("vendor_issuer_patterns", DEFAULT_PATTERN_CONFIG["vendor_issuer_patterns"])]
CUSTOMER_BRANCH_PATTERNS = [p["regex"] for p in pattern_config.get("customer_branch_patterns", DEFAULT_PATTERN_CONFIG["customer_branch_patterns"])]
TAXID_PATTERNS = [p["regex"] for p in pattern_config.get("taxid_patterns", DEFAULT_PATTERN_CONFIG["taxid_patterns"])]
EXCLUDE_TAX_IDS = set(pattern_config.get("exclude_tax_ids", DEFAULT_PATTERN_CONFIG["exclude_tax_ids"]))
STOP_KEYWORDS = pattern_config.get("stop_keywords", DEFAULT_PATTERN_CONFIG["stop_keywords"])
SUPPLIER_SKIP_KEYWORDS = pattern_config.get("supplier_skip_keywords", DEFAULT_PATTERN_CONFIG["supplier_skip_keywords"])
SUPPLIER_NAME_PATTERNS = [p["regex"] for p in pattern_config.get("supplier_name_patterns", DEFAULT_PATTERN_CONFIG["supplier_name_patterns"])]

def normalize_internal_company_branch(value):
    """Use the same five-digit company branch code in CLI, PO rules and Excel."""
    text = str(value or "").strip()
    if not text:
        return ""
    if not re.fullmatch(r"[0-9]{5}", text):
        raise ValueError("Company branch must contain exactly 5 digits")
    return text


def company_branch_to_display(value):
    """Company branch codes are already in the five-digit Excel format."""
    return normalize_internal_company_branch(value)


def get_po_patterns_by_branch(branch_code):
    """Use only the configured masks for this five-digit company branch."""
    branch_code = normalize_internal_company_branch(branch_code)
    return PO_PATTERN_CONFIG.get(branch_code, [])


def get_other_branch(branch_code):
    branch_code = normalize_internal_company_branch(branch_code)

    if branch_code == "00000":
        return "00001"
    if branch_code == "00001":
        return "00000"

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
    # THAI KOITO Prachinburi: ถ้ามีเลขที่ 555
    # ให้ตัดทุกอย่างก่อน 555 ออกก่อนหาเลขที่บ้านแบบ Generic
    # ตัวอย่าง:
    #   1 : 555 MOO 7 ... -> 555 MOO 7 ...
    #   ADDRESS : 555 หมู่ 7 ... -> 555 หมู่ 7 ...
    # ใช้ word boundary เพื่อไม่จับ 555 ที่เป็นส่วนหนึ่งของเลขอื่น
    # ----------------------------------------------------------
    koito_555 = re.search(r"\b555\b", address)
    if koito_555:
        address = address[koito_555.start():].strip(" ,;:-")

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
    python script.py testfile 00000

    ถ้า CompanyAddress มีเลข 370
        expected = 00000

    ถ้าไม่มี 370
        expected = 00001

    return:
        Company_branch
        is_correct
        expected_branch
    """

    Company_address = str(Company_address or "").strip()
    input_branch = normalize_internal_company_branch(input_branch)

    # กำหนด branch ที่ควรจะเป็นจาก Address
    if re.search(r"\b(?:370|10570)\b", Company_address):
        expected_branch = "00000"
    else:
        expected_branch = "00001"

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



def clean_customer_address_postcode(address):
    """
    Clean CustomerAddress หลัง OCR

    - เริ่มตั้งแต่เลขที่ 555 ถ้ามี
    - ถ้าพบรหัสไปรษณีย์ 5 หลัก ให้จบที่เลข 5 หลักทันที
      แม้ OCR จะเอาขยะมาติดท้าย เช่น 25140abc -> 25140
    - สำหรับที่อยู่ Prachinburi/ปราจีนบุรี ถ้า OCR อ่าน 25140 เป็น
      2514 + อักขระขยะ ให้ normalize กลับเป็น 25140 แล้วจบ
    - ถ้าไม่มี postcode ให้ตัดเมื่อเจอ section ถัดไป เช่น Term/Tel/Fax/Tax ID
    """
    text = re.sub(r"\s+", " ", str(address or "")).strip(" ,;:-")
    if not text:
        return ""

    # ตัดทุกอย่างก่อนเลขที่ 555
    m555 = re.search(r"(?<!\d)555(?!\d)", text)
    if m555:
        text = text[m555.start():]

    # รหัสไปรษณีย์ 5 หลัก: ไม่ใช้ word boundary ด้านหลัง
    # เพื่อรองรับ 25140ulumsuis / 251400ulumsthis
    postal = re.search(r"(?<!\d)([1-9]\d{4})", text)
    if postal:
        text = text[:postal.end(1)]
        return re.sub(r"\s+", " ", text).strip(" ,;:-")

    # OCR ของที่อยู่ THAI KOITO บางใบอ่าน 25140 เหลือ 2514 แล้วมีขยะต่อท้าย
    # จำกัดการแก้เฉพาะที่อยู่ Prachinburi/ปราจีนบุรี เพื่อลดผลกระทบเอกสารอื่น
    if re.search(r"PRACHINBURI|ปราจีนบุรี", text, re.IGNORECASE):
        bad_2514 = re.search(r"(?<!\d)2514(?=\D|$)", text)
        if bad_2514:
            text = text[:bad_2514.start()] + "25140"
            return re.sub(r"\s+", " ", text).strip(" ,;:-")

    # ไม่มี postcode: ตัด section ที่ไม่ใช่ address ถ้าติดท้ายมา
    stop = re.search(
        r"\s+(?:TERM\b|TEL\b|TELEPHONE\b|PHONE\b|FAX\b|TAX\s*ID\b|"
        r"INVOICE\b|P/?O\s*(?:NO)?\b|DESCRIPTION\b|QUANTITY\b|UNIT\s+PRICE\b)",
        text,
        re.IGNORECASE,
    )
    if stop:
        text = text[:stop.start()]

    return re.sub(r"\s+", " ", text).strip(" ,;:-")


def is_valid_customer_address(address):
    """
    ตรวจว่า CustomerAddress ที่ Azure/OCR คืนมาดูเป็นที่อยู่จริงหรือไม่
    ป้องกันค่าหลุด เช่น "7" หรือเลขสั้น ๆ ถูกนำไปใช้เป็น CustomerAddress
    """
    text = re.sub(r"\s+", " ", str(address or "")).strip(" ,;:-")

    if not text:
        return False

    # เลขเดี่ยว/ตัวเลขล้วนไม่ถือเป็น address
    if re.fullmatch(r"\d{1,6}", text):
        return False

    # สั้นเกินไปมีโอกาสเป็น token จาก OCR มากกว่าที่อยู่
    if len(text) < 10:
        return False

    # ที่อยู่ควรมีเลขและตัวอักษรอย่างน้อยอย่างละหนึ่งส่วน
    if not re.search(r"\d", text):
        return False
    if not re.search(r"[A-Za-zก-๙]", text):
        return False

    return True


def extract_customer_address_by_taxid(invoices, customer_tax_id="", ocr_cache=None):
    """
    OCR fallback สำหรับ CustomerAddress โดยใช้ Customer Tax ID เป็น anchor
    แล้วหา address ที่เริ่มด้วยเลขที่ 555 ในบริเวณใกล้เคียง

    เหมาะกับเอกสารที่ Azure ไม่ map CustomerAddress และไม่มี label
    Account To / Bill To เช่น CBC / BKF

    สำคัญ: ค้นเฉพาะบริเวณ Customer Tax ID เพื่อไม่ดึง Vendor Address
    """
    lines = ocr_cache["lines"] if ocr_cache is not None else get_all_lines(invoices)
    customer_tax_id = clean_tax_id(customer_tax_id)

    if not customer_tax_id:
        return ""

    # หา Customer Tax ID เป็น anchor
    anchor_indexes = []
    for i, line in enumerate(lines):
        text = re.sub(r"\s+", " ", str(line or "")).strip()
        digits = re.sub(r"\D", "", text)
        if customer_tax_id and customer_tax_id in digits:
            anchor_indexes.append(i)

    if not anchor_indexes:
        return ""

    stop_patterns = [
        r"\bTEL\b", r"\bTELEPHONE\b", r"\bPHONE\b", r"\bFAX\b",
        r"\bTERM\s+OF\s+PAYMENT\b", r"\bDUE\s+DATE\b",
        r"\bINVOICE\s+(?:NO|NUMBER|DATE)\b", r"\bP/?O\s*(?:NO)?\b",
        r"\bDESCRIPTION\b", r"\bQUANTITY\b", r"\bUNIT\s+PRICE\b",
        r"\bDELIVERY\s+TO\b", r"\bSHIP\s+TO\b",
    ]

    for anchor in anchor_indexes:
        # จำกัด window รอบ Customer Tax ID เท่านั้น
        start = max(0, anchor - 4)
        end = min(len(lines), anchor + 10)
        window = lines[start:end]

        # หาเลขที่ 555 ซึ่งเป็นเลขที่ของ THAI KOITO ในเอกสารจริง
        addr_pos = None
        first_value = ""
        for j, line in enumerate(window):
            text = re.sub(r"\s+", " ", str(line or "")).strip()
            m = re.search(r"\b555\b", text)
            if m:
                addr_pos = j
                first_value = text[m.start():].strip(" ,;:-")
                break

        if addr_pos is None:
            continue

        address_lines = []
        if first_value:
            address_lines.append(first_value)

        # เก็บ address ต่ออีกไม่เกิน 3 บรรทัด
        for line in window[addr_pos + 1: addr_pos + 4]:
            text = re.sub(r"\s+", " ", str(line or "")).strip(" ,;:-")
            if not text:
                continue

            # ถ้าเป็น Tax ID / Branch / Phone/Fax / ตาราง ให้หยุด
            if re.search(r"\bTAX\s*ID\b|เลขประจำตัวผู้เสียภาษี", text, re.IGNORECASE):
                break
            if any(re.search(p, text, re.IGNORECASE) for p in stop_patterns):
                break

            address_lines.append(text)

            # เจอ postcode ของ Prachinburi แล้วพอ
            if re.search(r"\b2514\d\b", text):
                break

        result = " ".join(address_lines)
        result = re.sub(r"\s+", " ", result).strip(" ,;:-")

        if result:
            return clean_company_address_start(result)

    return ""

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


def extract_company_name_from_pages(invoices, ocr_cache=None):
    """OCR fallback สำหรับดึงชื่อ Company/Customer จากโซนลูกค้า"""

    lines = ocr_cache["lines"] if ocr_cache is not None else get_all_lines(invoices)

    patterns = [
        r"(?:Customer\s*Name|Company\s*Name)\s*[:：]\s*(.+)",
        r"(?:ชื่อลูกค้า|ชื่อผู้ซื้อ)\s*[:：]\s*(.+)",
    ]

    for line in lines:
        text = re.sub(r"\s+", " ", str(line or "")).strip()

        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if not match:
                continue

            company_name = match.group(1).strip(" ,;:-")
            if company_name:
                return company_name

    return ""



def is_valid_customer_company_name(name):
    """ตรวจว่าชื่อ Customer ดูเป็นชื่อบริษัทที่สมบูรณ์ ไม่ใช่ OCR fragment เช่น THAI"""
    text = re.sub(r"\s+", " ", str(name or "")).strip()
    if not text:
        return False

    thai_ok = bool(re.search(r"บริษัท.+?(?:จำกัด|มหาชน)", text, re.IGNORECASE))
    eng_ok = bool(re.search(
        r"\b(?:CO\.?\s*,?\s*LTD\.?|COMPANY\s+LIMITED|PUBLIC\s+COMPANY\s+LIMITED|LIMITED)\b",
        text,
        re.IGNORECASE,
    ))
    return thai_ok or eng_ok


def extract_company_name_by_taxid(invoices, company_tax_id="", ocr_cache=None):
    """
    หา Customer/Company Name โดยใช้ Customer Tax ID เป็น anchor
    รองรับกรณี OCR แยกชื่อบริษัทหลาย line/token และกรณีชื่ออยู่บรรทัดเดียวกับ Tax ID
    ไม่ hard-code ชื่อบริษัท และไม่ใช้ชื่อ Vendor เป็น Customer
    """
    lines = ocr_cache["lines"] if ocr_cache is not None else get_all_lines(invoices)
    tax_id = clean_tax_id(company_tax_id)

    if not re.fullmatch(r"0\d{12}", tax_id):
        return ""

    anchor_indexes = []
    for i, line in enumerate(lines):
        text = re.sub(r"\s+", " ", str(line or "")).strip()
        digits = re.sub(r"\D", "", text)
        if tax_id in digits:
            anchor_indexes.append(i)

    # รูปแบบชื่อบริษัทแบบ generic
    english_company = re.compile(
        r"([A-Z][A-Z0-9&.'(),\-/ ]{1,120}?"
        r"(?:PUBLIC\s+COMPANY\s+LIMITED|COMPANY\s+LIMITED|CO\.?\s*,?\s*LTD\.?|LIMITED))",
        re.IGNORECASE,
    )
    thai_company = re.compile(
        r"(บริษัท\s+.{1,120}?(?:จำกัด(?:\s*\(\s*มหาชน\s*\))?|มหาชน))",
        re.IGNORECASE,
    )

    for anchor in anchor_indexes:
        # เน้นหลัง Tax ID เพราะหลาย template วางชื่อ Customer ถัดลงมา
        # แต่เผื่อ OCR reorder จึงดูย้อนหลังเล็กน้อย
        start = max(0, anchor - 1)
        end = min(len(lines), anchor + 7)
        window_lines = [
            re.sub(r"\s+", " ", str(lines[i] or "")).strip()
            for i in range(start, end)
            if str(lines[i] or "").strip()
        ]

        # รวมหลาย OCR lines ก่อนค้นหา ป้องกัน THAI / KOITO CO., LTD. ถูกแยกกัน
        window_text = " ".join(window_lines)

        # ลบ Tax ID และ label ที่ไม่ใช่ชื่อออก แต่ไม่ทิ้งทั้งบรรทัด
        window_text = re.sub(re.escape(tax_id), " ", window_text)
        window_text = re.sub(
            r"(?:TAX\s*ID|TAXID|เลขประจำตัวผู้เสียภาษี|CUSTOMER\s*ID|รหัสลูกค้า)"
            r"\s*[:：#-]?\s*\d*",
            " ",
            window_text,
            flags=re.IGNORECASE,
        )
        window_text = re.sub(
            r"\(?\s*BRANCH\s*(?:NO\.?|NUMBER|CODE|ID)?\s*[:：#-]?\s*\d+\s*\)?",
            " ",
            window_text,
            flags=re.IGNORECASE,
        )
        window_text = re.sub(r"\s+", " ", window_text).strip()

        # หาไทยก่อน/อังกฤษจากข้อความรวม
        matches = []
        for pattern in (thai_company, english_company):
            for m in pattern.finditer(window_text):
                candidate = clean_company_name(m.group(1))
                if is_valid_customer_company_name(candidate):
                    matches.append(candidate)

        if matches:
            # เลือกชื่อที่กระชับที่สุด เพื่อลดการกิน label/ข้อความก่อนหน้า
            matches.sort(key=len)
            return matches[0]

        # fallback: ตรวจทีละชุด 1-3 บรรทัดหลัง anchor
        for i in range(anchor, min(len(lines), anchor + 6)):
            combined = " ".join(
                re.sub(r"\s+", " ", str(lines[j] or "")).strip()
                for j in range(i, min(len(lines), i + 3))
                if str(lines[j] or "").strip()
            )
            combined = re.sub(re.escape(tax_id), " ", combined)
            combined = re.sub(r"\s+", " ", combined).strip()
            for pattern in (thai_company, english_company):
                m = pattern.search(combined)
                if m:
                    candidate = clean_company_name(m.group(1))
                    if is_valid_customer_company_name(candidate):
                        return candidate

    return ""

def extract_company_tax_id_from_pages(invoices, ocr_cache=None, vendor_tax_id=""):
    """
    OCR fallback สำหรับ CustomerTaxID โดยค้นเฉพาะ Customer zone
    และรองรับกรณี OCR แยกคำว่า Tax ID กับเลข 13 หลักคนละบรรทัด

    กฎสำคัญ:
    - ไม่ใช้ VendorTaxId เป็น CustomerTaxID
    - ไม่หยิบ Tax ID จากส่วน Vendor ด้านบนแบบสุ่ม
    - รองรับ Customer / Bill To / Account To / Sold To / Invoice To / Buyer
    - รองรับเลข 13 หลักที่มี space / - / OCR O,I,l ปน
    """
    lines = ocr_cache["lines"] if ocr_cache is not None else get_all_lines(invoices)
    vendor_tax_id = clean_tax_id(vendor_tax_id)

    customer_patterns = [
        r"\bCUSTOMER(?:\s+(?:NAME|ADDRESS|ID|CODE))?\b",
        r"\bACCOUNT\s+TO\b",
        r"\bBILL\s+TO\b",
        r"\bSOLD\s+TO\b",
        r"\bINVOICE\s+TO\b",
        r"\bBUYER\b",
        r"\bPURCHASER\b",
        r"ชื่อลูกค้า",
        r"ที่อยู่ลูกค้า",
        r"รหัสลูกค้า",
        r"ชื่อผู้ซื้อ",
        r"ผู้ซื้อ",
    ]

    stop_patterns = [
        r"\bSHIP\s+TO\b",
        r"\bDELIVERY\s+TO\b",
        r"\bDESCRIPTION\b",
        r"\bPART\s+NO\b",
        r"\bPRODUCT\s+NO\b",
        r"\bITEM\b",
        r"\bQUANTITY\b",
        r"\bUNIT\s+PRICE\b",
    ]

    tax_label_pattern = re.compile(
        r"(?:TAX\s*(?:ID|NO)|TAXID|VAT\s*(?:ID|NO)|TIN|"
        r"เลขประจำตัวผู้เสียภาษี(?:อากร)?)",
        re.IGNORECASE,
    )

    def extract_valid_tax_id(text):
        # clean_tax_id รองรับ O/o -> 0 และ I/l -> 1
        digits = clean_tax_id(text)
        for m in re.finditer(r"0\d{12}", digits):
            tax_id = m.group(0)
            if vendor_tax_id and tax_id == vendor_tax_id:
                continue
            return tax_id
        return ""

    # ==========================================================
    # 1) หา Customer zone จาก label ที่ชัดเจน
    # ==========================================================
    customer_starts = []
    for i, line in enumerate(lines):
        text = re.sub(r"\s+", " ", str(line or "")).strip()
        if text and any(re.search(p, text, re.IGNORECASE) for p in customer_patterns):
            customer_starts.append(i)

    for customer_start in customer_starts:
        end = min(len(lines), customer_start + 25)

        for i in range(customer_start, end):
            text = re.sub(r"\s+", " ", str(lines[i] or "")).strip()
            if not text:
                continue

            if i > customer_start and any(
                re.search(p, text, re.IGNORECASE) for p in stop_patterns
            ):
                break

            # Tax ID อยู่บรรทัดเดียวกับ label
            if tax_label_pattern.search(text):
                tax_id = extract_valid_tax_id(text)
                if tax_id:
                    return tax_id

                # OCR อาจแยกเลข Tax ID ไป 1-3 บรรทัดถัดไป
                combined = " ".join(
                    re.sub(r"\s+", " ", str(lines[j] or "")).strip()
                    for j in range(i, min(i + 4, end))
                    if str(lines[j] or "").strip()
                )
                tax_id = extract_valid_tax_id(combined)
                if tax_id:
                    return tax_id

        # ใน Customer zone บาง template ไม่มีคำว่า Tax ID ชัดเจน
        # จึงตรวจเลข 13 หลักใน zone แต่ยังห้ามเท่ากับ VendorTaxId
        zone_text = " ".join(
            re.sub(r"\s+", " ", str(lines[j] or "")).strip()
            for j in range(customer_start, end)
            if str(lines[j] or "").strip()
        )
        tax_id = extract_valid_tax_id(zone_text)
        if tax_id:
            return tax_id

    # ==========================================================
    # 2) Fallback จาก Customer Address 555
    # ใช้เฉพาะบริเวณรอบ 555 และยังกัน VendorTaxId ออก
    # ==========================================================
    for i, line in enumerate(lines):
        text = re.sub(r"\s+", " ", str(line or "")).strip()
        if not re.search(r"\b555\b", text):
            continue

        start = max(0, i - 5)
        end = min(len(lines), i + 8)
        window = " ".join(
            re.sub(r"\s+", " ", str(lines[j] or "")).strip()
            for j in range(start, end)
            if str(lines[j] or "").strip()
        )

        # ให้ความสำคัญกับ Tax ID label ใน window
        if tax_label_pattern.search(window):
            tax_id = extract_valid_tax_id(window)
            if tax_id:
                return tax_id

    return ""

def normalize_company_branch_invoice(value):
    """
    แปลงข้อความ Branch ที่อ่านจาก Invoice ให้เป็นรหัสสาขา 5 หลัก

    ตัวอย่าง:
    Branch: 00001       -> 00001
    Branch : 1          -> 00001
    Branch No. 00001    -> 00001
    สาขา 00001          -> 00001
    สาขาที่ 1           -> 00001
    Head Office         -> 00000
    สำนักงานใหญ่        -> 00000
    00001               -> 00001
    """

    text = re.sub(r"\s+", " ", str(value or "")).strip()

    if not text:
        return ""

    # =========================================================
    # Head Office
    # =========================================================
    if re.search(
        r"สำนักงานใหญ่|HEAD\s*OFFICE|HEADOFFICE",
        text,
        re.IGNORECASE
    ):
        return "00000"

    # =========================================================
    # Branch
    # =========================================================
    patterns = [
        # Branch: 00001
        # Branch : 00001
        # Branch No. 00001
        # Branch No: 00001
        # Branch Number 00001
        r"\bBRANCH\s*(?:NO\.?|NUMBER|CODE|ID)?\s*[:：#-]?\s*(\d{1,10})\b",

        # สาขา 00001
        # สาขาที่ 00001
        # สาขาเลขที่ 00001
        r"สาขา(?:ที่|เลขที่)?\s*[:：#-]?\s*(\d{1,10})",
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)

        if match:
            branch = match.group(1)

            # ทำให้เป็น 5 หลัก
            return branch.zfill(5)[-5:]

    # IMPORTANT:
    # ฟังก์ชันนี้ใช้กับ CustomerAddress ดังนั้นห้ามนำเลขเดี่ยว/ตัวเลขล้วน
    # เช่น "7" จาก "MOO 7" มาแปลงเป็น Branch 0007
    # Branch ต้องมีคำว่า Branch / สาขา กำกับเท่านั้น
    return ""

def extract_company_branch_invoice_from_pages(invoices, ocr_cache=None, company_tax_id=""):
    """
    ดึง Branch ของ Customer จาก Invoice โดยใช้ Customer Tax ID เป็น anchor

    ไม่ใช้ prebuilt-layout และไม่ยิง Azure เพิ่ม
    ใช้ OCR lines จากผลลัพธ์ prebuilt-invoice รอบเดียว

    ตัวอย่าง:
        TAX ID No. 0105529030059 Branch: 00001

    ผลลัพธ์:
        00001
    """

    lines = (
        ocr_cache["lines"]
        if ocr_cache is not None
        else get_all_lines(invoices)
    )

    # =========================================================
    # Customer Tax ID
    # =========================================================
    company_tax_id = clean_tax_id(company_tax_id)

    print(
        f"🔎 Search CompanyBranch_invoice using CustomerTaxID = {company_tax_id}"
    )

    # =========================================================
    # Branch Regex
    # =========================================================
    # Customer-only patterns: แยกจาก VendorBranch โดยเด็ดขาด
    branch_patterns = list(CUSTOMER_BRANCH_PATTERNS)
    # OCR fallback ของ Customer เท่านั้น
    branch_patterns.append(r"\bBRANCH\s+(\d{1,10})\b")

    def find_branch(text):
        text = re.sub(r"\s+", " ", str(text or "")).strip()

        if not text:
            return ""

        for pattern in branch_patterns:
            match = re.search(pattern, text, re.IGNORECASE)

            if match:
                branch = match.group(1)

                if branch:
                    return branch.zfill(5)[-5:]

        return ""

    # =========================================================
    # วิธีที่ 1: หา Customer Tax ID ก่อน
    # แล้วตรวจ Branch ในบรรทัดเดียวกัน / ใกล้เคียง
    # =========================================================
    if company_tax_id:

        for i, line in enumerate(lines):
            text = re.sub(r"\s+", " ", str(line or "")).strip()

            if not text:
                continue

            # clean_tax_id จะเหลือเฉพาะตัวเลข
            # จึงใช้ตรวจว่า OCR line นี้มี Customer Tax ID หรือไม่
            clean_line_tax = clean_tax_id(text)

            if company_tax_id in clean_line_tax:
                print(
                    f"🔎 CustomerTaxID found at OCR line {i}: {text}"
                )

                # Branch อยู่บรรทัดเดียวกับ Tax ID
                branch = find_branch(text)

                if branch:
                    print(
                        f"✅ CompanyBranch_invoice found same line: {branch}"
                    )
                    return branch

                # Branch อาจอยู่ก่อนหรือหลัง Customer Tax ID
                # เช่น DAISIN: ชื่อ/ที่อยู่ Customer มี "(สาขาที่ 00001)"
                # แล้ว TAX ID อยู่บรรทัดถัดลงมา
                start_index = max(0, i - 6)
                end_index = min(len(lines), i + 5)

                for j in range(start_index, end_index):
                    if j == i:
                        continue

                    nearby_text = re.sub(
                        r"\s+", " ", str(lines[j] or "")
                    ).strip()

                    branch = find_branch(nearby_text)

                    if branch:
                        print(
                            f"✅ CompanyBranch_invoice found near CustomerTaxID "
                            f"line {j}: {branch} | {nearby_text}"
                        )
                        return branch

    # =========================================================
    # วิธีที่ 2: Fallback หา Customer/Sold To
    # ใช้กรณี Tax ID ไม่ถูกอ่านเป็น field
    # =========================================================
    customer_patterns = [
        r"\bINVOICE\s+TO\b",
        r"\bSOLD\s+TO\b",
        r"\bCUSTOMER(?:\s+NAME)?\b",
        r"\bACCOUNT\s+TO\b",
        r"\bBILL\s+TO\b",
        r"\bBUYER\b",
        r"\bPURCHASER\b",
        r"ชื่อลูกค้า",
        r"ชื่อผู้ซื้อ",
        r"ผู้ซื้อ",
    ]

    customer_start = -1

    for i, line in enumerate(lines):
        text = re.sub(r"\s+", " ", str(line or "")).strip()

        if any(
            re.search(pattern, text, re.IGNORECASE)
            for pattern in customer_patterns
        ):
            customer_start = i

            print(
                f"🔎 Customer zone fallback found: {text}"
            )
            break

    # =========================================================
    # Scan 20 lines หลัง Customer/Sold To
    # =========================================================
    if customer_start >= 0:
        end_index = min(customer_start + 21, len(lines))

        for i in range(customer_start, end_index):
            text = re.sub(
                r"\s+", " ", str(lines[i] or "")
            ).strip()

            branch = find_branch(text)

            if branch:
                print(
                    f"✅ CompanyBranch_invoice fallback found: {branch}"
                )
                return branch

    # =========================================================
    # ไม่พบ
    # =========================================================
    print("⚠️ CompanyBranch_invoice not found")

    return ""


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
    - Azure เรียกเฉพาะ prebuilt-invoice; ไม่เรียก prebuilt-layout
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
    
    
    # Business rule: BOI code in invoice line items
    # Examples: [BOI06], BOI06, BOI 06, BOI-06
    if re.search(
        r"(?<![A-Za-z0-9])\[?BOI[\s-]*\d{1,3}\]?(?![A-Za-z0-9])",
        full_text,
        re.IGNORECASE
    ):
        return "หักภาษี ณ ที่จ่าย"

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

def clean_company_name(name):
    """
    ทำความสะอาด Customer / Company Name

    ภาษาไทย:
        นามลูกค้า บริษัท ไทย โคะอิโท จำกัด (สำนักงานใหญ่)
        -> บริษัท ไทย โคะอิโท จำกัด

    ภาษาอังกฤษ:
        THAI KOITO CO., LTD. (HEAD OFFICE)
        -> THAI KOITO CO., LTD.

        THAI KOITO COMPANY LIMITED HEAD OFFICE
        -> THAI KOITO COMPANY LIMITED
    """

    if not name:
        return ""

    name = str(name)

    # กรณีข้อมูลมาจาก HTML / Excel
    name = re.sub(r"<br\s*/?>", " ", name, flags=re.IGNORECASE)

    # รวม space
    name = re.sub(r"\s+", " ", name).strip(" ,;:-")

    # ==========================================================
    # ภาษาไทย
    # ==========================================================
    company_pos = name.find("บริษัท")

    if company_pos >= 0:
        # ตัดทุกอย่างก่อนคำว่า บริษัท
        name = name[company_pos:].strip()

        # เก็บถึงคำว่า จำกัด
        match = re.search(
            r"บริษัท.*?จำกัด(?:\s*\(\s*มหาชน\s*\))?",
            name,
            re.IGNORECASE,
        )

        if match:
            return re.sub(
                r"\s+",
                " ",
                match.group(0)
            ).strip(" ,;:-")

    # ==========================================================
    # ภาษาอังกฤษ
    # ==========================================================
    english_endings = [
        # PUBLIC COMPANY LIMITED
        r"\bPUBLIC\s+COMPANY\s+LIMITED\b",

        # COMPANY LIMITED
        r"\bCOMPANY\s+LIMITED\b",

        # CO., LTD. / CO.,LTD / CO LTD / CO. LTD.
        r"\bCO\.?\s*,?\s*LTD\.?",

        # LIMITED
        r"\bLIMITED\b",

        # LTD.
        r"\bLTD\.?",
    ]

    for pattern in english_endings:
        match = re.search(pattern, name, re.IGNORECASE)

        if match:
            name = name[:match.end()]

            name = re.sub(r"\s+", " ", name).strip(" ,;:-")

            return name

    return name

# ============================================================
# VendorBranch diagnostics (local log only; no additional AI calls)
# Set to False after troubleshooting.
# ============================================================
DEBUG_VENDOR_BRANCH = True
print(f'📌 VendorBranch patched script: {os.path.abspath(__file__)} | version 2026-11-03-branch-v5-issuer-priority')

def vendor_branch_debug(event, **details):
    if not DEBUG_VENDOR_BRANCH:
        return
    try:
        log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'DebugVendorBranch')
        os.makedirs(log_dir, exist_ok=True)
        payload = {'timestamp': datetime.now().isoformat(timespec='seconds'), 'event': event, **details}
        with open(os.path.join(log_dir, 'vendor_branch_debug.jsonl'), 'a', encoding='utf-8') as handle:
            handle.write(json.dumps(payload, ensure_ascii=False, default=str) + '\n')
        print(f'[VENDOR DEBUG] {event}')
    except Exception as exc:
        print(f'⚠️ VendorBranch debug logging failed: {exc}')

def extract_vendor_branch(invoices, layout_result=None, ocr_cache=None,
                          vendor_tax_id="", supplier_name="", pdf_path=None):
    """Resolve the *seller's* issuing branch from ONE prebuilt-invoice response.

    A vendor's directory of offices is not proof of which office issued an invoice.
    Local PDF pixel inspection (when available) is NOT an Azure/API request.
    """
    vendor_tax_id = clean_tax_id(vendor_tax_id)
    customer_ids = {clean_tax_id(x) for x in EXCLUDE_TAX_IDS if clean_tax_id(x)}
    head_re = re.compile(r"สำนักงานใหญ่|HEAD\s*OFFICE|HEADOFFICE", re.I)
    number_re = re.compile(r"(?:สาขา(?:ที่|เลขที่)?|BRANCH\s*(?:NO\.?|NUMBER|CODE|ID)?)\s*[:：#-]?\s*(\d{1,5})", re.I)
    issuer_re = re.compile(r"ออกโดย|สาขาที่ออก|สาขาผู้ออก|ISSUED\s+BY", re.I)
    buyer_re = re.compile(r"\b(?:CUSTOMER|BILL\s+TO|SHIP\s+TO|SOLD\s+TO|ACCOUNT\s+TO|INVOICE\s+TO|BUYER|PURCHASER|DELIVERY\s+TO)\b|ชื่อลูกค้า|ชื่อผู้ซื้อ|ผู้ซื้อ|รหัสลูกค้า|ส่งให้|ส่งถึง|ผู้รับสินค้า|ชื่อลูกค้า|ที่อยู่ลูกค้า", re.I)
    tick_re = re.compile(r"☑|☒|✓|✔|\[\s*[xX✓✔]\s*\]|\(\s*[xX✓✔]\s*\)")

    def bounds(obj):
        pts = []
        polygon = getattr(obj, 'polygon', None) or []
        if polygon and isinstance(polygon[0], (int, float)):
            polygon = list(zip(polygon[::2], polygon[1::2]))
        for pt in polygon:
            try:
                pts.append((float(pt.x), float(pt.y)) if hasattr(pt, 'x')
                           else (float(pt[0]), float(pt[1])))
            except (TypeError, ValueError, IndexError):
                pass
        if not pts:
            return None
        return min(x for x, y in pts), min(y for x, y in pts), max(x for x, y in pts), max(y for x, y in pts)

    def parse_options(text):
        result = []
        if head_re.search(text):
            result.append('00000')
        for m in number_re.finditer(text):
            val = m.group(1).zfill(5)
            if val not in result:
                result.append(val)
        return result

    pages = list(getattr(invoices, 'pages', []) or [])
    rows = []
    for pi, page in enumerate(pages):
        for line in (getattr(page, 'lines', []) or []):
            txt = re.sub(r'\s+', ' ', str(getattr(line, 'content', '') or '')).strip()
            if txt:
                rows.append((pi, txt, bounds(line)))
    if not rows:
        rows = [(0, str(t).strip(), None) for t in (ocr_cache or {}).get('lines', []) if str(t).strip()]
    vendor_branch_debug('OCR_ROWS', pdf_path=str(pdf_path or ''), vendor_tax_id=vendor_tax_id, supplier_name=supplier_name, rows=[{'page': pi+1, 'text': txt, 'bounds': rect} for pi, txt, rect in rows])
    if not rows:
        vendor_branch_debug('NO_OCR_ROWS', pdf_path=str(pdf_path or ''))
        return ''

    def buyer_tax(text):
        digits = clean_tax_id(text)
        return any(t in digits for t in customer_ids)

    # Highest priority: an explicit seller-issued tax invoice branch statement.
    # Example: "สาขาที่ออกใบกำกับภาษี : สาขาที่ 00003" (IRC).
    # This outranks office-address lists and checkbox heuristics.
    explicit_issuer_label = re.compile(
        r"สาขา\s*(?:ที่\s*)?ออก\s*(?:ใบ\s*กำกับ\s*ภาษี|ใบ\s*ส่ง\s*สินค้า|เอกสาร)"
        r"|สาขา\s*ผู้\s*ออก\s*(?:ใบ\s*กำกับ\s*ภาษี|เอกสาร)?"
        r"|ใบ\s*กำกับ\s*ภาษี\s*ออก\s*โดย"
        r"|ที่\s*อ่อน\s*ใช้\s*กำกับ\s*ภาษี"  # damaged Thai OCR: ที่อ่อนใช้กำกับภาษี
        r"|TAX\s*INVOICE\s*ISSUING\s*BRANCH"
        r"|INVOICE\s*ISSUING\s*BRANCH",
        re.I,
    )
    branch_value_after_label = re.compile(
        r"(?:สาขา\s*(?:ที่|เลขที่)?|BRANCH\s*(?:NO\.?|NUMBER)?)\s*[:：#-]?\s*(\d{1,5})",
        re.I,
    )
    for i, (pi, txt, _) in enumerate(rows):
        label_match = explicit_issuer_label.search(txt)
        if not label_match or buyer_re.search(txt) or buyer_tax(txt):
            continue
        tail = txt[label_match.end():]
        # HEAD OFFICE is an explicit value, NOT a prompt to read the next row's
        # item number / quantity. This caused YPC 00001 and 00008.
        if head_re.search(tail):
            vendor_branch_debug('EXPLICIT_TAX_INVOICE_ISSUER', pdf_path=str(pdf_path or ''),
                                text=txt, result='00000', evidence='head_office_same_line')
            return '00000'
        candidates = [m.group(1).zfill(5) for m in branch_value_after_label.finditer(tail)]
        if not candidates:
            bare = re.match(r"\s*[:：#-]?\s*(\d{1,5})\b", tail)
            if bare:
                candidates = [bare.group(1).zfill(5)]
        if not candidates:
            # Only a labelled branch or HEAD OFFICE on the next line qualifies.
            # Never consume a standalone number from the next OCR line.
            for pj, following, _ in rows[i + 1:i + 3]:
                if pj != pi or buyer_re.search(following) or buyer_tax(following):
                    break
                if head_re.search(following):
                    candidates = ['00000']
                    break
                matches = branch_value_after_label.findall(following)
                if matches:
                    candidates = [v.zfill(5) for v in matches]
                    break
        if len(set(candidates)) == 1:
            value = candidates[0]
            vendor_branch_debug('EXPLICIT_TAX_INVOICE_ISSUER', pdf_path=str(pdf_path or ''),
                                text=txt, result=value, evidence='issuer_branch_label')
            return value
        vendor_branch_debug('EXPLICIT_ISSUER_UNRESOLVED', pdf_path=str(pdf_path or ''),
                            text=txt, candidates=candidates)

    # Seller header office declaration can be on a separate line from the
    # supplier name (NITTO MATEX: HEAD OFFICE : NO. 700/611).
    # Do not apply to headers listing several offices/branches; those need
    # explicit issuing evidence or a marked checkbox.
    first_buyer = next((i for i, (_, text, _) in enumerate(rows)
                        if buyer_re.search(text) or buyer_tax(text)), len(rows))
    seller_prefix = rows[:first_buyer]
    seller_offices = [(pi, text, rect) for pi, text, rect in seller_prefix
                      if head_re.search(text)]
    seller_other_branches = [(pi, text, rect) for pi, text, rect in seller_prefix
                             if number_re.search(text) and
                             any(v != '00000' for v in parse_options(text))]
    if seller_offices and not seller_other_branches:
        vendor_branch_debug('SELLER_HEADER_UNIQUE_HEAD_OFFICE',
                            pdf_path=str(pdf_path or ''),
                            evidence=[text for _, text, _ in seller_offices],
                            result='00000')
        return '00000'

    # Explicit supplier-name anchor, restricted to top-of-page, excludes buyer rows.
    # This remains below the explicit issuing-branch priority.
    supplier_head_rows = []
    supplier_tokens = [w for w in re.findall(r"[A-Za-z]{4,}|[ก-๙]{4,}", str(supplier_name or '').upper())
                       if w not in {'COMPANY', 'LIMITED', 'PUBLIC', 'บริษัท', 'จำกัด'}]
    for pi, txt, rect in rows:
        if buyer_re.search(txt) or any(cid in re.sub(r"\D", "", txt) for cid in customer_ids):
            continue
        if not head_re.search(txt):
            continue
        if rect and rect[1] > 2.0:  # top header, page coordinates in inches
            continue
        if supplier_tokens and any(token in txt.upper() for token in supplier_tokens):
            supplier_head_rows.append({'page': pi+1, 'text': txt, 'bounds': rect})
    if supplier_head_rows:
        vendor_branch_debug('SELLER_NAME_HEAD_OFFICE', pdf_path=str(pdf_path or ''),
                            matches=supplier_head_rows, result='00000', evidence='supplier_name_anchor')
        return '00000'

    # A seller's company-name line explicitly marked HEAD OFFICE is stronger
    # evidence than branch numbers in the buyer's INVOICE TO block.
    # Restrict this to the vendor header BEFORE the first buyer-zone marker.
    # This is deliberately generic: it does not hardcode DAISIN or any tax ID.
    seller_header = []
    for pi, txt, rect in rows:
        if buyer_re.search(txt) or any(cid in re.sub(r"\D", "", txt) for cid in customer_ids):
            break
        seller_header.append((pi, txt, rect))
    company_name_re = re.compile(
        r"บริษัท\s*.+?(?:จำกัด|มหาชน)|\b(?:CO\.?\s*,?\s*LTD\.?|COMPANY\s+LIMITED|PUBLIC\s+COMPANY\s+LIMITED)\b",
        re.I,
    )
    header_name_head_office = [
        {'page': pi + 1, 'text': txt, 'bounds': rect}
        for pi, txt, rect in seller_header
        if head_re.search(txt) and company_name_re.search(txt)
        and not re.search(r"\b(?:BILL|INVOICE|SHIP|SOLD)\s+TO\b|ผู้ซื้อ|ลูกค้า", txt, re.I)
    ]
    # Do not override an explicit issuing-branch declaration later in the
    # header (e.g. IRC invoices with multiple registered office addresses).
    header_has_issuer = any(
        re.search(r"สาขา\s*(?:ที่\s*)?ออก\s*(?:ใบ\s*กำกับ\s*ภาษี|ใบ\s*ส่ง\s*สินค้า|เอกสาร)|ใบ\s*กำกับ\s*ภาษี\s*ออก\s*โดย|TAX\s*INVOICE\s*ISSUING\s*BRANCH", txt, re.I)
        for _, txt, _ in seller_header
    )
    if header_name_head_office and not header_has_issuer:
        vendor_branch_debug('SELLER_NAME_HEAD_OFFICE',
                            pdf_path=str(pdf_path or ''),
                            matches=header_name_head_office, result='00000')
        print('✅ VendorBranch seller name marked Head Office: 00000')
        return '00000'

    # Seller tax-registration footer: e.g. UNION NIFCO's
    # "สำนักงานใหญ่ TAX ID : 0105531086805" near the bottom of the page.
    # The footer is OUTSIDE the header/vendor-zone used by generic detection.
    # Match the seller's own 13-digit Tax ID on the SAME line as head-office,
    # never a customer Tax ID or a generic "HEAD OFFICE" elsewhere.
    if re.fullmatch(r"\d{13}", vendor_tax_id) and vendor_tax_id not in customer_ids:
        footer_candidates = []
        for pi, txt, rect in rows:
            if not head_re.search(txt):
                continue
            digits = re.sub(r"\D", "", txt)
            if vendor_tax_id not in digits:
                continue
            if any(cid in digits for cid in customer_ids):
                continue
            if re.search(r"(?:BRANCH|สาขา(?:ที่|เลขที่)?)\s*[:：#-]?\s*\d{1,5}", txt, re.I):
                continue
            footer_candidates.append({'page': pi + 1, 'text': txt, 'bounds': rect})
        if footer_candidates:
            vendor_branch_debug('SELLER_TAX_ID_HEAD_OFFICE',
                                pdf_path=str(pdf_path or ''),
                                vendor_tax_id=vendor_tax_id,
                                matches=footer_candidates,
                                result='00000')
            print('✅ VendorBranch seller Tax ID + Head Office: 00000')
            return '00000'


    # Only a branch stamp next to an invoice NUMBER is issuing evidence.
    # Never treat a branch address in the vendor letterhead as the issuer.
    # OCR may split "เลขที่", the invoice number and "สาขาที่ 00003"
    # into three consecutive rows. Require an actual invoice-number token to
    # avoid interpreting a registered-office directory as the issuing branch.
    invoice_number_re = re.compile(r"^(?:[A-Z]{1,4}\d{5,}[/-]?\d*|\d{8,12})$", re.I)
    branch_stamp_re = re.compile(r"^\s*(?:สาขา(?:ที่|เลขที่)?|BRANCH\s*(?:NO\.?)?)\s*[:：#-]?\s*(\d{1,5})\s*$", re.I)
    for i, (pi, txt, rect) in enumerate(rows):
        if not re.search(r"(?:เลขที่|INVOICE\s*(?:NO\.?|NUMBER))", txt, re.I):
            continue
        if buyer_tax(txt):
            continue
        nearby = [t for pj, t, _ in rows[i:i+5] if pj == pi]
        has_invoice = any(invoice_number_re.fullmatch(t.strip()) for t in nearby)
        has_invoice |= bool(re.search(r"(?:เลขที่|INVOICE\s*(?:NO\.?|NUMBER))\s*[:：#-]?\s*[A-Z]{1,4}\d{5,}", txt, re.I))
        if not has_invoice:
            continue
        for near in nearby:
            m = branch_stamp_re.fullmatch(near)
            if m:
                result = m.group(1).zfill(5)
                vendor_branch_debug('ISSUER_NEAR_INVOICE', pdf_path=str(pdf_path or ''), invoice_context=nearby, result=result)
                return result

    for i, (pi, txt, rect) in enumerate(rows):
        if issuer_re.search(txt) and not buyer_re.search(txt) and not buyer_tax(txt):
            values = parse_options(txt)
            if len(values) == 1:
                print(f'✅ VendorBranch explicit issuer: {values[0]}')
                return values[0]
            for pj, near, _ in rows[i+1:i+3]:
                if pj == pi and not buyer_re.search(near) and not buyer_tax(near):
                    values = parse_options(near)
                    if len(values) == 1:
                        print(f'✅ VendorBranch issuer continuation: {values[0]}')
                        return values[0]

    # Restrict office choices to the vendor letterhead. Avoid customer HEAD OFFICE.
    vendor_rows = []
    for pi, txt, rect in rows:
        if buyer_re.search(txt) or buyer_tax(txt):
            break
        vendor_rows.append((pi, txt, rect))
    if not vendor_rows:
        print('⚠️ VendorBranch: no vendor header isolated')
        return ''

    options = []
    for pi, txt, rect in vendor_rows:
        values = parse_options(txt)
        if len(values) == 1:
            options.append((pi, values[0], txt, rect))
        elif len(values) > 1:
            # Never assign one choice from a merged multi-option OCR line.
            print(f'🔎 VendorBranch multiple choices in OCR line: {txt}')
    vendor_branch_debug('OCR_OPTIONS', pdf_path=str(pdf_path or ''), options=[{'page': pi+1, 'branch': val, 'text': txt, 'bounds': rect} for pi, val, txt, rect in options], vendor_rows=[{'page': pi+1, 'text': txt, 'bounds': rect} for pi, txt, rect in vendor_rows])
    if not options:
        vendor_branch_debug('NO_OPTIONS', pdf_path=str(pdf_path or ''))
        return ''

    selected = []
    for pi, val, txt, rect in options:
        if tick_re.search(txt[:25]):
            selected.append(val)
    # Azure selection_marks are optional in prebuilt-invoice. Check only marks
    # horizontally aligned to a vendor choice, not marks elsewhere on the page.
    for pi, page in enumerate(pages):
        vendor_branch_debug('AZURE_SELECTION_MARKS', pdf_path=str(pdf_path or ''), page=pi+1, marks=[{'state': str(getattr(mark, 'state', '')), 'bounds': bounds(mark)} for mark in (getattr(page, 'selection_marks', []) or [])])
        candidates = [(v, t, r) for p, v, t, r in options if p == pi and r]
        for mark in (getattr(page, 'selection_marks', []) or []):
            state = str(getattr(mark, 'state', '') or '').lower()
            if 'selected' not in state or 'unselected' in state:
                continue
            mb = bounds(mark)
            if not mb:
                continue
            mx, my = (mb[0]+mb[2])/2, (mb[1]+mb[3])/2
            scored = []
            for val, txt, r in candidates:
                line_h = max(r[3]-r[1], mb[3]-mb[1], 0.01)
                dy = abs(my-(r[1]+r[3])/2)/line_h
                dx = (r[0]-mx)/line_h
                if dy <= 0.85 and -1.0 <= dx <= 9:
                    scored.append((dy+max(0, dx)*0.025, val))
            scored.sort()
            if scored and (len(scored) == 1 or scored[1][0]-scored[0][0] > .22):
                selected.append(scored[0][1])

    # Local PDF image fallback for prebuilt-invoice OCR without polygons.
    # Match a vertical column of real checkbox outlines to the ordered vendor
    # branch choices, then compare their INSIDE ink (not their printed borders).
    if not selected and pdf_path and len({v for _, v, _, _ in options}) > 1:
        try:
            import cv2
            import numpy as np
            from pdf2image import convert_from_path
            from itertools import combinations

            for pi in sorted({p for p, _, _, _ in options}):
                page_options = [(v, t, r) for p, v, t, r in options if p == pi]
                if len(page_options) != 2:
                    vendor_branch_debug('IMAGE_SKIP_AMBIGUOUS_OPTIONS', pdf_path=str(pdf_path or ''), page=pi+1, count=len(page_options))
                    continue
                # PyMuPDF renders locally without Poppler. Fallback to the
                # original pdf2image path when PyMuPDF is not installed.
                try:
                    import fitz
                    with fitz.open(str(pdf_path)) as doc:
                        if pi >= len(doc):
                            raise IndexError(f'PDF page {pi+1} not found')
                        pix = doc[pi].get_pixmap(matrix=fitz.Matrix(200/72, 200/72),
                                                 colorspace=fitz.csGRAY, alpha=False)
                        gray = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width).copy()
                    vendor_branch_debug('IMAGE_RENDER_BACKEND', pdf_path=str(pdf_path or ''),
                                        page=pi+1, backend='pymupdf')
                except ImportError:
                    images = convert_from_path(pdf_path, dpi=200, first_page=pi+1,
                                               last_page=pi+1, grayscale=True)
                    if not images:
                        continue
                    gray = np.asarray(images[0].convert('L'))
                    vendor_branch_debug('IMAGE_RENDER_BACKEND', pdf_path=str(pdf_path or ''),
                                        page=pi+1, backend='pdf2image')
                h, w = gray.shape
                bw_img = cv2.threshold(gray, 195, 255, cv2.THRESH_BINARY_INV)[1]
                contours, _ = cv2.findContours(bw_img, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
                boxes = []
                for c in contours:
                    x, y, bw, bh = cv2.boundingRect(c)
                    if not (9 <= bw <= 55 and 9 <= bh <= 55 and 0.7 <= bw/bh <= 1.35):
                        continue
                    if x > w*.36 or y > h*.36:
                        continue
                    if cv2.contourArea(c) < bw*bh*.43:
                        continue
                    inset = max(2, round(min(bw, bh)*.25))
                    inside = bw_img[y+inset:y+bh-inset, x+inset:x+bw-inset]
                    if inside.size < 12:
                        continue
                    fill = float(np.mean(inside > 0))
                    if fill >= .8:
                        continue
                    boxes.append((x,y,bw,bh,fill))
                # Deduplicate nested contours for the same square.
                dedup = []
                for b in sorted(boxes, key=lambda z: z[2]*z[3], reverse=True):
                    if not any(abs(b[0]-a[0]) < 7 and abs(b[1]-a[1]) < 7 for a in dedup):
                        dedup.append(b)
                # A pair of checkboxes has the same left edge and is vertically
                # separated. Reject unrelated page boxes and large gaps.
                pairs = []
                for a,b in combinations(dedup,2):
                    a,b = sorted((a,b), key=lambda z:z[1])
                    dx = abs(a[0]-b[0]); dy = b[1]-a[1]
                    if dx > max(11, min(a[2],b[2])*.6) or not (a[3]*1.5 <= dy <= h*.14):
                        continue
                    if abs(a[2]-b[2]) > max(a[2],b[2])*.45:
                        continue
                    # Prefer a tight pair with similar checkbox dimensions and
                    # near the top of the page, avoiding table grid boxes.
                    score = dy/h*3 + a[1]/h*1.5 + dx/w*5 + a[0]/w*2 + abs(a[2]-b[2])/max(a[2],b[2])
                    pairs.append((score,a,b))
                vendor_branch_debug('IMAGE_BOX_SCAN', pdf_path=str(pdf_path or ''), page=pi+1,
                                    candidate_count=len(dedup), pair_count=len(pairs),
                                    candidates=dedup[:50])
                if not pairs:
                    continue
                pairs.sort(key=lambda z:z[0])
                _, a, b = pairs[0]
                choices = [a,b]
                for (branch, label, _), box in zip(page_options, choices):
                    vendor_branch_debug('OPENCV_CHECKBOX', pdf_path=str(pdf_path or ''),
                                        page=pi+1, branch=branch, label=label,
                                        checkbox_x=box[0], checkbox_y=box[1],
                                        fill_score=round(box[4],5))
                winner = max(range(2), key=lambda i: choices[i][4])
                delta = choices[winner][4]-choices[1-winner][4]
                if choices[winner][4] >= .10 and delta >= .07:
                    result = page_options[winner][0]
                    selected.append(result)
                    vendor_branch_debug('IMAGE_SELECTED', pdf_path=str(pdf_path or ''),
                                        page=pi+1, result=result, difference=round(delta,5))
                else:
                    vendor_branch_debug('IMAGE_UNCERTAIN', pdf_path=str(pdf_path or ''),
                                        page=pi+1, scores=[a[4],b[4]])
        except Exception as exc:
            vendor_branch_debug('IMAGE_SCAN_ERROR', pdf_path=str(pdf_path or ''), error=repr(exc))
            print(f'⚠️ VendorBranch local checkbox scan failed: {exc}')

    vendor_branch_debug('SELECTION_SUMMARY', pdf_path=str(pdf_path or ''), selected=selected, all_options=[v for _, v, _, _ in options])
    unique = set(selected)
    if len(unique) == 1:
        result = next(iter(unique))
        print(f'✅ VendorBranch selected option: {result}')
        return result
    if len(unique) > 1:
        print(f'⚠️ VendorBranch conflicting selected options {sorted(unique)}')
        return ''
    distinct = {v for _, v, _, _ in options}
    if len(distinct) == 1:
        result = next(iter(distinct))
        print(f'✅ VendorBranch unique vendor header: {result}')
        return result
    print(f'⚠️ VendorBranch multiple unconfirmed vendor options: {sorted(distinct)}')
    return ''

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

def extract_amounts_from_pages(invoices, ocr_cache=None):
    """
    OCR fallback สำหรับ Amount จากผล prebuilt-invoice รอบเดิม

    รองรับ label เช่น:
        TOTAL AMOUNT 330.60
        VAT 7.00%     23.14
        NET AMOUNT   353.74

    คืนค่า: (total_amount, vat_amount, amount_inc_vat)
    """
    lines = ocr_cache["lines"] if ocr_cache is not None else get_all_lines(invoices)
    full_text = ocr_cache["full_text"] if ocr_cache is not None else "\n".join(lines)

    def clean_amount_text(text):
        return str(text or "").replace(",", "")

    def first_money(text):
        # รองรับ 330.60 / 330 / 353.74 และไม่หยิบเปอร์เซ็นต์ 7.00%
        matches = re.findall(r"(?<![\d.])(\d+(?:\.\d{1,2})?)(?!\s*%)", clean_amount_text(text))
        for value in reversed(matches):
            try:
                number = float(value)
                if number >= 0:
                    return number
            except Exception:
                pass
        return 0.0

    def find_by_label(label_patterns):
        # 1) บรรทัดเดียวกับ label
        for i, line in enumerate(lines):
            text = re.sub(r"\s+", " ", clean_amount_text(line)).strip()
            if not any(re.search(p, text, re.IGNORECASE) for p in label_patterns):
                continue

            # ตัด label ออกก่อน เพื่อไม่ให้ VAT 7.00% กลายเป็น amount
            tail = text
            for p in label_patterns:
                m = re.search(p, tail, re.IGNORECASE)
                if m:
                    tail = tail[m.end():]
                    break

            value = first_money(tail)
            if value > 0:
                return value

            # 2) Azure OCR อาจแยกค่าตัวเลขเป็นบรรทัดถัดไป
            for j in range(i + 1, min(i + 4, len(lines))):
                nearby = re.sub(r"\s+", " ", clean_amount_text(lines[j])).strip()
                # ถ้าเข้าหา label Amount ตัวถัดไป ให้หยุด เพื่อไม่หยิบผิดช่อง
                if re.search(r"(?:TOTAL\s*AMOUNT|SUB\s*TOTAL|VAT|TOTAL\s*TAX|NET\s*AMOUNT|GRAND\s*TOTAL|INVOICE\s*TOTAL)", nearby, re.IGNORECASE):
                    break
                value = first_money(nearby)
                if value > 0:
                    return value

        # 3) fallback จาก full text โดยจำกัดระยะหลัง label
        text = clean_amount_text(full_text)
        for p in label_patterns:
            m = re.search(p + r"[\s\S]{0,60}?([0-9]+(?:\.[0-9]{1,2})?)", text, re.IGNORECASE)
            if m:
                try:
                    return float(m.group(1))
                except Exception:
                    pass
        return 0.0

    total_amount = find_by_label([
        r"\bTOTAL\s*AMOUNT\b",
        r"\bSUB\s*TOTAL\b",
        r"\bTOTAL\s+BEFORE\s+VAT\b",
        r"\bAMOUNT\s+BEFORE\s+VAT\b",
        r"รวมเงิน",
        r"รวมราคาสินค้า",
    ])

    vat_amount = find_by_label([
        r"\bVAT\s*(?:7(?:\.00)?\s*%)?",
        r"\bTOTAL\s*TAX\b",
        r"ภาษีมูลค่าเพิ่ม",
    ])

    amount_inc_vat = find_by_label([
        r"\bNET\s*AMOUNT\b",
        r"\bGRAND\s*TOTAL\b",
        r"\bINVOICE\s*TOTAL\b",
        r"ยอดรวม",
    ])

    return total_amount, vat_amount, amount_inc_vat


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
    3) ถ้ายังไม่เจอ -> ค้นหาใน PO_PATTERNS ที่รวมจากไฟล์เดียวกัน
       ทุก regex มาจาก patternsPO.json เท่านั้น

    return:
        (po_value, po_branch_mismatch)
    """

    branch_code = normalize_internal_company_branch(branch_code)

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
    # 3) Search the combined configured PO patterns
    #    No built-in or legacy PO patterns are used.
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

    if value and not is_po_value(value, PO_PATTERNS):
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

    if value and not is_po_value(str(value), PO_PATTERNS):
        return value

    return ""


def find_invoice_no_from_words(invoices):
    for page in invoices.pages:
        if not hasattr(page, "words"):
            continue

        for w in page.words:
            text = w.content.strip() if w.content else ""

            if is_po_value(text, PO_PATTERNS):
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

    # เก็บข้อความ CustomerAddress ก่อน Clean ไว้ใช้หา Branch ที่พิมพ์บน Invoice
    Company_address_raw = Company_address

    # บาง template Azure อาจ map ที่อยู่ผู้ซื้อไป BillingAddress
    if not Company_address:
        Company_address = get_address_value(
            invoice.fields.get("BillingAddress")
        )
        Company_address_raw = Company_address

    # รองรับกรณี SDK/model บางเวอร์ชันคืน key เดิมที่โปรแกรมเคยใช้งาน
    if not Company_address:
        Company_address = get_address_value(
            invoice.fields.get("CompanyAddress")
        )
        Company_address_raw = Company_address

    # ถ้า Azure ไม่มี field ให้ fallback ไป OCR จาก Account To / Bill To
    if not Company_address:
        Company_address = extract_Company_address_from_pages(invoices, ocr_cache)
        Company_address_raw = Company_address

    # ==========================================================
    # Company / Customer Name, Tax ID และ Branch จาก Invoice
    # Azure field มาก่อน ถ้าไม่มีจึง OCR fallback จาก Customer zone
    # ==========================================================
    company_name = str(
        get_field_value(invoice.fields.get("CustomerName")) or ""
    ).strip()

    if not company_name:
        company_name = extract_company_name_from_pages(invoices, ocr_cache)
    
    # ==========================================================
    # Clean Customer / Company Name
    # ==========================================================
    company_name = clean_company_name(company_name)

    company_tax_id = clean_tax_id(
        get_field_value(invoice.fields.get("CustomerTaxId"))
    )

    if not re.fullmatch(r"0\d{12}", company_tax_id):
        vendor_tax_id_for_customer = clean_tax_id(
            get_field_value(invoice.fields.get("VendorTaxId"))
        )
        company_tax_id = extract_company_tax_id_from_pages(
            invoices,
            ocr_cache,
            vendor_tax_id_for_customer,
        )

    # ==========================================================
    # Customer / Company Name fallback by Customer Tax ID
    # ใช้ทั้งกรณีชื่อว่าง และ Azure/OCR คืนชื่อไม่สมบูรณ์ เช่น "THAI"
    # ==========================================================
    if company_tax_id and not is_valid_customer_company_name(company_name):
        company_name_by_taxid = extract_company_name_by_taxid(
            invoices,
            company_tax_id,
            ocr_cache,
        )
        if company_name_by_taxid:
            company_name = clean_company_name(company_name_by_taxid)

    # ==========================================================
    # CustomerAddress fallback by Customer Tax ID
    # ใช้เมื่อ Azure + Account/Bill To fallback ยังหา Address ไม่ได้
    # ไม่ค้นจาก Vendor zone และไม่ hard-code ที่อยู่ลง Excel
    # ==========================================================
    # Azure บางใบอาจ map CustomerAddress ผิดเป็น token สั้น ๆ เช่น "7"
    # ถ้าค่าไม่ใช่ address จริง ให้ถือว่าไม่มีค่าและ OCR fallback ใหม่
    if not is_valid_customer_address(Company_address) and company_tax_id:
        customer_address_by_taxid = extract_customer_address_by_taxid(
            invoices,
            company_tax_id,
            ocr_cache,
        )
        if is_valid_customer_address(customer_address_by_taxid):
            Company_address = customer_address_by_taxid
            Company_address_raw = customer_address_by_taxid
        elif not is_valid_customer_address(Company_address):
            Company_address = ""
            Company_address_raw = ""

    # อ่าน Branch จาก CustomerAddress ได้เฉพาะเมื่อมี label Branch/สาขาเท่านั้น
    # ห้ามเอาเลขเดี่ยว เช่น MOO 7 ไปเป็น CompanyBranch_invoice = 0007
    company_branch_invoice = normalize_company_branch_invoice(Company_address_raw)

    # ถ้า Azure Address ไม่มีคำว่า Branch ให้หาใน Customer zone จาก OCR
    if not company_branch_invoice:
        company_branch_invoice = extract_company_branch_invoice_from_pages(
            invoices,
            ocr_cache,
            company_tax_id,
        )

    # ==========================================================
    # Clean CompanyAddress
    # ตัดข้อความก่อนเลขที่ตั้ง เช่น
    # ชื่อลูกค้า 370... -> 370...
    # HEADOFFICE:370... -> 370...
    # ==========================================================
    Company_address = clean_company_address_start(Company_address)
    Company_address = clean_customer_address_postcode(Company_address)

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

    if is_po_value(tax_invoice_no_clean, PO_PATTERNS):
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
        "VendorAddress": address,
        "CompanyAddress": Company_address,
        "CompanyName": company_name,
        "CompanyTaxID": company_tax_id,
        "CompanyBranch": "",
        "CompanyBranch_invoice": company_branch_invoice,
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
        "VendorAddress": invoice.get("VendorAddress", ""),
        "TotalAmount": invoice.get("TotalAmount", ""),
        "VATAmount": invoice.get("VATAmount", ""),
        "AmountIncVat": invoice.get("AmountIncVat", ""),
        "PurchaseOrderNo": invoice.get("PurchaseOrderNo", ""),
        "TaxRemark": invoice.get("TaxRemark", ""),
        "Name": invoice.get("CompanyName", ""),
        "TaxID": invoice.get("CompanyTaxID", ""),
        "Branch": invoice.get("CompanyBranch", ""),
        "CompanyBranch_invoice": invoice.get("CompanyBranch_invoice", ""),
        "Address": invoice.get("CompanyAddress", ""),
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
        "VendorAddress",
        "CompanyAddress",
        "CompanyName",
        "CompanyTaxID",
        "CompanyBranch",
        "CompanyBranch_invoice",
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
            if not invoice_data.get("VendorAddress"):
                invoice_data["VendorAddress"] = extract_vendor_address_from_pages(invoices, ocr_cache)

            if not invoice_data.get("CompanyAddress"):
                customer_tax_id_for_address = clean_tax_id(
                    invoice_data.get("CompanyTaxID", "")
                )

                customer_address_fallback = extract_customer_address_by_taxid(
                    invoices,
                    customer_tax_id_for_address,
                    ocr_cache,
                )

                if not customer_address_fallback:
                    customer_address_fallback = extract_Company_address_from_pages(
                        invoices,
                        ocr_cache,
                    )

                invoice_data["CompanyAddress"] = clean_company_address_start(
                    customer_address_fallback
                )

            if not invoice_data.get("CompanyName"):
                invoice_data["CompanyName"] = extract_company_name_from_pages(
                    invoices,
                    ocr_cache,
                )

            if not invoice_data.get("CompanyTaxID"):
                invoice_data["CompanyTaxID"] = extract_company_tax_id_from_pages(
                    invoices,
                    ocr_cache,
                    invoice_data.get("VendorTaxId", ""),
                )

            if not invoice_data.get("CompanyBranch_invoice"):
                invoice_data["CompanyBranch_invoice"] = (
                    extract_company_branch_invoice_from_pages(
                        invoices,
                        ocr_cache,
                        invoice_data.get("CompanyTaxID", ""),
                    )
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

            Company_branch, branch_correct, expected_branch = validate_Company_branch(
                Company_address,
                branch_email
            )

            invoice_data["CompanyBranch"] = company_branch_to_display(Company_branch)

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

            print(f"🏢 CompanyName            : {invoice_data.get('CompanyName', '')}")
            print(f"🆔 CompanyTaxID           : {invoice_data.get('CompanyTaxID', '')}")
            print(f"📍 CompanyAddress         : {Company_address}")
            print(f"🏬 CompanyBranch          : {company_branch_to_display(Company_branch)}")
            print(
                f"📄 CompanyBranch_invoice  : "
                f"{invoice_data.get('CompanyBranch_invoice', '')}"
            )

            if not invoice_data.get("InvoiceDate") and invoice_date_ocr:
                invoice_data["InvoiceDate"] = invoice_date_ocr
                invoice_data["PostingDate"] = invoice_date_ocr

            # ==================================================
            # Amount fallback - ใช้ OCR จาก prebuilt-invoice รอบเดิม
            # ไม่ยิง Azure เพิ่ม และไม่เขียนทับค่าที่ Azure อ่านได้แล้ว
            # ==================================================
            ocr_total, ocr_vat, ocr_net = extract_amounts_from_pages(
                invoices,
                ocr_cache
            )

            if normalize_number(invoice_data.get("TotalAmount", 0)) == 0:
                if ocr_total > 0:
                    invoice_data["TotalAmount"] = round(ocr_total, 2)
                    print(f"✅ TotalAmount OCR fallback: {ocr_total:.2f}")

            if normalize_number(invoice_data.get("VATAmount", 0)) == 0:
                if ocr_vat > 0:
                    invoice_data["VATAmount"] = round(ocr_vat, 2)
                    print(f"✅ VATAmount OCR fallback: {ocr_vat:.2f}")
                else:
                    vat_fallback = extract_vat_from_pages(invoices, ocr_cache)
                    if vat_fallback:
                        invoice_data["VATAmount"] = round(vat_fallback, 2)

            if normalize_number(invoice_data.get("AmountIncVat", 0)) == 0:
                if ocr_net > 0:
                    invoice_data["AmountIncVat"] = round(ocr_net, 2)
                    print(f"✅ AmountIncVat OCR fallback: {ocr_net:.2f}")

            # Safety fallback: ถ้า SubTotal ยังหาไม่ได้ แต่ VAT + Net มีครบ
            # ให้ TotalAmount = AmountIncVat - VATAmount
            total_now = normalize_number(invoice_data.get("TotalAmount", 0))
            vat_now = normalize_number(invoice_data.get("VATAmount", 0))
            net_now = normalize_number(invoice_data.get("AmountIncVat", 0))

            # ==================================================
            # TotalAmount arithmetic validation / repair
            #
            # บาง Invoice Azure อ่าน SubTotal ผิดเป็นเลขเล็ก ๆ เช่น 1
            # ทั้งที่ VATAmount และ AmountIncVat อ่านได้ถูกต้อง
            # ตัวอย่าง:
            #   TotalAmount(Azure) = 1
            #   VATAmount          = 1161.25
            #   AmountIncVat       = 17750.53
            #   ค่าที่ถูกต้อง       = 17750.53 - 1161.25 = 16589.28
            #
            # ป้องกันการแก้เอกสารที่มี Discount/ภาษีรูปแบบอื่น:
            # จะ repair เฉพาะเมื่อค่าที่คำนวณกลับมี VAT ใกล้ 7% เท่านั้น
            # ==================================================
            if vat_now > 0 and net_now > 0 and net_now > vat_now:
                calculated_total = round(net_now - vat_now, 2)
                calculated_vat = round(calculated_total * 0.07, 2)

                vat_is_7_percent = abs(calculated_vat - vat_now) <= 0.10
                total_is_wrong = (
                    total_now <= 0
                    or abs(total_now - calculated_total) > 0.10
                )

                if calculated_total > 0 and vat_is_7_percent and total_is_wrong:
                    old_total = total_now
                    invoice_data["TotalAmount"] = calculated_total
                    total_now = calculated_total
                    print(
                        f"✅ TotalAmount repaired by VAT validation: "
                        f"{old_total:.2f} -> {calculated_total:.2f} "
                        f"({net_now:.2f} - {vat_now:.2f})"
                    )

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
                ocr_cache=ocr_cache,
                vendor_tax_id=invoice_data.get("VendorTaxId", ""),
                supplier_name=invoice_data.get("SupplierName", ""),
                pdf_path=pdf_path
            )

            vendor_branch_debug(
                'FINAL_VENDOR_BRANCH',
                pdf_path=str(pdf_path),
                invoice_no=invoice_data.get('TaxInvoiceNo', ''),
                vendor_tax_id=invoice_data.get('VendorTaxId', ''),
                supplier_name=invoice_data.get('SupplierName', ''),
                vendor_branch=invoice_data.get('VendorBranch', '')
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
            if is_po_value(tax_invoice_no, PO_PATTERNS):
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
    dest_path = os.path.join(pdf_folder, pdf_name)

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
    "VendorTaxId",
    "VendorBranch",
    "VendorAddress",
    "TotalAmount",
    "VATAmount",
    "AmountIncVat",
    "PurchaseOrderNo",
    "TaxRemark",
    "Name",
    "TaxID",
    "Branch",
    # "CompanyBranch_invoice",
    "Address",
    "Assignment",
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

    # try:
    #     total_amount = normalize_number(row.get("TotalAmount", 0))
    #     vat_amount = normalize_number(row.get("VATAmount", 0))

    #     expected_vat = round(total_amount * 0.07, 2)

    #     if abs(vat_amount - expected_vat) > 0.01:
    #         errors.append(
    #             "กรุณาเช็ค Amount ทั้ง 3 ช่อง"
    #         )

    # VAT validation block is disabled above; no except clause is needed.

    row["Emessage"] = append_msg(
        row.get("Emessage", ""),
        " | ".join(errors)
    )

excel_rows = [build_excel_row(row) for row in all_data]

df = pd.DataFrame(excel_rows)
df = df.reindex(columns=EXCEL_COLUMNS)

df.to_excel(output_excel, index=False)

# ==========================================================
# ==========================================================
# 🔎 Validate OCR Excel with Matching_TextFile
# ==========================================================

# เลือกไฟล์ TXT ที่แก้ไขล่าสุดจาก Input Folder ของสาขาที่เลือก
matching_txt_files = [
    os.path.join(MATCHING_TEXT_FOLDER, filename)
    for filename in os.listdir(MATCHING_TEXT_FOLDER)
    if filename.lower().endswith(".txt")
    and os.path.isfile(os.path.join(MATCHING_TEXT_FOLDER, filename))
]

if not matching_txt_files:
    raise FileNotFoundError(
        f"ไม่พบไฟล์ .txt ในโฟลเดอร์: {MATCHING_TEXT_FOLDER}"
    )

MATCHING_TEXT_PATH = max(
    matching_txt_files,
    key=os.path.getmtime
)

if not os.path.isfile(MATCHING_SCRIPT):
    raise FileNotFoundError(
        f"ไม่พบ Matching_TextFile.py: {MATCHING_SCRIPT}"
    )

if not os.path.isfile(MATCHING_MASTER_PATH):
    raise FileNotFoundError(
        f"ไม่พบ Master Supplier: {MATCHING_MASTER_PATH}"
    )

# เก็บไฟล์ผลลัพธ์ของรอบนี้ไว้ในโฟลเดอร์เดียวกัน
os.makedirs(MATCHING_OUTPUT_FOLDER, exist_ok=True)

# คัดลอกไฟล์ OCR Excel เข้าโฟลเดอร์รอบปัจจุบัน
MATCHING_INPUT_EXCEL = os.path.join(
    MATCHING_OUTPUT_FOLDER,
    os.path.basename(output_excel)
)
if os.path.abspath(output_excel) != os.path.abspath(MATCHING_INPUT_EXCEL):
    shutil.copy2(output_excel, MATCHING_INPUT_EXCEL)

# กำหนดปลายทาง TXT แต่ยังไม่ย้ายจนกว่า Matching จะสำเร็จ
MATCHING_ARCHIVE_TEXT = os.path.join(
    MATCHING_OUTPUT_FOLDER,
    os.path.basename(MATCHING_TEXT_PATH)
)

MATCHING_OUTPUT_EXCEL = os.path.join(
    MATCHING_OUTPUT_FOLDER,
    os.path.splitext(os.path.basename(output_excel))[0]
    + "_matched.xlsx"
)

print("\n🔎 Starting Matching_TextFile validation...")
print(f"Excel input : {MATCHING_INPUT_EXCEL}")
print(f"TXT input   : {MATCHING_TEXT_PATH}")
print(f"Master      : {MATCHING_MASTER_PATH}")
print(f"Output      : {MATCHING_OUTPUT_EXCEL}")

# ส่ง path ทั้ง 4 ค่าให้ Matching_TextFile ทำงานต่อ
subprocess.run(
    [
        sys.executable,
        MATCHING_SCRIPT,
        MATCHING_INPUT_EXCEL,
        MATCHING_TEXT_PATH,
        MATCHING_MASTER_PATH,
        MATCHING_OUTPUT_EXCEL,
    ],
    check=True
)

if not os.path.isfile(MATCHING_OUTPUT_EXCEL):
    raise FileNotFoundError(
        f"Matching did not create output Excel: {MATCHING_OUTPUT_EXCEL}"
    )

# Matching_TextFile.py รุ่นใหม่ย้าย TXT ให้แล้ว; รองรับรุ่นเก่าที่ยังไม่ย้ายด้วย
if os.path.normcase(os.path.abspath(MATCHING_TEXT_PATH)) != os.path.normcase(os.path.abspath(MATCHING_ARCHIVE_TEXT)):
    if os.path.isfile(MATCHING_TEXT_PATH):
        if os.path.exists(MATCHING_ARCHIVE_TEXT):
            raise FileExistsError(
                f"ไม่ย้าย TXT เพราะมีไฟล์ชื่อเดียวกันอยู่ที่ปลายทางแล้ว: {MATCHING_ARCHIVE_TEXT}"
            )
        shutil.move(MATCHING_TEXT_PATH, MATCHING_ARCHIVE_TEXT)
        print(f"📦 Moved TXT: {MATCHING_TEXT_PATH} -> {MATCHING_ARCHIVE_TEXT}")
    elif os.path.isfile(MATCHING_ARCHIVE_TEXT):
        print(f"📦 TXT ถูกย้ายโดย Matching_TextFile.py แล้ว: {MATCHING_ARCHIVE_TEXT}")
    else:
        raise FileNotFoundError(f"ไม่พบ TXT ทั้งต้นทางและปลายทาง: {MATCHING_TEXT_PATH}")

print(f"\n✅ Matching completed. Final Excel: {MATCHING_OUTPUT_EXCEL}")

# Matching_TextFile.py ทำสี Emessage ในไฟล์ matched แล้ว (เวอร์ชัน fixed)
