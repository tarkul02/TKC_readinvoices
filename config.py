import os
from dotenv import load_dotenv


# ====================================================
# 📄 โหลด .env
# ====================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

ENV_FILE = os.path.join(BASE_DIR, ".env")

load_dotenv(ENV_FILE)


# ====================================================
# 🔐 Azure Document Intelligence
# ====================================================

AZURE_ENDPOINT = os.getenv("AZURE_ENDPOINT")
AZURE_KEY = os.getenv("AZURE_KEY")


# ====================================================
# 📁 Main Path
# ====================================================

MAIN_PATH = os.getenv("MAIN_PATH")


# ====================================================
# 🔎 ตรวจสอบ Config
# ====================================================

if not AZURE_ENDPOINT:
    raise ValueError("ไม่พบ AZURE_ENDPOINT ในไฟล์ .env")

if not AZURE_KEY:
    raise ValueError("ไม่พบ AZURE_KEY ในไฟล์ .env")

if not MAIN_PATH:
    raise ValueError("ไม่พบ MAIN_PATH ในไฟล์ .env")


# ====================================================
# 📁 สร้าง Path แยกตาม Process
# ====================================================

def get_paths(today_str, branch_email, process_name):

    # -----------------------------
    # INPUT ของแต่ละ Process
    # -----------------------------
    input_folder = os.path.join(
        MAIN_PATH,
        "INPUT",
        process_name,
        branch_email
    )

    # -----------------------------
    # Temp ของแต่ละ Process
    # -----------------------------
    temp_folder = os.path.join(
        MAIN_PATH,
        "TempSplit",
        process_name
    )

    # -----------------------------
    # Excel Output
    # -----------------------------
    output_excel = os.path.join(
        MAIN_PATH,
        "INPUT",
        process_name,
        branch_email,
        "OutputExcel",
        today_str,
        "Excel",
        "OCR",
        "All_Invoices.xlsx"
    )

    # -----------------------------
    # Destination
    # -----------------------------
    dest_folder = os.path.join(
        MAIN_PATH,
        "INPUT",
        process_name,
        branch_email,
        "OutputExcel",
        today_str
    )

    return {
        "input_folder": input_folder,
        "temp_folder": temp_folder,
        "output_excel": output_excel,
        "dest_folder": dest_folder,
    }