"""TKC invoice processing configuration and directory initialization."""
import os
from pathlib import Path
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / '.env')

AZURE_ENDPOINT = os.getenv('AZURE_ENDPOINT')
AZURE_KEY = os.getenv('AZURE_KEY')
MAIN_PATH = os.getenv('MAIN_PATH')

if not AZURE_ENDPOINT:
    raise ValueError('ไม่พบ AZURE_ENDPOINT ในไฟล์ .env')
if not AZURE_KEY:
    raise ValueError('ไม่พบ AZURE_KEY ในไฟล์ .env')
if not MAIN_PATH:
    raise ValueError('ไม่พบ MAIN_PATH ในไฟล์ .env')

MAIN_PATH = os.path.normpath(MAIN_PATH)


def get_paths(today_str, branch_email, process_name):
    """Return paths and create directories under MAIN_PATH as needed."""
    for label, value in [('today_str', today_str), ('branch_email', branch_email), ('process_name', process_name)]:
        value = str(value)
        if not value or value in ('.', '..') or '/' in value or '\\' in value:
            raise ValueError(f'Invalid {label}: {value!r}')

    input_folder = os.path.join(MAIN_PATH, 'INPUT', process_name, branch_email)
    temp_folder = os.path.join(MAIN_PATH, 'TempSplit', process_name)
    dest_folder = os.path.join(input_folder, 'OutputExcel', today_str)
    output_excel = os.path.join(dest_folder, 'Excel', 'OCR', 'All_Invoices.xlsx')
    master_supplier_folder = os.path.join(MAIN_PATH, 'MasterSuplier')
    master_supplier_path = os.path.join(master_supplier_folder, 'Master Vendor Name (MAT)_TH_ENG.xlsx')
    matching_script_path = str(BASE_DIR / 'Matching_TextFile.py')

    for folder in (input_folder, temp_folder, os.path.dirname(output_excel), master_supplier_folder):
        os.makedirs(folder, exist_ok=True)

    return {
        'input_folder': input_folder,
        'temp_folder': temp_folder,
        'output_excel': output_excel,
        'dest_folder': dest_folder,
        'report_txt_folder': input_folder,
        'master_supplier_path': master_supplier_path,
        'matching_script_path': matching_script_path,
    }


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='Create TKC invoice processing folders')
    parser.add_argument('today_str', help='Output folder name, e.g. 20261104_13')
    parser.add_argument('branch_email', help='Branch code, e.g. 00000')
    parser.add_argument('process_name', help='Process1, Process2 or Process3')
    args = parser.parse_args()
    for name, path in get_paths(args.today_str, args.branch_email, args.process_name).items():
        print(f'{name}: {path}')
