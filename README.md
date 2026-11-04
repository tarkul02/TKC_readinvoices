# TKC Read Invoices

โปรเจกต์สำหรับอ่านและประมวลผล Invoice โดยใช้ Azure AI Document Intelligence และส่งออกข้อมูลสำหรับใช้งานต่อในรูปแบบ Excel

## Requirements

- Python 3.x
- Azure AI Document Intelligence และ Azure endpoint/key ที่ใช้งานได้
- Poppler (สำหรับขั้นตอนที่แปลง PDF เป็นรูปภาพ)

## 1. Clone Repository

```bash
git clone <repository-url>
cd TKC_readinvoices
```

> แนะนำให้เก็บ Source Code ไว้ที่ `C:\TKC\TKC_readinvoices` และเก็บข้อมูลที่ประมวลผลไว้คนละตำแหน่ง เช่น `C:\TKC\Document` ตามค่า `MAIN_PATH`

## 2. Install Python Modules

```bash
python -m pip install -r requirements.txt
```

## 3. Create `.env`

สร้างไฟล์ `.env` ที่ root ของโปรเจกต์ (โฟลเดอร์เดียวกับ `config.py`)

```text
TKC_readinvoices/
├── .env
├── .gitignore
├── config.py
├── requirements.txt
├── main.py
└── ...
```

ตัวอย่าง `.env`:

```env
AZURE_ENDPOINT=https://your-resource.cognitiveservices.azure.com/
AZURE_KEY=YOUR_AZURE_KEY
MAIN_PATH=C:\TKC\Document
```

- `AZURE_ENDPOINT`: Endpoint ของ Azure AI Document Intelligence
- `AZURE_KEY`: Key สำหรับเรียกใช้ Azure AI Document Intelligence
- `MAIN_PATH`: Root folder สำหรับเก็บ Input, Output, Temp และ Master Supplier; เปลี่ยนได้ตามเครื่อง
- ไม่ต้องสร้าง `.env` ภายใน `MAIN_PATH`; ให้เก็บไว้ข้าง `config.py`
- ตัวอย่างข้างต้นเป็น placeholder ต้องแทนด้วยค่าจริงก่อนใช้งาน

**Security:** ห้าม Commit/Push `.env` หรือ Azure Key ขึ้น GitHub และหากเคยเผยแพร่ Key ต้อง Rotate/Regenerate ทันที

ตรวจสอบ `.gitignore`:

```gitignore
.env
.env.*
!.env.example
INPUT/
TempSplit/
DebugVendorBranch/
```

> กฎ `.gitignore` ป้องกันไฟล์ที่ยังไม่ถูก Git ติดตามเท่านั้น หากเคย Commit ไฟล์ความลับไปแล้ว ต้องจัดการประวัติและเปลี่ยน Key ด้วย

## 4. Project Configuration (`config.py`)

`config.py` จะโหลด `.env` จากโฟลเดอร์ Source Code ผ่าน `python-dotenv` และตรวจสอบว่ามี `AZURE_ENDPOINT`, `AZURE_KEY`, `MAIN_PATH` ครบถ้วน หากไม่มีจะหยุดด้วย `ValueError`

ไฟล์ `config.py` เวอร์ชันนี้มีฟังก์ชัน:

```python
def get_paths(today_str, branch_email, process_name):
    ...
```

เมื่อเรียก `get_paths()` โปรแกรมจะ **สร้างโฟลเดอร์ที่จำเป็นอัตโนมัติ** ภายใต้ `MAIN_PATH` ด้วย `os.makedirs(..., exist_ok=True)` และส่ง Dictionary ของ Path กลับมา โดยยังไม่สร้างไฟล์ Excel, PDF, TXT หรือ Master Supplier จริง

### 4.1 เตรียมโฟลเดอร์สำหรับ Process / Branch

ตัวอย่างบน PowerShell (รันจากโฟลเดอร์โปรเจกต์):

```powershell
python .\config.py 20261104_13 00000 Process3
```

พารามิเตอร์เรียงตามลำดับ:

1. `today_str`: ชื่อรอบหรือโฟลเดอร์ Output เช่น `20261104_13`
2. `branch_email`: รหัสสาขา เช่น `00000` หรือ `1100`
3. `process_name`: ชื่อ Process เช่น `Process1`, `Process2`, `Process3`

ตัวอย่างโครงสร้างที่สร้างเมื่อ `MAIN_PATH=C:\TKC\Document`:

```text
C:\TKC\Document\
├── INPUT\
│   └── Process3\
│       └── 00000\
│           └── OutputExcel\
│               └── 20261104_13\
│                   └── Excel\
│                       └── OCR\
├── TempSplit\
│   └── Process3\
└── MasterSuplier\
```

`get_paths()` สร้างเฉพาะ Process/Branch ที่ส่งเข้ามา ไม่ได้สร้างทุก Process/Branch ล่วงหน้า

### 4.2 Path ที่คืนจาก `get_paths()`

| Key | ตัวอย่าง Path / ความหมาย |
|---|---|
| `input_folder` | `C:\TKC\Document\INPUT\Process3\00000` — ตำแหน่งรับไฟล์ |
| `temp_folder` | `C:\TKC\Document\TempSplit\Process3` — ไฟล์ชั่วคราว |
| `output_excel` | `C:\TKC\Document\INPUT\Process3\00000\OutputExcel\20261104_13\Excel\OCR\All_Invoices.xlsx` |
| `dest_folder` | `C:\TKC\Document\INPUT\Process3\00000\OutputExcel\20261104_13` |
| `report_txt_folder` | โฟลเดอร์เดียวกับ `input_folder` |
| `master_supplier_path` | `C:\TKC\Document\MasterSuplier\Master Vendor Name (MAT)_TH_ENG.xlsx` |
| `matching_script_path` | ไฟล์ `Matching_TextFile.py` ในโฟลเดอร์ Source Code ข้าง `config.py` |

**สิ่งที่ต้องเตรียมเอง:** นำไฟล์ `Master Vendor Name (MAT)_TH_ENG.xlsx` ไปไว้ใน `MasterSuplier` และนำ PDF/TXT ไปไว้ใน Input Folder ที่ตรงกับ Process/Branch ที่จะรัน (ตามข้อกำหนดของสคริปต์ประมวลผล)

### 4.3 เรียกใช้จาก Python Script

```python
from config import AZURE_ENDPOINT, AZURE_KEY, get_paths

paths = get_paths(
    today_str="20261104_13",
    branch_email="00000",
    process_name="Process3",
)

print(paths["input_folder"])
print(paths["output_excel"])
```

**สำคัญ:** `get_paths()` จะสร้าง Directory เมื่อถูกเรียกใช้งาน ไม่ใช่สร้างทุก Folder ทันทีที่ Import `config.py` ส่วน `python config.py ...` เป็นวิธีสร้างโฟลเดอร์ล่วงหน้าจาก Command Line

### 4.4 เปลี่ยนเครื่องหรือเปลี่ยนตำแหน่งเก็บข้อมูล

แก้เพียง `MAIN_PATH` ใน `.env` เช่น `MAIN_PATH=D:\InvoiceData` จากนั้นเรียก `get_paths()` อีกครั้ง โครงสร้างใหม่จะถูกสร้างภายใต้ Path ใหม่ โดยไฟล์เก่าจะ **ไม่ถูกย้ายอัตโนมัติ** ต้องคัดลอกข้อมูลที่ต้องการเอง

## 5. Run Program

รูปแบบการรันสคริปต์ประมวลผลเดิม (ขึ้นกับสคริปต์แต่ละตัว):

```bash
python <script_name.py> <today_str> <branch_email>
```

ตัวอย่าง:

```powershell
python .\03_430000848.py testfile_444454 1100
```

- `testfile_444454`: ชื่อ Folder/Run Parameter
- `1100`: Branch Email / Branch Code

> คำสั่งสร้างโฟลเดอร์ `config.py` รับ **3 พารามิเตอร์** รวม `process_name` ส่วนสคริปต์ OCR ตัวอย่างข้างต้นรับ **2 พารามิเตอร์** ตามรูปแบบเดิม ต้องตรวจสอบว่าสคริปต์ OCR เรียก `get_paths()` โดยระบุ Process Name ถูกต้อง

## 6. Main Project Structure

```text
C:\TKC\TKC_readinvoices\
├── .env                       # ความลับ: ไม่ขึ้น Git
├── .gitignore
├── config.py                  # โหลด .env + สร้าง Folder
├── main.py
├── 03_430000848.py
├── 07_430000840.py
├── Matching_TextFile.py
├── patternsInvoice.json
├── requirements.txt
├── README.md
└── poppler-24.08.0\         # หากติดตั้งไว้ในโปรเจกต์

C:\TKC\Document\             # MAIN_PATH: เก็บข้อมูลแยกจาก Source Code
├── INPUT\
├── TempSplit\
└── MasterSuplier\
```

## 7. Update Dependencies

เมื่อเพิ่ม Python Module ใหม่ ให้เพิ่มใน `requirements.txt` แล้วติดตั้งอีกครั้ง:

```bash
python -m pip install -r requirements.txt
```

`config.py` ใช้แพ็กเกจ `python-dotenv` ดังนั้น `requirements.txt` ต้องมีแพ็กเกจนี้ด้วย

## 8. Troubleshooting

- **`ไม่พบ MAIN_PATH ในไฟล์ .env`**: ตรวจว่ามี `.env` ข้าง `config.py` และกำหนด `MAIN_PATH` แล้ว
- **`ModuleNotFoundError: No module named 'dotenv'`**: ติดตั้ง `python-dotenv` หรือรันคำสั่งติดตั้ง `requirements.txt`
- **`PermissionError` ระหว่างสร้าง Folder**: ตรวจสิทธิ์เขียนใน `MAIN_PATH`
- **ไม่พบ Master Supplier**: `config.py` สร้างโฟลเดอร์ `MasterSuplier` แต่ไม่ได้สร้างไฟล์ Master ต้องนำไฟล์จริงมาวาง
- **หา `Matching_TextFile.py` ไม่พบ**: ต้องมีไฟล์นี้อยู่ข้าง `config.py` ใน Source Code
- **Input/Output ยังชี้ Path เก่า**: ตรวจว่าสคริปต์อื่นไม่ได้ Hardcode Path และใช้ค่าที่คืนจาก `get_paths()` จริง

## Security

- ห้าม Push `.env` ขึ้น GitHub
- ห้ามเขียน `AZURE_KEY` ลงใน Source Code
- หาก Azure Key เคยถูก Commit หรือเผยแพร่ ให้ Rotate/Regenerate Key ทันที
- เก็บเฉพาะตัวอย่าง Configuration เช่น `.env.example` ใน Git
- ไม่ควร Commit ไฟล์ Input, Output หรือ Debug ที่อาจมีข้อมูล Invoice จริง
