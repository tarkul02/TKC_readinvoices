# TKC Read Invoices

โปรเจกต์สำหรับอ่านและประมวลผล Invoice โดยใช้ Azure AI Document Intelligence และส่งออกข้อมูลสำหรับใช้งานต่อในรูปแบบ Excel

## Requirements

* Python 3.x
* Azure AI Document Intelligence
* Poppler

## 1. Clone Repository

```bash
git clone <repository-url>
cd TKC_readinvoices
```

## 2. Install Python Modules

ติดตั้ง Python modules ที่จำเป็นจาก `requirements.txt`

```bash
python -m pip install -r requirements.txt
```

## 3. Create `.env`

สร้างไฟล์ชื่อ `.env` ไว้ที่ root ของ project

```text
TKC_readinvoices/
├── .env
├── .gitignore
├── config.py
├── requirements.txt
├── main.py
└── ...
```

เพิ่มค่าต่อไปนี้ลงใน `.env`

```env
AZURE_ENDPOINT=https://your-resource.cognitiveservices.azure.com/
AZURE_KEY=YOUR_AZURE_KEY

MAIN_PATH=C:\TKC\TKC_readinvoices
```

> **Security:** ห้าม Commit หรือ Push ไฟล์ `.env` ขึ้น GitHub เนื่องจากมี Azure Key และข้อมูล configuration ที่เป็นความลับ

ไฟล์ `.gitignore` ต้องมี:

```gitignore
.env
.env.*
!.env.example
```

## 4. Project Configuration

โปรแกรมจะโหลดค่าจาก `.env` ผ่าน `config.py`

ตัวอย่าง:

```python
from config import (
    AZURE_ENDPOINT,
    AZURE_KEY,
    get_paths
)
```

ไม่ควรเขียน Azure Key ลงใน Python source code โดยตรง

## 5. Run Program

รูปแบบคำสั่ง:

```bash
python <script_name.py> <today_str> <branch_email>
```

ตัวอย่าง:

```bash
python .\03_430000848.py testfile_444454 1100
```

โดย:

* `20260814` = ชื่อ folder/date parameter ที่ส่งให้โปรแกรม
* `1100` = Branch Email / Branch Code

## 6. Main Project Structure

```text
TKC_readinvoices/
├── .env                    # Environment variables (ไม่ขึ้น Git)
├── .gitignore
├── config.py               # Configuration
├── main.py
├── 03_430000848.py
├── 07_430000840.py
├── patternsInvoice.json    # Invoice patterns
├── requirements.txt        # Python dependencies
├── README.md
│
├── INPUT/                  # Input files (ไม่ขึ้น Git)
├── TempSplit/              # Temporary files (ไม่ขึ้น Git)
└── poppler-24.08.0/        # Poppler local installation
```

## 7. Update Dependencies

หากมีการเพิ่ม Python module ใหม่ ให้เพิ่ม module ลงใน `requirements.txt`

ติดตั้งทั้งหมดอีกครั้งด้วย:

```bash
python -m pip install -r requirements.txt
```

## Security

* ห้าม Push `.env` ขึ้น GitHub
* ห้ามเขียน `AZURE_KEY` ลงใน source code
* หาก Azure Key เคยถูก Commit หรือเผยแพร่ ให้ Rotate/Regenerate Key ทันที
* เก็บเฉพาะตัวอย่าง configuration เช่น `.env.example` ไว้ใน Git
