"""Send the completed invoice workbook with an Emessage-based summary."""
import os
import ssl
import smtplib
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path
from openpyxl import load_workbook
from dotenv import load_dotenv

load_dotenv(Path(__file__).resolve().parent / '.env')


def summarize_excel(excel_path):
    """Count nonempty invoice rows; Emessage is the source of truth."""
    wb = load_workbook(excel_path, read_only=True, data_only=True)
    try:
        ws = wb.active
        headers = {str(c.value).strip().casefold(): i for i, c in enumerate(ws[1]) if c.value is not None}
        if 'emessage' not in headers:
            raise ValueError('ไม่พบคอลัมน์ Emessage ในไฟล์ Excel')
        emsg_index = headers['emessage']
        total = valid = invalid = warning = 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not any(v is not None and str(v).strip() for v in row):
                continue
            total += 1
            msg = str(row[emsg_index] or '').strip()
            if not msg:
                valid += 1
            elif msg == 'เอกสารใบนี้เป็นสำเนา':
                warning += 1
            else:
                invalid += 1
        return {'total': total, 'valid': valid, 'invalid': invalid, 'warning': warning}
    finally:
        wb.close()


def send_invoice_email(excel_path):
    excel_path = Path(excel_path).resolve()
    if not excel_path.is_file():
        raise FileNotFoundError(f'ไม่พบไฟล์ Excel: {excel_path}')
    counts = summarize_excel(excel_path)
    sender = os.getenv('GMAIL_FROM_EMAIL', 'tarkulbeer@gmail.com').strip()
    recipients = [s.strip() for s in os.getenv('GMAIL_TO_EMAIL', '').replace(';', ',').split(',') if s.strip()]
    password = os.getenv('GMAIL_APP_PASSWORD', '').replace(' ', '').strip()
    if not recipients:
        raise ValueError('กรุณาตั้งค่า GMAIL_TO_EMAIL ในไฟล์ .env')
    if not password:
        raise ValueError('กรุณาตั้งค่า GMAIL_APP_PASSWORD ในไฟล์ .env')

    msg = EmailMessage()
    msg['From'] = sender
    msg['To'] = ', '.join(recipients)
    msg['Subject'] = f'Invoice Validation Result - {datetime.now():%Y-%m-%d %H:%M}'
    msg.set_content(
        'สรุปผลการตรวจสอบ Invoice\n\n'
        f'จำนวนข้อมูลทั้งหมด: {counts["total"]} ชุด\n'
        f'ข้อมูลถูกต้อง: {counts["valid"]} ชุด\n'
        f'ข้อมูลไม่ถูกต้อง: {counts["invalid"]} ชุด\n'
        f'คำเตือนเอกสารสำเนาอย่างเดียว: {counts["warning"]} ชุด\n\n'
        'หมายเหตุ: นับจากคอลัมน์ Emessage ใน Excel ที่แนบมา\n'
        f'ไฟล์แนบ: {excel_path.name}\n'
    )
    msg.add_attachment(excel_path.read_bytes(), maintype='application',
                       subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                       filename=excel_path.name)

    host = os.getenv('SMTP_HOST', 'smtp.gmail.com').strip()
    port = int(os.getenv('SMTP_PORT', '465'))
    context = ssl.create_default_context()
    if port == 465:
        with smtplib.SMTP_SSL(host, port, context=context, timeout=30) as server:
            server.login(sender, password)
            server.send_message(msg)
    else:
        with smtplib.SMTP(host, port, timeout=30) as server:
            server.ehlo()
            server.starttls(context=context)
            server.ehlo()
            server.login(sender, password)
            server.send_message(msg)
    print(f'📧 ส่ง Email สำเร็จ: {", ".join(recipients)} | {counts}')
    return counts
