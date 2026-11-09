import os
import smtplib
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from dotenv import load_dotenv

# Load .env from the same folder as this script.
load_dotenv(Path(__file__).resolve().parent / '.env')


def send_mail():
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    from_email = 'tarkulbeer@gmail.com'
    to_email = 'Beerbeerlovemusic@gmail.com'
    app_password = os.getenv('GMAIL_APP_PASSWORD', '').replace(' ', '').strip()

    if not app_password:
        raise ValueError('ไม่พบ GMAIL_APP_PASSWORD ในไฟล์ .env หรือ Environment Variable')

    message = MIMEMultipart()
    message['From'] = from_email
    message['To'] = to_email
    message['Subject'] = 'Test email!!222333'
    message.attach(MIMEText('<h1>Hi</h1>', 'html', 'utf-8'))

    try:
        with smtplib.SMTP(smtp_server, smtp_port, timeout=30) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(from_email, app_password)
            server.send_message(message)
        print('ส่ง Email สำเร็จ')
    except Exception as exc:
        print(f'ส่ง Email ไม่สำเร็จ: {exc}')
        raise


if __name__ == '__main__':
    send_mail()
