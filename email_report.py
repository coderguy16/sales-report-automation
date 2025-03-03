import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
from dotenv import load_dotenv

load_dotenv()

def send_email(report_file):
    msg = MIMEMultipart()
    msg['From'] = os.getenv('EMAIL_USER')
    msg['To'] = os.getenv('CLIENT_EMAIL')
    msg['Subject'] = 'Monthly Sales Report - Automated'
    
    # Attach report
    with open(report_file, 'rb') as f:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename="{report_file}"')
    msg.attach(part)
    
    # Connect to server (Gmail example)
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(os.getenv('EMAIL_USER'), os.getenv('EMAIL_PASSWORD'))
        server.send_message(msg)
