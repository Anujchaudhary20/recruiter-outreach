import os  # ← ADD THIS LINE
import json
import requests
from datetime import datetime, timedelta
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

GMAIL_EMAIL = os.getenv('ANUJ_EMAIL', 'Anuj04004@gmail.com')
GMAIL_APP_PASSWORD = os.getenv('GMAIL_APP_PASSWORD')

def send_email_with_resume(email_data, resume):
    """Send email via Gmail SMTP"""
    try:
        msg = MIMEMultipart()
        msg['From'] = GMAIL_EMAIL
        msg['To'] = email_data['to']
        msg['Subject'] = email_data['subject']
        msg.attach(MIMEText(email_data['body'], 'plain'))
        
        # Connect and send
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(GMAIL_EMAIL, GMAIL_APP_PASSWORD)
            server.send_message(msg)
        
        print(f"✅ Email sent to {email_data['to']}")
        return True
    except Exception as e:
        print(f"❌ Failed to send to {email_data['to']}: {e}")
        return False
