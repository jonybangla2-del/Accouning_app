import os
import json
import secrets
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
SMTP_CONFIG_FILE = os.path.join(BASE_DIR, 'smtp_config.json')
ADMIN_EMAIL = "jonybangla2@gmail.com"

DEFAULT_SMTP_CONFIG = {
    "enabled": True,
    "smtp_server": "smtp.gmail.com",
    "smtp_port": 587,
    "smtp_user": "jonybangla2@gmail.com",
    "smtp_password": "",
    "use_tls": True,
    "sender_name": "ZR-ERP Security System"
}

def get_smtp_config():
    if not os.path.exists(SMTP_CONFIG_FILE):
        save_smtp_config(DEFAULT_SMTP_CONFIG)
        return DEFAULT_SMTP_CONFIG.copy()
    try:
        with open(SMTP_CONFIG_FILE, 'r', encoding='utf-8') as f:
            cfg = json.load(f)
            return cfg
    except Exception as e:
        print(f"Error reading smtp_config.json: {e}")
        return DEFAULT_SMTP_CONFIG.copy()

def save_smtp_config(cfg):
    try:
        with open(SMTP_CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, indent=2)
        return True
    except Exception as e:
        print(f"Error saving smtp_config.json: {e}")
        return False

def generate_otp():
    """Generates a secure 6-digit numeric OTP."""
    return f"{secrets.randbelow(900000) + 100000}"

def send_otp_email(to_email=ADMIN_EMAIL, otp_code=None):
    """
    Dispatches OTP verification code to the target email.
    Attempts actual SMTP dispatch if configured, and always returns
    delivery status and diagnostics for transparent verification.
    """
    if not otp_code:
        otp_code = generate_otp()

    cfg = get_smtp_config()
    smtp_server = os.environ.get('SMTP_SERVER') or cfg.get('smtp_server', 'smtp.gmail.com')
    smtp_port = int(os.environ.get('SMTP_PORT') or cfg.get('smtp_port', 587))
    smtp_user = os.environ.get('SMTP_USER') or cfg.get('smtp_user', ADMIN_EMAIL)
    smtp_password = os.environ.get('SMTP_PASSWORD') or cfg.get('smtp_password', '')
    sender_name = cfg.get('sender_name', 'ZR-ERP Security System')
    use_tls = cfg.get('use_tls', True)

    subject = f"ZR-ERP Login Verification OTP: {otp_code}"
    
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: #f8fafc; color: #1e293b; margin: 0; padding: 20px; }}
            .container {{ max-width: 500px; margin: 0 auto; background: #ffffff; border-radius: 10px; border: 1px solid #e2e8f0; overflow: hidden; box-shadow: 0 4px 12px rgba(0,0,0,0.05); }}
            .header {{ background: #1e3a8a; color: #ffffff; padding: 24px; text-align: center; }}
            .brand {{ font-size: 24px; font-weight: 800; letter-spacing: 0.05em; }}
            .sub {{ font-size: 13px; opacity: 0.85; margin-top: 4px; }}
            .content {{ padding: 30px 24px; text-align: center; }}
            .otp-box {{ background: #f1f5f9; border: 2px dashed #3b82f6; border-radius: 8px; padding: 18px; margin: 24px auto; display: inline-block; }}
            .otp-code {{ font-family: 'Courier New', Courier, monospace; font-size: 36px; font-weight: 800; color: #1e40af; letter-spacing: 8px; }}
            .warning {{ font-size: 13px; color: #64748b; line-height: 1.6; margin-top: 15px; }}
            .footer {{ background: #f8fafc; padding: 15px; text-align: center; font-size: 11px; color: #94a3b8; border-top: 1px solid #e2e8f0; }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <div class="brand">⚖️ ZR-ERP Enterprise</div>
                <div class="sub">Two-Factor Authentication Security Alert</div>
            </div>
            <div class="content">
                <h3 style="margin-top: 0; color: #0f172a;">Administrator Sign-In Verification</h3>
                <p style="font-size: 14px; color: #475569;">
                    A login attempt was initiated for your Administrator account in <strong>ZR-ERP</strong>. Use the One-Time Password (OTP) below to authenticate your session:
                </p>
                <div class="otp-box">
                    <div class="otp-code">{otp_code}</div>
                </div>
                <div class="warning">
                    ⏱️ This OTP is valid for <strong>10 minutes</strong>.<br>
                    🛡️ Never share this verification code with anyone.
                </div>
            </div>
            <div class="footer">
                Dispatched to {to_email} &bull; ZR-ERP Automated Security Dispatcher &bull; {datetime.utcnow().strftime('%d-%b-%Y %H:%M:%S UTC')}
            </div>
        </div>
    </body>
    </html>
    """

    text_content = f"""
ZR-ERP Enterprise - Login Verification OTP

Your 6-digit Administrator OTP is: {otp_code}

This code is valid for 10 minutes. Dispatched to {to_email}.
If you did not request this login, please change your password immediately.
    """

    # If SMTP password is provided, attempt actual SMTP dispatch
    smtp_dispatched = False
    delivery_error = None

    if smtp_password and smtp_user:
        try:
            msg = MIMEMultipart("alternative")
            msg["Subject"] = subject
            msg["From"] = f"{sender_name} <{smtp_user}>"
            msg["To"] = to_email

            part1 = MIMEText(text_content, "plain")
            part2 = MIMEText(html_content, "html")
            msg.attach(part1)
            msg.attach(part2)

            context = ssl.create_default_context()
            if smtp_port == 465:
                with smtplib.SMTP_SSL(smtp_server, smtp_port, context=context, timeout=15) as server:
                    server.login(smtp_user, smtp_password)
                    server.sendmail(smtp_user, to_email, msg.as_string())
            else:
                with smtplib.SMTP(smtp_server, smtp_port, timeout=15) as server:
                    if use_tls:
                        server.starttls(context=context)
                    server.login(smtp_user, smtp_password)
                    server.sendmail(smtp_user, to_email, msg.as_string())
            smtp_dispatched = True
        except Exception as e:
            delivery_error = str(e)
            print(f"[SMTP Error] Failed to send email to {to_email}: {e}")

    return {
        'success': True,
        'smtp_dispatched': smtp_dispatched,
        'otp': otp_code,
        'email': to_email,
        'error': delivery_error,
        'timestamp': datetime.utcnow().isoformat()
    }
