import smtplib
import pandas as pd
from email.message import EmailMessage
from pathlib import Path
import sys

# Optional: Set UTF-8 encoding (works in some Windows setups)
# sys.stdout.reconfigure(encoding='utf-8')

# === CONFIGURATION ===
EMAIL_ADDRESS = 'kartikmahamuni09@gmail.com'         # 🔁 Your Gmail
APP_PASSWORD = 'utvx fsnp kmep hfuy'                 # 🔁 16-digit app password
EXCEL_FILE = 'recipients.xlsx'                       # Excel file with Name + Email
ATTACHMENT_PATH = 'attachments/Resume.pdf'           # Attachment file
HTML_TEMPLATE_PATH = 'templates/welcome_template.html'  # HTML email template
LOG_FILE = 'sent_log.txt'                            # Log file
EMAIL_SUBJECT = 'Application for React JS Developer Position'
# ======================

def load_recipients():
    try:
        df = pd.read_excel(EXCEL_FILE)
        return df[['Name', 'Email']].dropna().to_dict(orient='records')
    except Exception as e:
        print(f"Error: Failed to load Excel file: {e}")
        return []

def load_html_template(name):
    try:
        with open(HTML_TEMPLATE_PATH, 'r', encoding='utf-8') as f:
            html = f.read()
            return html.replace('{{name}}', name)
    except Exception as e:
        print(f"Error: Failed to load HTML template: {e}")
        return f"<p>Hello {name},</p><p>This is a fallback message.</p>"

def send_email(smtp, to_name, to_email):
    msg = EmailMessage()
    msg['Subject'] = EMAIL_SUBJECT
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = to_email

    # Set plain and HTML content
    msg.set_content(f"Hello {to_name},\nThis is a test email from Python.")
    msg.add_alternative(load_html_template(to_name), subtype='html')

    # Attach file if it exists
    file = Path(ATTACHMENT_PATH)
    if file.exists():
        with open(file, 'rb') as f:
            msg.add_attachment(f.read(), maintype='application', subtype='octet-stream', filename=file.name)

    # Send and log
    smtp.send_message(msg)
    with open(LOG_FILE, 'a', encoding='utf-8') as log:
        log.write(f"Sent to {to_email} ({to_name})\n")
    print(f"Sent to {to_email}")

def main():
    recipients = load_recipients()
    if not recipients:
        print("Warning: No recipients found.")
        return

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(EMAIL_ADDRESS, APP_PASSWORD)
            for person in recipients:
                try:
                    send_email(smtp, person['Name'], person['Email'])
                except Exception as e:
                    print(f"Error sending to {person['Email']}: {e}")
    except Exception as e:
        print(f"Error: Failed to connect or login: {e}")

if __name__ == '__main__':
    main()
