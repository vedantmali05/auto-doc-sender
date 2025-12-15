import os
import json
import pandas as pd
import smtplib
from email.message import EmailMessage

# ---------------- PATHS ----------------
CONFIG_PATH = "config/event_config.json"
DATA_PATH = "data/participants.xlsx"
OUTPUT_DIR = "output"

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# ---------------- LOAD CONFIG ----------------
if not os.path.exists(CONFIG_PATH):
    raise FileNotFoundError("Config file not found")

with open(CONFIG_PATH, "r") as f:
    config = json.load(f)

template_cfg = config["template"]
email_cfg = config["email"]
testing_cfg = config.get("testing", {"enabled": False})

# ---------------- EMAIL SETUP ----------------
sender_email = email_cfg["sender"]["address"]
sender_name = email_cfg["sender"]["display_name"]
sender_password = os.getenv("EMAIL_PASSWORD")

if not sender_password:
    raise RuntimeError("EMAIL_PASSWORD not set")

# ---------------- LOAD DATA ----------------
if testing_cfg.get("enabled"):
    rows = testing_cfg.get("receivers", [])
else:
    if not os.path.exists(DATA_PATH):
        raise FileNotFoundError("participants.xlsx not found")
    df = pd.read_excel(DATA_PATH)
    rows = df.to_dict(orient="records")

if not rows:
    raise RuntimeError("No data to process")

# ---------------- SEND EMAIL ----------------
def send_email(row):
    filename = template_cfg["output_name_format"].format(**row)
    filename = filename.replace(" ", "_")
    pdf_path = os.path.join(OUTPUT_DIR, filename)

    if not os.path.exists(pdf_path):
        print(f"❌ PDF not found: {filename}")
        return

    msg = EmailMessage()
    msg["From"] = f"{sender_name} <{sender_email}>"
    msg["To"] = row["email"]
    msg["Subject"] = email_cfg["subject"]

    msg.set_content(email_cfg["body"].format(**row))

    with open(pdf_path, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="pdf",
            filename=filename
        )

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(msg)

    print(f"Email sent: {filename}")

# ---------------- MAIN ----------------
for row in rows:
    send_email(row)

print("✅ Email sending completed")
