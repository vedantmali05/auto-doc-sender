import os
import json
import pandas as pd
import smtplib
from email.message import EmailMessage

print("✉️ Running send_email.py")

# ---------------- PATHS ----------------
CONFIG_PATH = "config/event_config.json"
OUTPUT_DIR = "output"

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# ---------------- LOAD CONFIG ----------------
if not os.path.exists(CONFIG_PATH):
    raise FileNotFoundError("Config file not found")

with open(CONFIG_PATH, "r") as f:
    config = json.load(f)

mode = config["mode"]
dry_run = config.get("dry_run", False)

template_cfg = config["template"]
email_cfg = config["email"]

# ---------------- DATA PATH (MODE BASED) ----------------
DATA_PATH = config["data"][mode]

if not os.path.exists(DATA_PATH):
    raise FileNotFoundError(f"Data file not found: {DATA_PATH}")

# ---------------- EMAIL SETUP ----------------
sender_email = email_cfg["sender"]["address"]
sender_name = email_cfg["sender"]["display_name"]
receiver_column = email_cfg["receiver_email_column"]

body_text_template = email_cfg.get("body_text", "")
body_html_template = email_cfg.get("body_html", "")

sender_password = os.getenv("EMAIL_PASSWORD")
if not sender_password:
    raise RuntimeError("EMAIL_PASSWORD not set in environment")

# ---------------- LOAD DATA ----------------
df = pd.read_excel(DATA_PATH)
rows = df.to_dict(orient="records")

if not rows:
    raise RuntimeError("No data to process")

# ---------------- SEND EMAIL ----------------
def send_email(row):
    if receiver_column not in row or not row[receiver_column]:
        print("❌ Missing receiver email, skipping row")
        return

    filename = template_cfg["output_pdf_name"].format(**row).replace(" ", "_")
    pdf_path = os.path.join(OUTPUT_DIR, filename)

    if not os.path.exists(pdf_path):
        print(f"❌ PDF not found (skipping): {filename}")
        return

    if dry_run:
        print(f"🟡 DRY RUN → Would send to {row[receiver_column]} | {filename}")
        return

    msg = EmailMessage()
    msg["From"] = f"{sender_name} <{sender_email}>"
    msg["To"] = row[receiver_column]
    msg["Subject"] = email_cfg["subject"]

    body_text = body_text_template.format(**row)
    body_html = body_html_template.format(**row)

    msg.set_content(body_text)

    if body_html:
        msg.add_alternative(body_html, subtype="html")

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

    print(f"✅ Email sent: {row[receiver_column]}")

# ---------------- MAIN ----------------
for row in rows:
    send_email(row)

print("✅ Email sending completed")
