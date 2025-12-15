import os
import json
import pandas as pd
import fitz  # PyMuPDF

print("🧾 Running generate_pdf.py")

# ---------------- PATHS ----------------
CONFIG_PATH = "config/event_config.json"
OUTPUT_DIR = "output"

os.makedirs(OUTPUT_DIR, exist_ok=True)

# ---------------- LOAD CONFIG ----------------
with open(CONFIG_PATH, "r") as f:
    config = json.load(f)

template_cfg = config["template"]
fields_cfg = config["fields"]
testing_cfg = config.get("testing", {"enabled": False})

DATA_PATH = "data/test.xlsx" if testing_cfg.get("enabled") else "data/participants.xlsx"

template_pdf = template_cfg["input_pdf"]
if not os.path.exists(template_pdf):
    raise FileNotFoundError("Template PDF not found")

if not os.path.exists(DATA_PATH):
    raise FileNotFoundError(f"Data file not found: {DATA_PATH}")

# ---------------- LOAD DATA ----------------
df = pd.read_excel(DATA_PATH)
rows = df.to_dict(orient="records")

if not rows:
    raise RuntimeError("No data to process")

# ---------------- PDF GENERATION ----------------
def create_pdf(row, output_path):
    doc = fitz.open(template_pdf)
    page = doc[0]
    page_height = page.rect.height

    for field_cfg in fields_cfg.values():
        col = field_cfg["source_column"]
        if col not in row:
            continue

        text = str(row[col])

        x = field_cfg["x"]
        y = page_height - field_cfg["y"]

        font_size = field_cfg["font_size"]
        font_name = field_cfg.get("font", "Times-Roman")
        align = field_cfg.get("align", "left")

        if field_cfg.get("auto_shrink") and len(text) > 22:
            font_size = max(font_size - (len(text) - 22), font_size - 6)

        if align == "center":
            text_width = fitz.get_text_length(text, fontname=font_name, fontsize=font_size)
            x = x - text_width / 2

        page.insert_text(
            fitz.Point(x, y),
            text,
            fontsize=font_size,
            fontname=font_name,
            color=(0, 0, 0),
            overlay=True
        )

    doc.save(output_path, incremental=False, deflate=False, garbage=0)
    doc.close()

# ---------------- MAIN ----------------
for row in rows:
    filename = template_cfg["output_name_format"].format(**row).replace(" ", "_")
    output_path = os.path.join(OUTPUT_DIR, filename)

    create_pdf(row, output_path)
    print(f"Generated PDF: {filename}")

print("✅ PDF generation completed (NO EMAILS SENT)")
