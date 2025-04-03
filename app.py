import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
from io import BytesIO
from collections import defaultdict

st.title("Royal Wine Extractor - AMS Mode")

# Predefined value sets
valid_vintages = {str(y) for y in range(1990, 2026)}
valid_case_sizes = {"1", "3", "6", "12", "24", "36", "48"}

# Regex patterns
re_item = re.compile(r"^\d{5}$")
re_vintage = re.compile(r"^(199\d|20[0-2]\d|2025)$")
re_case_size = re.compile(rf"^({'|'.join(valid_case_sizes)})\s*/\s*\d+[A-Z]*$")
re_price_pair = re.compile(r"^\d+\.\d{2} \d+\.\d{2}$")
re_discount = re.compile(r"^\$\d+\.\d{2} on \d+cs$")

# Header/Footer blocks to ignore
SKIP_LINES = [
    "ROYAL WINE CORP.", "C & R DISTRIBUTORS", "BEVERAGE MEDIA",
    "TEL:", "FAX:", "Lic#", "Order Department", "CR-WINES"
]

# Extract function
def extract_ams_style(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    words = []
    for page in doc:
        words += page.get_text("words")

    # sort by vertical then horizontal position
    words.sort(key=lambda w: (round(w[1], 1), w[0]))
    lines_by_y = defaultdict(list)
    for w in words:
        lines_by_y[round(w[1], 1)].append(w)

    y_sorted = sorted(lines_by_y.keys())
    blocks = []
    debug_log = []
    current_brand = None
    current_region = None
    record = {}

    for y in y_sorted:
        line_words = lines_by_y[y]
        text = " ".join(w[4] for w in sorted(line_words, key=lambda x: x[0])).strip()
        debug_log.append({"Y": y, "Line": text})

        # skip headers/footers
        if any(skip in text for skip in SKIP_LINES):
            continue

        # region block (all caps, white on black — inferred by height and boldness)
        if text.isupper() and len(text.split()) <= 4 and any(w[3] - w[1] > 20 for w in line_words):
            current_region = text
            continue

        # brand line
        if text.isupper() and len(text.split()) <= 6:
            current_brand = text
            continue

        # item number starts new entry
        if re_item.match(text):
            if record:
                blocks.append(record)
            record = {
                "Region": current_region,
                "Brand": current_brand,
                "Item#": text,
                "Vintage": "",
                "Product Name": "",
                "Bottles per Case": "",
                "Bottle Size": "",
                "Case Price": "",
                "Bottle Price": "",
                "Discounts": ""
            }
            continue

        if not record:
            continue  # Skip any line before first item#

        # vintage
        if re_vintage.match(text):
            record["Vintage"] = text
            continue

        # case size
        if re_case_size.match(text):
            bpc, size = text.split("/")
            record["Bottles per Case"] = bpc.strip()
            record["Bottle Size"] = size.strip()
            continue

        # price pair
        if re_price_pair.match(text):
            cp, bp = text.split()
            record["Case Price"] = cp
            record["Bottle Price"] = bp
            continue

        # discount line
        if re_discount.match(text):
            record["Discounts"] += text + "; "
            continue

        # fallback to product name
        record["Product Name"] += text + " "

    if record:
        blocks.append(record)

    return blocks, debug_log

# Streamlit upload and display
uploaded_file = st.file_uploader("Upload Royal Wine PDF (AMS Style)", type="pdf")
debug = st.checkbox("🔍 Show Debug Log")

if uploaded_file:
    pdf_bytes = uploaded_file.read()
    try:
        rows, log = extract_ams_style(pdf_bytes)
        if not rows:
            st.warning("No items extracted. Check PDF structure.")
        else:
            df = pd.DataFrame(rows)
            st.success(f"✅ Extracted {len(df)} entries.")
            st.dataframe(df)

            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Extracted")
                pd.DataFrame(log).to_excel(writer, index=False, sheet_name="Debug Log")

            st.download_button("📥 Download Excel", buffer.getvalue(), "royal_wine_data_ams.xlsx")
    except Exception as e:
        st.error(f"Extraction failed: {e}")

    if debug:
        st.subheader("🔍 Debug Log (First 100 lines)")
        st.dataframe(pd.DataFrame(log).head(100))
