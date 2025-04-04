import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
from io import BytesIO

st.title("Royal Wine ETS PDF Extractor")

# Patterns based on ETS-generated structure
re_item = re.compile(r"^\d{5}$")
re_vintage = re.compile(r"^(199\d|20[0-2]\d|2025)$")
re_case_size = re.compile(r"^(\d{1,2})\s*/\s*(\d+(\.\d+)?(L|ML)?)$")
re_price_pair = re.compile(r"^(\d+\.\d{2})\s+(\d+\.\d{2})$")
re_discount = re.compile(r"^\$\d+\.\d{2} on \d+cs$")

SKIP_LINES = [
    "ROYAL WINE CORP.", "C & R DISTRIBUTORS", "BEVERAGE MEDIA",
    "TEL:", "FAX:", "Lic#", "Order Department", "CR-WINES", "APRIL, 2025 PRICES",
    "Item# Vint BPC Size Case Bottle"
]

# Clean and flatten PDF text
@st.cache_data
def extract_ets_pdf(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    lines = []
    for page in doc:
        blocks = page.get_text("blocks")
        for b in blocks:
            for line in b[4].split("\n"):
                text = line.strip()
                if text:
                    lines.append(text)

    items = []
    debug_log = []
    current_item = None

    for line in lines:
        debug_log.append({"Line": line})

        # skip known garbage
        if any(skip in line for skip in SKIP_LINES):
            continue

        if re_item.fullmatch(line):
            if current_item:
                items.append(current_item)
            current_item = {
                "Item#": line,
                "Vintage": "",
                "Product Name": "",
                "Bottles per Case": "",
                "Bottle Size": "",
                "Case Price": "",
                "Bottle Price": "",
                "Discounts": ""
            }
            continue

        if not current_item:
            continue

        if re_vintage.fullmatch(line):
            current_item["Vintage"] = line
            continue

        m = re_case_size.fullmatch(line)
        if m:
            current_item["Bottles per Case"] = m.group(1)
            current_item["Bottle Size"] = m.group(2)
            continue

        m = re_price_pair.fullmatch(line)
        if m:
            current_item["Case Price"] = m.group(1)
            current_item["Bottle Price"] = m.group(2)
            continue

        if re_discount.fullmatch(line):
            current_item["Discounts"] += line + "; "
            continue

        current_item["Product Name"] += line + " "

    if current_item:
        items.append(current_item)

    return items, debug_log

uploaded_file = st.file_uploader("Upload Royal PDF (ETS Format)", type="pdf")
debug = st.checkbox("Show debug log")

if uploaded_file:
    pdf_bytes = uploaded_file.read()
    try:
        data, log = extract_ets_pdf(pdf_bytes)
        if not data:
            st.warning("No data extracted. Please verify the PDF structure.")
        else:
            df = pd.DataFrame(data)
            st.success(f"Extracted {len(df)} items.")
            st.dataframe(df)

            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Extracted")
                pd.DataFrame(log).to_excel(writer, index=False, sheet_name="Debug Log")
            st.download_button("📥 Download Excel", buffer.getvalue(), "ets_export.xlsx")
    except Exception as e:
        st.error(f"Extraction error: {e}")

    if debug:
        st.subheader("Debug Log")
        st.dataframe(pd.DataFrame(log).head(100))
