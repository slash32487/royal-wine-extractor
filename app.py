import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
from io import BytesIO
from collections import defaultdict

st.title("Royal Wine PDF to Excel Extractor (First Page Test)")

valid_vintages = {str(y) for y in range(1990, 2026)}
valid_case_sizes = {"1", "3", "6", "12", "24", "36", "48"}

def is_valid_case_size_format(text):
    return re.match(rf"^({'|'.join(valid_case_sizes)})\s*/\s*\d+[A-Z]*$", text)

def is_discount_line(text):
    return re.match(r"\$\d+\.\d{2} on \d+cs", text)

def is_price_pair(text):
    return re.fullmatch(r"\d+\.\d{2} \d+\.\d{2}", text)

def is_item_number(text):
    return re.fullmatch(r"\d{5}", text)

def is_vintage(text):
    return text in valid_vintages

def is_centered_brand_line(line_words):
    text = " ".join(w[4] for w in line_words)
    return text.isupper() and 4 <= len(line_words) <= 6 and all(150 < w[0] < 450 for w in line_words)

def is_region_block(w):
    return w[3] - w[1] > 20 and w[4].isupper() and w[0] < 200

def extract_pdf_data(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    first_page = doc[0]
    words = first_page.get_text("words")
    words.sort(key=lambda w: (round(w[1], 1), w[0]))
    lines_by_y = defaultdict(list)
    for w in words:
        lines_by_y[round(w[1], 1)].append(w)

    y_sorted = sorted(lines_by_y.keys())
    data = []
    debug_log = []
    current_region = None
    current_brand = None
    temp = defaultdict(str)

    def save_entry():
        if temp.get("Item#"):
            data.append({
                "Region": current_region,
                "Brand": current_brand,
                "Item#": temp["Item#"],
                "Vintage": temp["Vintage"],
                "Product Name": temp["Name"].strip(),
                "Bottles per Case": temp["BPC"],
                "Bottle Size": temp["BottleSize"],
                "Case Price": temp["CasePrice"],
                "Bottle Price": temp["BottlePrice"],
                "Discounts": temp["Discounts"].strip("; ")
            })
            temp.clear()

    for i, y in enumerate(y_sorted):
        line_words = lines_by_y[y]
        line = " ".join(w[4] for w in sorted(line_words, key=lambda x: x[0])).strip()
        debug_log.append({"Y": y, "Line": line})

        if any(is_region_block(w) for w in line_words):
            current_region = line.strip()
            continue

        if is_centered_brand_line(line_words):
            current_brand = line
            continue

        if any(is_item_number(w[4]) for w in line_words):
            save_entry()
            temp["Item#"] = next(w[4] for w in line_words if is_item_number(w[4]))
            continue

        if any(is_vintage(w[4]) for w in line_words):
            temp["Vintage"] = next(w[4] for w in line_words if is_vintage(w[4]))
            continue

        if any(is_valid_case_size_format(w[4]) for w in line_words):
            bpc_size = next(w[4] for w in line_words if is_valid_case_size_format(w[4]))
            bpc, size = bpc_size.split("/")
            temp["BPC"] = bpc.strip()
            temp["BottleSize"] = size.strip()
            continue

        if is_price_pair(line):
            cp, bp = line.split()
            temp["CasePrice"] = cp
            temp["BottlePrice"] = bp
            continue

        if is_discount_line(line):
            temp["Discounts"] += line + "; "
            continue

        temp["Name"] += line + " "

    save_entry()
    return data, debug_log

uploaded_file = st.file_uploader("Upload Full Royal Wine PDF", type="pdf")
debug = st.checkbox("🔍 Enable Debug Preview")

if uploaded_file:
    pdf_bytes = uploaded_file.read()
    try:
        extracted_data, debug_info = extract_pdf_data(pdf_bytes)
    except Exception as e:
        st.error(f"Extraction failed: {e}")
        st.stop()

    if not extracted_data:
        st.error("❌ No items extracted from first page. Check formatting.")
    else:
        df = pd.DataFrame(extracted_data)
        st.success(f"✅ Extracted {len(df)} wine entries from first page.")
        st.dataframe(df)

        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Extracted")
            pd.DataFrame(debug_info).to_excel(writer, index=False, sheet_name="Debug Log")

        st.download_button(
            "📥 Download Excel", buffer.getvalue(), "royal_wine_data_first_page.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    if debug:
        st.subheader("🔍 Raw Extracted Lines from First Page")
        st.dataframe(pd.DataFrame(debug_info).head(100))
