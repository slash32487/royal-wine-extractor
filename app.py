import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

def extract_items_from_pdf(file):
    with pdfplumber.open(file) as pdf:
        lines = []
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                lines.extend(text.split('\n'))

    region = "California"
    brand = None
    results = []
    i = 0

    def is_brand_line(line):
        return line.isupper() and not any(char.isdigit() for char in line) and len(line.split()) <= 5

    def is_combo_line(line):
        return "COMBO PACK" in line.upper() or "GIFT PACK" in line.upper() or "BOTTLES EACH" in line.upper()

    while i < len(lines):
        line = lines[i].strip()

        # Detect brand
        if is_brand_line(line):
            brand = line
            i += 1
            continue

        # Skip combo packs
        if is_combo_line(line):
            i += 1
            continue

        # Match main product line
        match = re.match(r"(\d{5})\s+(\d{4}|NV)\s+(\d+)\s*/\s*(\d+[\w\s]*)\s+(\d+\.\d{2})(?:\s+(\d+\.\d{2}))?", line)
        if match:
            # Find the real product name (look backwards until non-empty non-award line)
            pname = ""
            for k in range(i-1, max(i-6, -1), -1):
                prev = lines[k].strip()
                if prev and not any(x in prev.upper() for x in ["RATED", "AWARD", "CHALLENGE", "GOLD", "SILVER", "PLATINUM", "DOUBLE"]):
                    pname = prev
                    break

            item = {
                "Region": region,
                "Brand": brand,
                "Item#": match.group(1),
                "Vintage": match.group(2),
                "Product Name": pname,
                "Bottles per Case": match.group(3),
                "Bottle Size": match.group(4),
                "Case Price": match.group(5),
                "Bottle Price": match.group(6) if match.group(6) else "",
                "Discounts": ""
            }

            discount_lines = []
            j = i + 1
            while j < len(lines):
                dline = lines[j].strip()
                dmatch = re.match(r"\$(\d+\.\d{2}) on (\d+cs)\s+(\d+\.\d{2})\s+(\d+\.\d{2})", dline)
                if dmatch:
                    discount_str = f"${dmatch.group(1)} on {dmatch.group(2)}: {dmatch.group(3)} / {dmatch.group(4)}"
                    discount_lines.append(discount_str)
                    j += 1
                else:
                    break

            if discount_lines:
                item["Discounts"] = "; ".join(discount_lines)

            results.append(item)
            i = j
        else:
            i += 1

    return pd.DataFrame(results)

st.title("Royal Wine PDF to Excel Extractor")

uploaded_file = st.file_uploader("Upload Royal Wine PDF", type="pdf")
if uploaded_file:
    with st.spinner("Processing PDF..."):
        df = extract_items_from_pdf(uploaded_file)

    st.success("Extraction complete!")
    st.subheader("Preview of Extracted Data")
    st.dataframe(df)

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='WineData')
    st.download_button(
        label="Download Excel File",
        data=output.getvalue(),
        file_name="royal_wine_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
