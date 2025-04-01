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

    region = "California"  # Hardcoded based on your spec
    brand = None
    results = []
    i = 0

    while i < len(lines):
        line = lines[i].strip()

        # Detect brand
        if line.isupper() and not any(char.isdigit() for char in line):
            brand = line
            i += 1
            continue

        # Match main wine line with Item#, Vintage, BPC/Size, Case, Bottle Price
        match = re.match(r"(\d{5})\s+(\d{4}|NV)\s+(\d+)\s*/\s*(\d+[\s\w]*)\s+(\d+\.\d{2})(?:\s+(\d+\.\d{2}))?", line)
        if match:
            item = {
                "Region": region,
                "Brand": brand,
                "Item#": match.group(1),
                "Vintage": match.group(2),
                "Product Name": lines[i-1].strip(),  # assume previous line is the wine name
                "Bottles per Case": match.group(3),
                "Bottle Size": match.group(4),
                "Case Price": match.group(5),
                "Bottle Price": match.group(6) if match.group(6) else "",
                "Discounts": ""
            }

            # Collect discount lines
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

    # Excel export
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='WineData')
    st.download_button(
        label="Download Excel File",
        data=output.getvalue(),
        file_name="royal_wine_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
