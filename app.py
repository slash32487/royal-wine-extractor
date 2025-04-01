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

    def is_brand_line(line, following_lines):
        banned_phrases = ["ROYAL WINE CORP", "TEL:", "WWW.ROYALWINES.COM", "FAX:", "NASSAU", "BROOKLYN"]
        if any(b in line.upper() for b in banned_phrases):
            return False
        if line.isupper() and not any(char.isdigit() for char in line) and len(line.split()) <= 5:
            for lookahead in following_lines[:2]:
                if re.search(r"\d{5}\s+(\d{4}|NV)\s+\d+\s*/\s*\d+", lookahead):
                    return True
        return False

    def is_combo_line(line):
        return "COMBO PACK" in line.upper() or "GIFT PACK" in line.upper() or "BOTTLES EACH" in line.upper()

    def split_multiple_items(line):
        return re.findall(r"(\d{5}\s+(?:\d{4}|NV)\s+\d+\s*/\s*\d+\s+\d+\.\d{2}(?:\s+\d+\.\d{2})?)", line)

    while i < len(lines):
        line = lines[i].strip()

        if is_brand_line(line, lines[i+1:i+3]):
            brand = line
            i += 1
            continue

        if is_combo_line(line):
            i += 1
            continue

        items_in_line = split_multiple_items(line)
        for idx, item_str in enumerate(items_in_line):
            match = re.match(r"(\d{5})\s+(\d{4}|NV)\s+(\d+)\s*/\s*(\d+)\s+(\d+\.\d{2})(?:\s+(\d+\.\d{2}))?", item_str)
            if match:
                pname_lines = []
                if idx == 0:
                    for k in range(i-1, max(i-6, -1), -1):
                        prev = lines[k].strip()
                        if prev and not any(x in prev.upper() for x in ["RATED", "AWARD", "CHALLENGE", "GOLD", "SILVER", "PLATINUM", "DOUBLE"]):
                            if not is_brand_line(prev, lines[k+1:k+3]):
                                pname_lines.insert(0, prev)
                            else:
                                break
                pname = " ".join(pname_lines).strip()

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

                results.append(item)

        if len(items_in_line) == 1:
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
                results[-1]["Discounts"] = "; ".join(discount_lines)
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
