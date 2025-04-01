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

    results = []
    current_region = None
    current_brand = None

    def is_region_line(line):
        return line.isupper() and len(line.split()) <= 3 and not re.search(r'[\d\$/]', line)

    def is_brand_line(line):
        banned = ["ROYAL WINE CORP", "TEL:", "WWW.ROYALWINES.COM", "FAX:", "NASSAU", "BROOKLYN"]
        return (
            line.isupper()
            and not any(b in line for b in banned)
            and not re.search(r'[\d\$/]', line)
            and 1 <= len(line.split()) <= 5
        )

    def is_combo_line(line):
        return any(x in line.upper() for x in ["COMBO PACK", "GIFT PACK", "BOTTLES EACH"])

    def is_item_line(line):
        return re.match(r"\d{5}\s+(\d{4}|NV)\s+\d+\s*/\s*\d+\s+\d+\.\d{2}(\s+\d+\.\d{2})?", line)

    def is_discount_line(line):
        return re.match(r"\$\d+\.\d{2} on \d+cs\s+\d+\.\d{2}\s+\d+\.\d{2}", line)

    i = 0
    while i < len(lines):
        line = lines[i].strip()

        # Set region if it's likely a region header
        if is_region_line(line):
            current_region = line
            i += 1
            continue

        # Set brand if it's valid
        if is_brand_line(line):
            current_brand = line
            i += 1
            continue

        # Skip combo/gift lines
        if is_combo_line(line):
            i += 1
            continue

        # Handle item lines
        if is_item_line(line):
            item_match = re.match(r"(\d{5})\s+(\d{4}|NV)\s+(\d+)\s*/\s*(\d+)\s+(\d+\.\d{2})(?:\s+(\d+\.\d{2}))?", line)
            if item_match:
                pname_lines = []
                for k in range(i - 1, max(i - 8, -1), -1):
                    prev = lines[k].strip()
                    if is_brand_line(prev) or is_combo_line(prev):
                        break
                    if re.search(r'[\d]', prev):
                        continue
                    pname_lines.insert(0, prev)

                pname = " ".join(pname_lines).strip() or "[MISSING NAME]"

                item = {
                    "Region": current_region or "[UNKNOWN REGION]",
                    "Brand": current_brand or "[UNKNOWN BRAND]",
                    "Item#": item_match.group(1),
                    "Vintage": item_match.group(2),
                    "Product Name": pname,
                    "Bottles per Case": item_match.group(3),
                    "Bottle Size": item_match.group(4),
                    "Case Price": item_match.group(5),
                    "Bottle Price": item_match.group(6) if item_match.group(6) else "",
                    "Discounts": ""
                }

                # Gather discounts
                discount_lines = []
                j = i + 1
                while j < len(lines):
                    next_line = lines[j].strip()
                    if is_discount_line(next_line):
                        dmatch = re.match(r"\$(\d+\.\d{2}) on (\d+cs)\s+(\d+\.\d{2})\s+(\d+\.\d{2})", next_line)
                        if dmatch:
                            dstr = f"${dmatch.group(1)} on {dmatch.group(2)}: {dmatch.group(3)} / {dmatch.group(4)}"
                            discount_lines.append(dstr)
                            j += 1
                        else:
                            break
                    else:
                        break

                if discount_lines:
                    item["Discounts"] = "; ".join(discount_lines)
                results.append(item)
                i = j
                continue

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
