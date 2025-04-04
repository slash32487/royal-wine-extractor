import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.title("Royal Wine ETS Excel Extractor")

# Regex patterns for identifying fields
re_item = re.compile(r"^\d{5}$")
re_vintage = re.compile(r"^(199\d|20[0-2]\d|2025)$")
re_case_size = re.compile(r"^(\d{1,2})\s*/\s*(\d+(?:\.\d+)?(?:L|ML)?)$")
re_price = re.compile(r"^\d+\.\d{2}$")
re_discount = re.compile(r"^\$\d+\.\d{2} on \d+cs$")

@st.cache_data
def extract_from_excel(file):
    df_raw = pd.read_excel(file, header=None)
    df_raw.dropna(how='all', inplace=True)
    df_raw.fillna("", inplace=True)

    items = []
    current = {}

    for _, row in df_raw.iterrows():
        if re_item.fullmatch(str(row[0]).strip()):
            if current:
                items.append(current)
            current = {
                "Item#": str(row[0]).strip(),
                "Product Name": "",
                "Vintage": "",
                "Bottles per Case": "",
                "Bottle Size": "",
                "Case Price": "",
                "Bottle Price": "",
                "Discounts": ""
            }
            continue

        if not current:
            continue

        for cell in row:
            text = str(cell).strip()
            if not text:
                continue
            if re_vintage.fullmatch(text):
                current["Vintage"] = text
            elif re_case_size.fullmatch(text):
                m = re_case_size.match(text)
                current["Bottles per Case"] = m.group(1)
                current["Bottle Size"] = m.group(2)
            elif re_price.fullmatch(text):
                if not current["Case Price"]:
                    current["Case Price"] = text
                elif not current["Bottle Price"]:
                    current["Bottle Price"] = text
            elif re_discount.fullmatch(text):
                current["Discounts"] += text + "; "
            else:
                current["Product Name"] += text + " "

    if current:
        items.append(current)

    return pd.DataFrame(items)

uploaded_file = st.file_uploader("Upload Royal Excel File (ETS Exported)", type="xlsx")

if uploaded_file:
    try:
        df = extract_from_excel(uploaded_file)
        if df.empty:
            st.warning("No data extracted. Please verify the Excel content.")
        else:
            st.success(f"Extracted {len(df)} items.")
            st.dataframe(df)

            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Extracted")
            st.download_button("📥 Download Excel", buffer.getvalue(), "ets_export.xlsx")
    except Exception as e:
        st.error(f"Extraction error: {e}")
