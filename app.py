import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
from io import BytesIO

FOOTER_PATTERNS = [
    "ROYAL WINE CORP", "BEVERAGE MEDIA", "ORDER DEPT", "WWW.ROYALWINES.COM",
    "TEL:", "FAX:", "NASSAU"
]

def extract_items_from_pdf(file):
    results = []
    current_region = None
    current_brand = None
    last_valid_product_name = None

    doc = fitz.open(stream=file, filetype="pdf")
    for page in doc:
        blocks = page.get_text("dict")["blocks"]
        text_lines = []
        for b in blocks:
            for l in b.get("lines", []):
                line_text = " ".join([span["text"] for span in l["spans"] if span["text"].strip()])
                font_sizes = [span.get("size", 0) for span in l["spans"] if "size" in span]
                text_lines.append((l["bbox"][1], line_text.strip(), font_sizes))

        sorted_lines = sorted(text_lines, key=lambda x: x[0])

        def get_line_type(text):
            if not text or any(p in text for p in FOOTER_PATTERNS):
                return "skip"
            if re.fullmatch(r"\d{5}", text):
                return "item_id"
            if re.fullmatch(r"\d{4}|NV", text):
                return "vintage"
            if re.fullmatch(r"\d+\s*/\s*\d+", text):
                return "size"
            if re.fullmatch(r"\d+\.\d{2}(\s+\d+\.\d{2})?", text):
                return "price"
            if re.match(r"\$\d+\.\d{2} on \d+cs", text):
                return "discount"
            return "text"

        i = 0
        while i < len(sorted_lines) - 4:
            group = [sorted_lines[i + j][1] for j in range(5)]
            types = [get_line_type(t) for t in group]

            if types[0] == "item_id" and types[1] == "vintage" and types[2] == "size" and "price" in types[3:]:
                item_id = group[0]
                vintage = group[1]
                size_split = re.split(r"\s*/\s*", group[2])
                bottles_per_case = size_split[0]
                bottle_size = size_split[1]
                prices = " ".join(group[3:5]).split()
                case_price = prices[0]
                bottle_price = prices[1] if len(prices) > 1 else ""

                pname_lines = []
                for j in range(i - 1, max(i - 8, -1), -1):
                    prev_line = sorted_lines[j][1]
                    if get_line_type(prev_line) in ["item_id", "vintage", "size", "price", "discount", "skip"]:
                        break
                    pname_lines.insert(0, prev_line.strip())

                pname = " ".join(pname_lines).strip()
                inferred = False
                if not pname:
                    pname = last_valid_product_name or "[MISSING NAME]"
                    inferred = True
                else:
                    last_valid_product_name = pname

                item = {
                    "Region": current_region or "[UNKNOWN REGION]",
                    "Brand": current_brand or "[UNKNOWN BRAND]",
                    "Item#": item_id,
                    "Vintage": vintage,
                    "Product Name": pname,
                    "Bottles per Case": bottles_per_case,
                    "Bottle Size": bottle_size,
                    "Case Price": case_price,
                    "Bottle Price": bottle_price,
                    "Discounts": "",
                    "Name Inferred": "Yes" if inferred else "No"
                }

                # Look for following discount lines
                discounts = []
                k = i + 5
                while k < len(sorted_lines):
                    d_line = sorted_lines[k][1]
                    if get_line_type(d_line) == "discount":
                        discounts.append(d_line)
                        k += 1
                    else:
                        break
                if discounts:
                    item["Discounts"] = "; ".join(discounts)
                results.append(item)
                i = k
            else:
                i += 1

    return pd.DataFrame(results)

st.title("Royal Wine PDF to Excel Extractor")

uploaded_file = st.file_uploader("Upload Royal Wine PDF", type="pdf")
if uploaded_file:
    pdf_bytes = uploaded_file.read()
    with st.spinner("Analyzing and Extracting using PyMuPDF..."):
        df = extract_items_from_pdf(BytesIO(pdf_bytes))

    if df.empty:
        st.warning("No items extracted. Previewing first 80 lines for debug:")
        doc = fitz.open(stream=BytesIO(pdf_bytes), filetype="pdf")
        for page in doc:
            blocks = page.get_text("dict")["blocks"]
            text_lines = []
            for b in blocks:
                for l in b.get("lines", []):
                    line_text = " ".join([span["text"] for span in l["spans"] if span["text"].strip()])
                    font_sizes = [span.get("size", 0) for span in l["spans"] if "size" in span]
                    text_lines.append((l["bbox"][1], line_text.strip(), font_sizes))
            for top, line, fonts in sorted(text_lines, key=lambda x: x[0])[:80]:
                st.text(f"{line}  | Sizes: {fonts}")
            break
    else:
        st.success("Extraction complete!")
        st.dataframe(df)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="WineData")
        st.download_button(
            label="Download Excel File",
            data=output.getvalue(),
            file_name="royal_wine_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
