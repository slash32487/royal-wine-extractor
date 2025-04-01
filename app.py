import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
from io import BytesIO

FOOTER_PATTERNS = [
    "ROYAL WINE CORP", "BEVERAGE MEDIA", "ORDER DEPT", "WWW.ROYALWINES.COM",
    "TEL:", "FAX:", "NASSAU"
]

REGION_FONT_SIZE = 21.95
BRAND_FONT_SIZE = 11.04


def extract_items_from_pdf(file):
    results = []
    current_region = None
    current_brand = None
    last_valid_product_name = None

    doc = fitz.open(stream=file, filetype="pdf")
    for page in doc:
        blocks = page.get_text("dict")['blocks']
        lines = []
        for b in blocks:
            for l in b.get("lines", []):
                spans = l.get("spans", [])
                text = " ".join([span["text"] for span in spans if span.get("text")])
                if not text.strip():
                    continue
                font_sizes = list(set(span.get("size", 0) for span in spans))
                top = l["bbox"][1]
                lines.append((top, text.strip(), font_sizes))

        sorted_lines = sorted(lines, key=lambda x: x[0])

        def get_line_type(text, fonts):
            if not text or any(p in text for p in FOOTER_PATTERNS):
                return "skip"
            if text == "NEW":
                return "skip"
            if REGION_FONT_SIZE in fonts:
                return "region"
            if BRAND_FONT_SIZE in fonts:
                return "brand"
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
        while i < len(sorted_lines):
            top, text, fonts = sorted_lines[i]
            line_type = get_line_type(text, fonts)

            if line_type == "region":
                current_region = text.strip()
                i += 1
                continue
            elif line_type == "brand":
                current_brand = text.strip()
                i += 1
                continue
            elif line_type == "item_id":
                try:
                    item_id = text.strip()
                    vintage = sorted_lines[i + 1][1].strip()
                    size_parts = sorted_lines[i + 2][1].strip().split("/")
                    bottles_per_case = size_parts[0].strip()
                    bottle_size = size_parts[1].strip()
                    price_line = sorted_lines[i + 3][1].strip().split()
                    case_price = price_line[0]
                    bottle_price = price_line[1] if len(price_line) > 1 else ""

                    discounts = []
                    j = i + 4
                    while j < len(sorted_lines):
                        d_text = sorted_lines[j][1].strip()
                        if get_line_type(d_text, []) == "discount":
                            discounts.append(d_text)
                            j += 1
                        else:
                            break

                    # Collect name from lines above item_id
                    pname_lines = []
                    for k in range(i - 1, max(i - 10, -1), -1):
                        prev_text = sorted_lines[k][1].strip()
                        prev_fonts = sorted_lines[k][2]
                        if get_line_type(prev_text, prev_fonts) in ["item_id", "vintage", "size", "price", "discount", "skip"]:
                            break
                        pname_lines.insert(0, prev_text)
                    pname = " ".join(pname_lines).strip()
                    if not pname:
                        pname = last_valid_product_name or "[MISSING NAME]"
                        inferred = True
                    else:
                        last_valid_product_name = pname
                        inferred = False

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
                        "Discounts": "; ".join(discounts),
                        "Name Inferred": "Yes" if inferred else "No"
                    }
                    results.append(item)
                    i = j
                except Exception as e:
                    i += 1
                    continue
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
            blocks = page.get_text("dict")['blocks']
            lines = []
            for b in blocks:
                for l in b.get("lines", []):
                    spans = l.get("spans", [])
                    text = " ".join([span["text"] for span in spans if span.get("text")])
                    font_sizes = list(set(span.get("size", 0) for span in spans))
                    top = l["bbox"][1]
                    lines.append((top, text.strip(), font_sizes))
            for top, text, fonts in sorted(lines, key=lambda x: x[0])[:80]:
                st.text(f"{text}  | Sizes: {fonts}")
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
