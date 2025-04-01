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
    last_known_product_name = ""

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
            if text.strip().upper() == "NEW":
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
            if re.fullmatch(r"\d+\.\d{2}", text):
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
                item_id = text.strip()
                vintage = ""
                bottles_per_case = ""
                bottle_size = ""
                case_price = ""
                bottle_price = ""
                discounts = []

                # Gather product name from lines ABOVE only
                pname_lines = []
                for k in range(i - 1, max(i - 10, -1), -1):
                    pt = sorted_lines[k][1].strip()
                    pf = sorted_lines[k][2]
                    if pt.upper() == "NEW" or "COMBINE" in pt.upper():
                        continue
                    if get_line_type(pt, pf) in ["item_id", "vintage", "size", "price", "discount", "skip"]:
                        break
                    pname_lines.insert(0, pt)
                pname = " ".join(pname_lines).strip()
                if pname:
                    last_known_product_name = pname
                else:
                    pname = last_known_product_name

                # Scan downward for fields tied to this item
                j = i + 1
                while j < len(sorted_lines):
                    t = sorted_lines[j][1].strip()
                    f = sorted_lines[j][2]
                    t_type = get_line_type(t, f)

                    if t_type == "vintage" and not vintage:
                        vintage = t
                    elif t_type == "size" and not bottles_per_case:
                        parts = t.split("/")
                        if len(parts) == 2:
                            bottles_per_case = parts[0].strip()
                            bottle_size = parts[1].strip()
                    elif t_type == "price" and not case_price:
                        prices = t.split()
                        case_price = prices[0]
                        bottle_price = prices[1] if len(prices) > 1 else ""
                        if bottles_per_case == "1":
                            bottle_price = case_price  # fallback for 1/bottle case
                    elif t_type == "discount":
                        if "$" in t and "on" in t:
                            discounts.append(t)
                    elif t_type == "item_id":
                        break
                    j += 1

                item = {
                    "Region": current_region or "[UNKNOWN REGION]",
                    "Brand": current_brand or "[UNKNOWN BRAND]",
                    "Item#": item_id,
                    "Vintage": vintage,
                    "Product Name": pname or "[MISSING NAME]",
                    "Bottles per Case": bottles_per_case,
                    "Bottle Size": bottle_size,
                    "Case Price": case_price,
                    "Bottle Price": bottle_price,
                    "Discounts": "; ".join(discounts),
                }
                results.append(item)
                i = j
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
