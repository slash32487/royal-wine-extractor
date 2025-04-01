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

        def get_line_type(line, sizes):
            text = line.strip()
            if not text or any(p in text for p in FOOTER_PATTERNS):
                return "skip"
            if re.match(r"\d{5}\s+(?:\d{4}|NV)\s+\d+\s*/\s*\d+\s+\d+\.\d{2}(?:\s+\d+\.\d{2})?", text):
                return "item"
            if re.match(r"\$\d+\.\d{2} on \d+cs\s+\d+\.\d{2}\s+\d+\.\d{2}", text):
                return "discount"
            if text.upper() == "NEW":
                return "skip"
            if any(x in text.upper() for x in ["COMBO PACK", "GIFT PACK", "BOTTLES EACH", "VARIATION"]):
                return "combo"
            return "product"

        i = 0
        while i < len(sorted_lines):
            _, line, sizes = sorted_lines[i]
            ltype = get_line_type(line, sizes)

            if ltype == "region":
                current_region = line
                i += 1
                continue
            elif ltype == "brand":
                current_brand = line
                last_valid_product_name = None
                i += 1
                continue
            elif ltype in ["combo", "skip"]:
                i += 1
                continue
            elif ltype == "item":
                item_match = re.match(r"(\d{5})\s+(\d{4}|NV)\s+(\d+)\s*/\s*(\d+)\s+(\d+\.\d{2})(?:\s+(\d+\.\d{2}))?", line)
                pname_lines = []
                for j in range(i - 1, max(i - 8, -1), -1):
                    _, prev_line, prev_sizes = sorted_lines[j]
                    if get_line_type(prev_line, prev_sizes) in ["item", "combo", "skip"]:
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
                    "Item#": item_match.group(1),
                    "Vintage": item_match.group(2),
                    "Product Name": pname,
                    "Bottles per Case": item_match.group(3),
                    "Bottle Size": item_match.group(4),
                    "Case Price": item_match.group(5),
                    "Bottle Price": item_match.group(6) if item_match.group(6) else "",
                    "Discounts": "",
                    "Name Inferred": "Yes" if inferred else "No"
                }

                discounts = []
                j = i + 1
                while j < len(sorted_lines):
                    _, next_line, next_sizes = sorted_lines[j]
                    if get_line_type(next_line, next_sizes) == "discount":
                        dmatch = re.match(r"\$(\d+\.\d{2}) on (\d+cs)\s+(\d+\.\d{2})\s+(\d+\.\d{2})", next_line)
                        if dmatch:
                            discounts.append(f"${dmatch.group(1)} on {dmatch.group(2)}: {dmatch.group(3)} / {dmatch.group(4)}")
                            j += 1
                        else:
                            break
                    else:
                        break
                if discounts:
                    item["Discounts"] = "; ".join(discounts)
                results.append(item)
                i = j
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
        st.warning("No items extracted. Previewing first 20 lines for debug:")
        doc = fitz.open(stream=BytesIO(pdf_bytes), filetype="pdf")
        for page in doc:
            blocks = page.get_text("dict")["blocks"]
            text_lines = []
            for b in blocks:
                for l in b.get("lines", []):
                    line_text = " ".join([span["text"] for span in l["spans"] if span["text"].strip()])
                    font_sizes = [span.get("size", 0) for span in l["spans"] if "size" in span]
                    text_lines.append((l["bbox"][1], line_text.strip(), font_sizes))
            for top, line, fonts in sorted(text_lines, key=lambda x: x[0])[:20]:
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
