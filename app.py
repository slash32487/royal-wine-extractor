import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

def extract_items_from_pdf(file):
    with pdfplumber.open(file) as pdf:
        results = []
        current_region = None
        current_brand = None
        last_valid_product_name = None

        for page in pdf.pages:
            words = page.extract_words(use_text_flow=True, keep_blank_chars=False)
            lines = {}

            for w in words:
                top = round(w["top"])
                lines.setdefault(top, []).append(w)

            sorted_lines = []
            for top in sorted(lines.keys()):
                line = " ".join([w["text"] for w in sorted(lines[top], key=lambda x: x["x"])])
                fonts = set((w["fontname"], int(float(w["size"]))) for w in lines[top])
                sorted_lines.append((top, line.strip(), fonts))

            def get_line_type(line, fonts):
                text = line.strip()
                if not text:
                    return "skip"
                if re.match(r"\d{5}\s+(\d{4}|NV)\s+\d+\s*/\s*\d+\s+\d+\.\d{2}(\s+\d+\.\d{2})?", text):
                    return "item"
                if re.match(r"\$\d+\.\d{2} on \d+cs\s+\d+\.\d{2}\s+\d+\.\d{2}", text):
                    return "discount"
                if "NEW" == text.upper():
                    return "skip"
                if any(x in text.upper() for x in ["COMBO PACK", "GIFT PACK", "BOTTLES EACH", "VARIATION"]):
                    return "combo"
                sizes = [f[1] for f in fonts]
                if len(sizes) > 0:
                    avg_size = sum(sizes) / len(sizes)
                    if avg_size > 11:
                        return "region" if text.isupper() else "brand"
                return "product"

            buffer = []
            i = 0
            while i < len(sorted_lines):
                _, line, fonts = sorted_lines[i]
                ltype = get_line_type(line, fonts)

                if ltype == "region":
                    current_region = line
                    i += 1
                    continue
                elif ltype == "brand":
                    current_brand = line
                    last_valid_product_name = None
                    i += 1
                    continue
                elif ltype == "combo" or ltype == "skip":
                    i += 1
                    continue
                elif ltype == "item":
                    item_match = re.match(r"(\d{5})\s+(\d{4}|NV)\s+(\d+)\s*/\s*(\d+)\s+(\d+\.\d{2})(?:\s+(\d+\.\d{2}))?", line)
                    pname_lines = []
                    for j in range(i - 1, max(i - 8, -1), -1):
                        _, prev_line, prev_fonts = sorted_lines[j]
                        prev_type = get_line_type(prev_line, prev_fonts)
                        if prev_type in ["item", "combo", "skip"]:
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
                        _, next_line, next_fonts = sorted_lines[j]
                        if get_line_type(next_line, next_fonts) == "discount":
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
    with st.spinner("Analyzing and Extracting..."):
        df = extract_items_from_pdf(uploaded_file)

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
