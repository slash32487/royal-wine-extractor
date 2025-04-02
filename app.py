import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
from io import BytesIO
from collections import defaultdict

st.title("Royal Wine PDF to Excel Extractor (First Page Test)")

valid_vintages = {str(y) for y in range(1990, 2026)}
valid_case_sizes = {"1", "3", "6", "12", "24", "36", "48"}

known_regions = {"CALIFORNIA", "FRANCE", "ISRAEL", "ITALY", "SPAIN", "SOUTH AFRICA"}
known_brands = {"STOUDEMIRE", "WEINSTOCK", "HAGAFEN CELLARS", "PHILIPPE LE HARDI"}

uploaded_file = st.file_uploader("Upload Full Royal Wine PDF", type="pdf")

if uploaded_file:
    doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
    first_page = doc[0]  # only use the first page
    data = []
    debug_log = []
    current_region = None
    current_brand = None
    temp = defaultdict(str)

    def save_entry():
        if temp.get("Item#"):
            data.append({
                "Region": current_region,
                "Brand": current_brand,
                "Item#": temp["Item#"],
                "Vintage": temp["Vintage"],
                "Product Name": temp["Name"].strip(),
                "Bottles per Case": temp["BPC"],
                "Bottle Size": temp["BottleSize"],
                "Case Price": temp["CasePrice"],
                "Bottle Price": temp["BottlePrice"],
                "Discounts": temp["Discounts"].strip("; ")
            })
            temp.clear()

    words = first_page.get_text("words")
    words.sort(key=lambda w: (round(w[1], 1), w[0]))
    lines_by_y = defaultdict(list)
    for w in words:
        lines_by_y[round(w[1], 1)].append((w[0], w[4]))

    for y in sorted(lines_by_y):
        line = " ".join([w[1] for w in sorted(lines_by_y[y])]).strip()
        debug_log.append({"Y": y, "Line": line})

        if line.upper() in known_regions:
            current_region = line.upper()
            continue
        if line.upper() in known_brands:
            current_brand = line.upper()
            continue
        if re.fullmatch(r"\d{5}", line):
            save_entry()
            temp["Item#"] = line
            continue
        if line in valid_vintages:
            temp["Vintage"] = line
            continue
        if re.match(rf"^({'|'.join(valid_case_sizes)}) /\d+[A-Z]*$", line):
            bpc, size = line.split("/")
            temp["BPC"] = bpc.strip()
            temp["BottleSize"] = size.strip()
            continue
        if re.fullmatch(r"\d+\.\d{2} \d+\.\d{2}", line):
            case_price, bottle_price = line.split()
            temp["CasePrice"] = case_price
            temp["BottlePrice"] = bottle_price
            continue
        if re.match(r"\$\d+\.\d{2} on \d+cs", line):
            temp["Discounts"] += line + "; "
            continue
        temp["Name"] += line + " "

    save_entry()

    if not data:
        st.error("❌ No items extracted from first page. Check formatting.")
    else:
        df = pd.DataFrame(data)
        st.success(f"✅ Extracted {len(df)} wine entries from first page.")
        st.dataframe(df)

        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Extracted")
            pd.DataFrame(debug_log).to_excel(writer, index=False, sheet_name="Debug Log")

        st.download_button(
            "📥 Download Excel", buffer.getvalue(), "royal_wine_data_first_page.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
