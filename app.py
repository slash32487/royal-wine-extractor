import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
from io import BytesIO
from collections import Counter

st.title("Royal Wine PDF to Excel Extractor")

# Predefined region and brand databases
known_regions = {
    "CALIFORNIA", "FRANCE", "BORDEAUX", "BOURGOGNE", "ISRAEL", "ITALY",
    "NEW ZEALAND", "SOUTH AFRICA", "SPAIN", "CHAMPAGNE", "COTES DU RHONE",
    "MARGAUX", "MEDOC", "PAUILLAC", "POMEROL", "SAUTERNES", "ST. EMILION",
    "ST. ESTEPHE", "COTES DE PROVENCE"
}

known_brands = {
    "HAGAFEN CELLARS", "HAJDU", "MARCIANO ESTATE", "PADIS VINEYARDS",
    "SONOMA LOEB", "STOUDEMIRE", "WEINSTOCK", "WEINSTOCK - BY W", "WEINSTOCK-CELLAR SELECT",
    "BOKOBSA SELECTIONS", "ROLLAN DE BY", "PHILIPPE LE HARDI", "DOMAINE TERNYNCK",
    "ANOMIS", "B SAINT BEATRICE", "CHATEAU ROUBINE", "NADIV WINERY",
    "DOMAINE DU CASTEL", "RAZIEL BY CASTEL", "YATIR WINERY", "CARMEL", "SHILOH",
    "NETOFA", "BINNUN", "CHATEAU GOLAN", "MATAR", "NANA WINERY", "TEPERBERG",
    "TULIP", "TZUBA", "VITKIN", "CANTINA GIULIANO", "LA REGOLA", "MASSERIA",
    "TERRA DI SETA", "ESSA", "RIMAPERE", "CAPCANES", "RAMON CARDOVA",
    "RASHI WINES", "BEN AMI", "KING DAVID"
}

uploaded_file = st.file_uploader("Upload Royal Wine PDF", type="pdf")

if uploaded_file:
    pdf_bytes = uploaded_file.read()
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    data = []
    debug_log = []
    current_region = None
    current_brand = None

    all_regions = set()
    all_brands = set()
    item_count = 0
    missing_data_rows = []
    failed_parses = []

    for page in doc:
        words = page.get_text("words")
        lines = {}

        for w in words:
            y = round(w[1], 1)
            text = w[4].strip()
            if not text:
                continue
            lines.setdefault(y, []).append((w[0], text))

        sorted_y = sorted(lines.keys())
        for idx, y in enumerate(sorted_y):
            line_words = sorted(lines[y], key=lambda x: x[0])
            line = " ".join([w[1] for w in line_words])

            debug_log.append({"Y": y, "Text": line})

            if re.search(r"^\s*Item#", line, re.IGNORECASE):
                continue

            if line.strip() in known_regions:
                current_region = line.strip()
                all_regions.add(current_region)
                continue

            if line.strip() in known_brands:
                current_brand = line.strip()
                all_brands.add(current_brand)
                continue

            match = re.match(r"(\d{5})\s+(\d{4}|NV)?\s+(\d+ /\d+[A-Z]*)\s+(\d+\.\d{2})\s+(\d+\.\d{2})", line)
            if match:
                item = {
                    "Region": current_region,
                    "Brand": current_brand,
                    "Item#": match.group(1),
                    "Vintage": match.group(2) or "",
                    "Size": match.group(3),
                    "Case Price": match.group(4),
                    "Bottle Price": match.group(5),
                    "Product Name": "",
                    "Discounts": ""
                }

                pname = []
                for back_y in sorted_y[max(0, idx - 3):idx][::-1]:
                    pt = " ".join([w[1] for w in sorted(lines[back_y], key=lambda x: x[0])])
                    if not re.search(r"\d{5}|\d+ /\d+|\d+\.\d{2}|cs", pt):
                        pname.insert(0, pt)
                item["Product Name"] = " ".join(pname)

                for fy in sorted_y[idx + 1:idx + 4]:
                    ft = " ".join([w[1] for w in sorted(lines[fy], key=lambda x: x[0])])
                    if re.match(r"\$\d+\.\d{2} on \d+cs", ft):
                        item["Discounts"] += (ft + "; ")
                item["Discounts"] = item["Discounts"].strip("; ")

                try:
                    size_parts = item["Size"].split("/")
                    item["Bottles per Case"] = size_parts[0].strip()
                    item["Bottle Size"] = size_parts[1].strip()
                except Exception as e:
                    item["Bottles per Case"] = ""
                    item["Bottle Size"] = ""

                del item["Size"]
                item_count += 1

                if not all([item["Region"], item["Brand"], item["Product Name"]]):
                    missing_data_rows.append(item)

                data.append(item)
            elif re.match(r"\d{5}", line):
                failed_parses.append(line)

    if not data:
        st.warning("No data found.")
    else:
        df = pd.DataFrame(data)
        st.success(f"Extraction complete! {item_count} items extracted.")
        st.dataframe(df)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="WineData")
            if missing_data_rows:
                pd.DataFrame(missing_data_rows).to_excel(writer, index=False, sheet_name="Missing Fields")
            if failed_parses:
                pd.DataFrame(failed_parses, columns=["Failed Line"]).to_excel(writer, index=False, sheet_name="Failed Matches")
            pd.DataFrame(debug_log).to_excel(writer, index=False, sheet_name="Debug Log")

        st.download_button(
            label="Download Excel File",
            data=output.getvalue(),
            file_name="royal_wine_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.subheader("Regions Found")
        st.write(sorted(list(all_regions)))

        st.subheader("Brands Found")
        st.write(sorted(list(all_brands)))

        if missing_data_rows:
            st.warning(f"⚠️ {len(missing_data_rows)} items are missing region, brand, or product name")
            st.dataframe(pd.DataFrame(missing_data_rows))

        if failed_parses:
            st.error(f"❌ {len(failed_parses)} lines matched item# but failed to parse fully.")
            st.dataframe(pd.DataFrame(failed_parses, columns=["Failed Line"]))
