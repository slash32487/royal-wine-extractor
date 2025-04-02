import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
from io import BytesIO
from collections import Counter

st.title("Royal Wine PDF to Excel Extractor (4-Column Split Mode)")

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

uploaded_file = st.file_uploader("Upload 4-Column Split PDF", type="pdf")

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
        text = page.get_text("text")
        lines = text.split("\n")

        for idx, line in enumerate(lines):
            debug_log.append({"Line": line})
            line = line.strip()
            if not line:
                continue

            if re.search(r"^\s*Item#", line, re.IGNORECASE):
                continue

            if line in known_regions:
                current_region = line
                all_regions.add(current_region)
                continue

            if line in known_brands:
                current_brand = line
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

                # Look backward for product name (up to 4 lines)
                pname = []
                for back in range(max(0, idx - 4), idx):
                    pt = lines[back].strip()
                    if pt and not re.search(r"\d{5}|\d+ /\d+|\d+\.\d{2}|cs", pt):
                        pname.append(pt)
                item["Product Name"] = " ".join(pname)

                # Look ahead for discount(s)
                for fwd in range(idx + 1, min(idx + 6, len(lines))):
                    ft = lines[fwd].strip()
                    if re.match(r"\$\d+\.\d{2} on \d+cs", ft):
                        item["Discounts"] += ft + "; "

                item["Discounts"] = item["Discounts"].strip("; ")

                try:
                    size_parts = item["Size"].split("/")
                    item["Bottles per Case"] = size_parts[0].strip()
                    item["Bottle Size"] = size_parts[1].strip()
                except:
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
