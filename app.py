import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
from io import BytesIO

st.title("Royal Wine PDF to Excel Extractor")

uploaded_file = st.file_uploader("Upload Royal Wine PDF", type="pdf")

if uploaded_file:
    pdf_bytes = uploaded_file.read()
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    data = []
    current_region = None
    current_brand = None

    for page in doc:
        words = page.get_text("words")  # (x0, y0, x1, y1, word, block_no, line_no, word_no)
        lines = {}

        for w in words:
            y = round(w[1], 1)
            text = w[4].strip()
            if not text:
                continue
            lines.setdefault(y, []).append((w[0], text))

        for y in sorted(lines.keys()):
            line = " ".join([w[1] for w in sorted(lines[y], key=lambda x: x[0])])

            if re.search(r"^\s*Item#", line, re.IGNORECASE):
                continue

            if re.fullmatch(r"[A-Z\s\-&]+", line) and len(line.split()) <= 5:
                current_region = line.strip()
                continue

            if re.match(r"^[A-Z][a-z]+(?: [A-Z][a-z]+)*$", line) and len(line.split()) <= 4:
                current_brand = line.strip()
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
                # backtrack for product name
                prior_lines = sorted([k for k in lines if k < y], reverse=True)
                pname = []
                for py in prior_lines[:3]:
                    pt = " ".join([w[1] for w in sorted(lines[py], key=lambda x: x[0])])
                    if not re.search(r"\d{5}|\d+ /\d+|\d+\.\d{2}|cs", pt):
                        pname.insert(0, pt)
                item["Product Name"] = " ".join(pname)

                # forward grab discounts
