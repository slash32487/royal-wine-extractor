import streamlit as st
import pandas as pd
from doctr.io import DocumentFile
from doctr.models import ocr_predictor

st.set_page_config(page_title="PDF OCR Extractor with DocTR", layout="wide")
st.title("📄 PDF OCR Extractor (DocTR)")

uploaded_file = st.file_uploader("Upload a PDF file", type=["pdf"])

if uploaded_file:
    with st.spinner("Processing with DocTR OCR model..."):
        # Load and parse the PDF
        doc = DocumentFile.from_pdf(uploaded_file.read())
        model = ocr_predictor(pretrained=True)
        result = model(doc)

        # Parse all blocks of text with coordinates
        rows = []
        for page_idx, page in enumerate(result.pages):
            for block in page.blocks:
                for line in block.lines:
                    text = " ".join([word.value for word in line.words])
                    bbox = line.geometry
                    rows.append({
                        "Page": page_idx + 1,
                        "Text": text,
                        "X0": round(bbox[0][0], 4),
                        "Y0": round(bbox[0][1], 4),
                        "X1": round(bbox[1][0], 4),
                        "Y1": round(bbox[1][1], 4),
                    })

        df = pd.DataFrame(rows)
        st.success(f"Extracted {len(df)} text blocks across {len(result.pages)} page(s).")
        st.dataframe(df.head(100), use_container_width=True)

        # Download button
        csv = df.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Download CSV", csv, "ocr_extracted.csv", "text/csv")

        # Debug preview
        with st.expander("Show Full Table"):
            st.dataframe(df, use_container_width=True)
else:
    st.info("Upload a PDF file to begin.")
