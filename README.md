# Royal Wine ETS Excel Extractor

This Streamlit app extracts structured product data from Excel sheets generated from ETS for Windows (via Adobe Acrobat Pro export).

## Features
- Accepts `.xlsx` files exported from Acrobat Pro
- Identifies:
  - Product Name
  - Item#
  - Vintage (1990–2025)
  - Bottle Size & Quantity
  - Case/Bottle Price
  - Discounts
- Debug log for raw row preview
- Export to clean Excel file

## How to Use
```bash
pip install -r requirements.txt
streamlit run app.py
