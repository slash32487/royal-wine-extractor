# Royal Wine PDF to Excel Extractor

This Streamlit app extracts structured data from Royal Wine PDF price lists and converts it to Excel. It identifies regions, brands, product names, item numbers, sizes, pricing, and discounts using layout-based logic (fonts, formatting, etc.).

## Features
- Automatically detects region and brand hierarchy
- Combines multi-line product names and awards
- Skips combo/gift packs and irrelevant headers/footers
- Flags inferred product names when missing

## Installation
Create a virtual environment and install dependencies:

```bash
pip install -r requirements.txt
