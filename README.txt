GSTR-3B Consolidator – Setup & Usage
======================================

FILES:
  gstr3b_app.py     – Streamlit app (main UI)
  gstr3b_parser.py  – PDF extraction engine
  gstr3b_excel.py   – Excel workbook builder

INSTALL DEPENDENCIES:
  pip install streamlit pdfplumber openpyxl

RUN THE APP:
  streamlit run gstr3b_app.py

FEATURES:
  • Upload any number of GSTR-3B PDFs at once
  • Extracts all 6 sections with zero data loss
  • Generates a 4-sheet consolidated Excel workbook:
      Sheet 1 – Outward Supplies (3.1 / 3.1.1 / 3.2)
      Sheet 2 – Eligible ITC (Section 4)
      Sheet 3 – Exempt Inward + Interest (Section 5 & 5.1)
      Sheet 4 – Payment of Tax (Section 6.1)
  • SUM formulas for annual totals
  • Colour-coded headers matching your sample format
  • Handles PDF number-splitting artefacts automatically

SUPPORTED PDFs:
  Standard GSTR-3B PDFs downloaded from GST portal (any GSTIN)
  Tested on: Kapston Services Limited – GSTIN 36AADCK5952F1ZH
