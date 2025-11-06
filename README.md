MB-Clean
=======

MB-Clean is a single-page Streamlit application for Mindbase teams to standardise customer data. Upload an Excel workbook or CSV export, let the app tidy each sheet automatically, and download a formatted `.xlsx` file that complies with internal reporting rules.

Features
--------
- Removes duplicate rows after trimming whitespace and fixing casing discrepancies.
- Title-cases columns whose headers contain “name”, normalises “date” columns to ISO 8601, and validates contact details.
- Supports multi-sheet workbooks, preserving original sheet names while applying identical formatting.
- Generates a navy-themed workbook with wrapped text, bold headers, gridlines, and a one-cell margin.

Project Structure
-----------------
```
MB-Clean/
├── app.py
├── cleaning/
│   ├── __init__.py
│   ├── ai_normalize.py
│   └── detect_fix.py
├── requirements.txt
├── README.md
└── sample_data/
    └── sample.xlsx
```

Getting Started
---------------
1. Create and activate a Python 3.11 virtual environment.
2. Install dependencies: `pip install -r requirements.txt`.
3. Launch the app: `streamlit run app.py`.

Using MB-Clean
--------------
1. Upload an `.xlsx`, `.xls`, or `.csv` file.
2. Wait while the app shows “Cleaning in progress…” and performs duplicate removal, trimming, name normalisation, “date” column conversion, and contact validation across all sheets.
3. Review the summary message that lists duplicate removals, cleared contact fields, and final row counts.
4. Use the **Download Cleaned File** button to save the cleaned workbook with standard formatting starting at cell B2.

Deployment Notes
----------------
- Streamlit Community Cloud: point the app to `app.py`, select Python 3.11, and deploy.
- No secrets are required; the application runs fully offline.
