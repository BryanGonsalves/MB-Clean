MB-Clean
=======

MB-Clean is a single-page Streamlit application for Mindbase teams to standardise customer data. Upload an Excel workbook or CSV export, let the app tidy each sheet automatically, and download a formatted `.xlsx` file that complies with internal reporting rules.

Features
--------
- Removes duplicate rows after trimming whitespace and fixing casing discrepancies.
- Title-cases name columns, normalises dates to ISO 8601, and validates contact details.
- Supports multi-sheet workbooks, preserving original sheet names while applying identical formatting.
- Generates a navy-themed workbook with wrapped text, bold headers, gridlines, and a one-cell margin.
- Optional OpenAI-assisted clean-up for stubborn name and date values (disabled by default).

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
2. Wait while the app shows “Cleaning in progress…” and performs duplicate removal, trimming, name normalisation, date conversion, and contact validation across all sheets.
3. Review the summary message that lists duplicate removals, invalid contact drops, and final row counts.
4. Use the **Download Cleaned File** button to save the cleaned workbook with standard formatting starting at cell B2.

AI Normalisation (Optional)
---------------------------
- Toggle **Enable AI Normalization for Names and Dates** in the sidebar to allow OpenAI to refine values the deterministic rules could not fix.
- Provide the API key via `OPENAI_API_KEY` in Streamlit secrets or environment variables. If the key is absent, the toggle is ignored and the deterministic pipeline still runs.

Deployment Notes
----------------
- Streamlit Community Cloud: point the app to `app.py`, select Python 3.11, and supply secrets if needed.
- OPENAI_API_KEY can be left unset; the application functions entirely offline without it.
