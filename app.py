from __future__ import annotations

from pathlib import Path
from typing import Dict

import streamlit as st

from cleaning import clean_workbook, export_cleaned_workbook, load_uploaded_data

PAGE_TITLE = "MB-Clean"
THEME_NAVY = "#0A1D44"


def configure_page() -> None:
    st.set_page_config(page_title=PAGE_TITLE, page_icon="🧼", layout="centered")

    st.markdown(
        f"""
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Gopher:wght@400;600;700&display=swap');
            html, body, [data-testid="stAppViewContainer"], [data-testid="block-container"] {{
                font-family: 'Gopher', sans-serif;
                font-size: 10pt;
                color: #000000;
                background-color: #FFFFFF;
            }}
            .app-shell {{
                border: 2px solid {THEME_NAVY};
                padding: 1.5rem 2rem;
                margin: 2rem auto;
                max-width: 760px;
                background-color: #FFFFFF;
                box-shadow: 0 0 0 4px rgba(10, 29, 68, 0.06);
            }}
            .app-header {{
                background-color: {THEME_NAVY};
                color: #FFFFFF;
                font-weight: 700;
                text-align: center;
                padding: 0.9rem 1rem;
                margin-bottom: 1.25rem;
                letter-spacing: 0.05rem;
                border-radius: 2px;
            }}
            .section-title {{
                font-weight: 600;
                color: #000000;
                margin-top: 1.25rem;
                margin-bottom: 0.25rem;
                text-transform: uppercase;
                letter-spacing: 0.04rem;
            }}
            .body-text {{
                margin-bottom: 0.75rem;
                color: #000000;
                line-height: 1.5;
            }}
            .summary-box {{
                border: 1px solid {THEME_NAVY};
                padding: 1rem 1.25rem;
                margin-top: 0.75rem;
                color: #000000;
                background-color: #FFFFFF;
                border-radius: 2px;
            }}
            .summary-box ul {{
                margin: 0.25rem 0 0 1rem;
                padding: 0;
                line-height: 1.5;
            }}
            .notice {{
                border: 1px solid {THEME_NAVY};
                padding: 0.75rem 1rem;
                margin: 1rem 0;
                color: #000000;
                background-color: #FFFFFF;
                border-radius: 2px;
            }}
            .notice.error {{
                border-style: dashed;
            }}
            .footer {{
                text-align: center;
                color: #000000;
                margin-top: 1.5rem;
                font-size: 10pt;
            }}
            div.stButton > button, div[data-testid="stDownloadButton"] > button {{
                background-color: {THEME_NAVY};
                color: #FFFFFF;
                border: 1px solid {THEME_NAVY};
                border-radius: 4px;
                padding: 0.55rem 1.5rem;
                font-weight: 600;
                letter-spacing: 0.02rem;
            }}
            div.stButton > button:hover, div[data-testid="stDownloadButton"] > button:hover {{
                border-color: #000000;
            }}
            [data-testid="stFileUploaderDropzone"] {{
                border: 1px solid {THEME_NAVY};
            }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_footer() -> None:
    st.markdown('<div class="footer">Internal Use Only – Mindbase</div>', unsafe_allow_html=True)


def render_summary(summaries: Dict[str, Dict[str, int]]) -> None:
    total = summaries.get("TOTAL", {})
    duplicates = total.get("duplicates_removed", 0)
    invalids = total.get("invalid_contacts_cleared", 0)
    final_rows = total.get("final_rows", 0)

    primary_lines = [
        f"{duplicates} duplicate rows removed",
        f"{invalids} contact fields cleared",
        f"{final_rows} rows ready for download",
    ]

    detail_lines = [
        f"{sheet_name}: {metrics.get('final_rows', 0)} rows (duplicates removed {metrics.get('duplicates_removed', 0)}, contacts cleared {metrics.get('invalid_contacts_cleared', 0)})"
        for sheet_name, metrics in summaries.items()
        if sheet_name != "TOTAL"
    ]

    html = "<ul>" + "".join(f"<li>{line}</li>" for line in primary_lines) + "</ul>"
    if detail_lines:
        html += f"<hr style='border: none; border-top: 1px solid {THEME_NAVY}; margin: 0.6rem 0;'/>"
        html += "<ul>" + "".join(f"<li>{line}</li>" for line in detail_lines) + "</ul>"

    st.markdown('<div class="section-title">Summary</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="summary-box">{html}</div>', unsafe_allow_html=True)


def main() -> None:
    configure_page()

    if "export_bytes" not in st.session_state:
        st.session_state["export_bytes"] = None
    if "summaries" not in st.session_state:
        st.session_state["summaries"] = None
    if "source_name" not in st.session_state:
        st.session_state["source_name"] = None

    st.markdown('<div class="app-shell">', unsafe_allow_html=True)
    st.markdown('<div class="app-header">MB-Clean</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-title">Upload</div>', unsafe_allow_html=True)
    st.markdown('<div class="body-text">Upload an Excel or CSV file to begin cleaning.</div>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader(
        "Upload an Excel or CSV file",
        type=["xlsx", "xls", "csv"],
        label_visibility="collapsed",
    )

    if uploaded_file is not None:
        try:
            with st.spinner("Cleaning in progress…"):
                sheets = load_uploaded_data(uploaded_file)
                cleaned_sheets, summaries = clean_workbook(
                    sheets,
                )
                export_bytes = export_cleaned_workbook(cleaned_sheets)
        except Exception as exc:
            st.session_state["export_bytes"] = None
            st.session_state["summaries"] = None
            st.session_state["source_name"] = None
            st.markdown(
                f'<div class="notice error">Unable to clean the file: {exc}</div>',
                unsafe_allow_html=True,
            )
        else:
            st.session_state["export_bytes"] = export_bytes
            st.session_state["summaries"] = summaries
            st.session_state["source_name"] = uploaded_file.name

    if st.session_state["summaries"]:
        render_summary(st.session_state["summaries"])
        st.markdown('<div class="section-title">Download</div>', unsafe_allow_html=True)
        download_name = f"{Path(st.session_state['source_name']).stem}_cleaned.xlsx"
        st.download_button(
            "Download Cleaned File",
            data=st.session_state["export_bytes"],
            file_name=download_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.markdown("</div>", unsafe_allow_html=True)
    render_footer()


if __name__ == "__main__":
    main()
