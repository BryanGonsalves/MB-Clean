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
                color: #0F172A;
                background-color: #F4F6FB;
            }}
            .app-shell {{
                background-color: #FFFFFF;
                border-radius: 16px;
                padding: 2rem 2.5rem;
                margin: 2rem auto;
                max-width: 820px;
                box-shadow: 0 20px 45px rgba(15, 23, 42, 0.08);
            }}
            .app-header {{
                color: {THEME_NAVY};
                font-weight: 700;
                text-align: center;
                font-size: 1.5rem;
                margin-bottom: 0.4rem;
                letter-spacing: 0.02rem;
            }}
            .app-subhead {{
                text-align: center;
                color: #475569;
                font-size: 0.95rem;
                margin-bottom: 1.5rem;
            }}
            .section-title {{
                font-weight: 600;
                color: #0F172A;
                margin-top: 1.5rem;
                margin-bottom: 0.5rem;
                text-transform: uppercase;
                letter-spacing: 0.04rem;
                font-size: 0.85rem;
            }}
            .upload-card {{
                background-color: #F8FAFF;
                border: 1px solid rgba(10, 29, 68, 0.12);
                border-radius: 12px;
                padding: 1rem;
                min-height: 230px;
                display: flex;
                flex-direction: column;
                gap: 0.5rem;
            }}
            .upload-title {{
                font-weight: 600;
                font-size: 1rem;
                color: #0F172A;
            }}
            .upload-body {{
                color: #475569;
                font-size: 0.9rem;
                line-height: 1.4;
            }}
            .summary-box {{
                border: 1px solid rgba(10, 29, 68, 0.18);
                padding: 1rem 1.25rem;
                margin-top: 0.75rem;
                color: #0F172A;
                background-color: #FFFFFF;
                border-radius: 12px;
                text-align: center;
            }}
            .summary-line {{
                margin: 0.2rem 0;
                line-height: 1.45;
            }}
            .notice {{
                border: 1px solid rgba(10, 29, 68, 0.3);
                padding: 0.9rem 1rem;
                margin: 1rem 0;
                color: #0F172A;
                background-color: #FFFFFF;
                border-radius: 12px;
                text-align: center;
            }}
            .notice.error {{
                border-style: dashed;
            }}
            .footer {{
                text-align: center;
                color: #475569;
                margin-top: 1.5rem;
                font-size: 10pt;
            }}
            div.stButton > button, div[data-testid="stDownloadButton"] > button {{
                background-color: {THEME_NAVY};
                color: #FFFFFF;
                border: none;
                border-radius: 999px;
                padding: 0.65rem 1.75rem;
                font-weight: 600;
                letter-spacing: 0.01rem;
            }}
            div.stButton > button:hover, div[data-testid="stDownloadButton"] > button:hover {{
                background-color: #020817;
            }}
            [data-testid="stFileUploaderDropzone"] {{
                border: 1px dashed rgba(10, 29, 68, 0.35);
                background-color: #FFFFFF;
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

    html = "".join(f"<div class='summary-line'>{line}</div>" for line in primary_lines)
    if detail_lines:
        html += f"<hr style='border: none; border-top: 1px solid {THEME_NAVY}; margin: 0.6rem 0;'/>"
        html += "".join(f"<div class='summary-line'>{line}</div>" for line in detail_lines)

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
    st.markdown(
        '<div class="app-subhead">Generate the weekly missed-session report with one upload.</div>',
        unsafe_allow_html=True,
    )

    st.markdown('<div class="section-title">Upload Files</div>', unsafe_allow_html=True)
    col_export, col_master = st.columns(2, gap="large")
    with col_export:
        st.markdown("<div class='upload-card'>", unsafe_allow_html=True)
        st.markdown("<div class='upload-title'>Missed Session Export</div>", unsafe_allow_html=True)
        st.markdown(
            "<div class='upload-body'>Upload the weekly missed-session export exactly as you download it.</div>",
            unsafe_allow_html=True,
        )
        missed_file = st.file_uploader(
            "Missed session export",
            type=["xlsx", "xls", "csv"],
            key="missed-upload",
            label_visibility="collapsed",
        )
        st.markdown("</div>", unsafe_allow_html=True)

    with col_master:
        st.markdown("<div class='upload-card'>", unsafe_allow_html=True)
        st.markdown("<div class='upload-title'>Master Student Data</div>", unsafe_allow_html=True)
        st.markdown(
            "<div class='upload-body'>Upload the master roster that contains PS Number, mentor, and ADEK advisor.</div>",
            unsafe_allow_html=True,
        )
        master_file = st.file_uploader(
            "Master student data",
            type=["xlsx", "xls", "csv"],
            key="master-upload",
            label_visibility="collapsed",
        )
        st.markdown("</div>", unsafe_allow_html=True)

    if missed_file is not None and master_file is not None:
        try:
            with st.spinner("Building report…"):
                missed_sheets = load_uploaded_data(missed_file)
                master_sheets = load_uploaded_data(master_file)
                cleaned_sheets, summaries = clean_workbook(
                    missed_sheets,
                    master_sheets,
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
            st.session_state["source_name"] = missed_file.name
    elif missed_file is not None or master_file is not None:
        st.markdown(
            "<div class='notice'>Please upload both files to generate the report.</div>",
            unsafe_allow_html=True,
        )

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
