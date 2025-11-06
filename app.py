from __future__ import annotations

import os
from pathlib import Path
from typing import Dict, Optional

import streamlit as st

from cleaning import clean_workbook, export_cleaned_workbook, load_uploaded_data
from cleaning.ai_normalize import AINormalizer, AIUnavailableError

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
                border: 3px solid {THEME_NAVY};
                padding: 1.5rem 2rem;
                margin: 1.5rem auto;
                max-width: 900px;
                background-color: #FFFFFF;
            }}
            .app-header {{
                background-color: {THEME_NAVY};
                color: #FFFFFF;
                font-weight: 700;
                text-align: center;
                padding: 1rem;
                margin-bottom: 1.5rem;
                letter-spacing: 0.05rem;
            }}
            .section-title {{
                font-weight: 600;
                color: #000000;
                margin-top: 1.5rem;
                margin-bottom: 0.5rem;
            }}
            .body-text {{
                margin-bottom: 0.75rem;
            }}
            .summary-box {{
                border: 1px solid {THEME_NAVY};
                padding: 1rem;
                margin-top: 1rem;
                color: #000000;
                background-color: #FFFFFF;
            }}
            .status-message {{
                border: 1px solid {THEME_NAVY};
                padding: 0.75rem;
                margin-top: 1rem;
                color: #000000;
                background-color: #FFFFFF;
            }}
            .status-message.error {{
                border-style: dashed;
            }}
            .footer {{
                text-align: center;
                color: #000000;
                margin-top: 2rem;
                font-size: 10pt;
            }}
            div.stButton > button, div[data-testid="stDownloadButton"] > button {{
                background-color: {THEME_NAVY};
                color: #FFFFFF;
                border: 1px solid {THEME_NAVY};
                border-radius: 4px;
                padding: 0.6rem 1.5rem;
                font-weight: 600;
            }}
            div.stButton > button:hover, div[data-testid="stDownloadButton"] > button:hover {{
                border-color: #000000;
            }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def resolve_ai_normalizer(enable_ai: bool) -> tuple[Optional[AINormalizer], Optional[str]]:
    if not enable_ai:
        return None, None

    api_key = None
    if "OPENAI_API_KEY" in st.secrets:
        api_key = st.secrets["OPENAI_API_KEY"]
    if not api_key:
        api_key = os.getenv("OPENAI_API_KEY")

    try:
        if api_key:
            return AINormalizer.from_api_key(api_key), None
        return AINormalizer.from_env(), None
    except AIUnavailableError as err:
        return None, str(err)
    except Exception:
        return None, "AI normalization is currently unavailable."


def render_footer() -> None:
    st.markdown('<div class="footer">Internal Use Only – Mindbase</div>', unsafe_allow_html=True)


def render_summary(summaries: Dict[str, Dict[str, int]]) -> None:
    total = summaries.get("TOTAL", {})
    duplicates = total.get("duplicates_removed", 0)
    invalids = total.get("invalid_rows_removed", 0)
    final_rows = total.get("final_rows", 0)

    messages = [
        f"{duplicates} duplicate rows removed.",
        f"{invalids} rows dropped due to invalid contact information.",
        f"{final_rows} rows ready for download.",
    ]

    detail_lines = []
    for sheet_name, metrics in summaries.items():
        if sheet_name == "TOTAL":
            continue
        detail_lines.append(
            f"{sheet_name}: {metrics.get('duplicates_removed', 0)} duplicates removed, "
            f"{metrics.get('invalid_rows_removed', 0)} invalid rows dropped, "
            f"{metrics.get('final_rows', 0)} rows cleaned."
        )

    summary_text = " ".join(messages)
    details_text = " ".join(detail_lines)

    st.markdown('<div class="section-title">Summary</div>', unsafe_allow_html=True)
    st.markdown(
        f'<div class="summary-box">{summary_text}'
        f"{(' ' + details_text) if details_text else ''}</div>",
        unsafe_allow_html=True,
    )


def main() -> None:
    configure_page()

    if "export_bytes" not in st.session_state:
        st.session_state["export_bytes"] = None
    if "summaries" not in st.session_state:
        st.session_state["summaries"] = None
    if "source_name" not in st.session_state:
        st.session_state["source_name"] = None

    st.sidebar.markdown("### Options")
    ai_toggle = st.sidebar.toggle("Enable AI Normalization for Names and Dates", value=False)
    ai_normalizer, ai_warning = resolve_ai_normalizer(ai_toggle)
    if ai_warning:
        st.sidebar.markdown(f'<div class="status-message">{ai_warning}</div>', unsafe_allow_html=True)

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
                    enable_ai=ai_toggle and ai_normalizer is not None,
                    ai_normalizer=ai_normalizer,
                )
                export_bytes = export_cleaned_workbook(cleaned_sheets)
        except Exception as exc:
            st.session_state["export_bytes"] = None
            st.session_state["summaries"] = None
            st.session_state["source_name"] = None
            st.markdown(
                f'<div class="status-message error">Unable to clean the file: {exc}</div>',
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
