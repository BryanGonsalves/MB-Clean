from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Dict, Optional, Tuple

import streamlit as st

from cleaning import (clean_workbook, export_cleaned_workbook,
                      load_uploaded_data)

PAGE_TITLE = "MB-Clean"
THEME_NAVY = "#0A1D44"


def _ordinal(n: int) -> str:
    suffix = "th"
    if n % 100 not in (11, 12, 13):
        if n % 10 == 1:
            suffix = "st"
        elif n % 10 == 2:
            suffix = "nd"
        elif n % 10 == 3:
            suffix = "rd"
    return f"{n}{suffix}"


def _format_export_artifacts(start: Optional[date], end: Optional[date]) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    if not start or not end:
        return None, None, None
    if end < start:
        start, end = end, start

    if start.month == end.month and start.year == end.year:
        file_range = f"{start.strftime('%B')} {start.day} - {end.day}"
        sheet_range = f"{start.day} - {end.day} {end.strftime('%b')}"
    else:
        file_range = f"{start.strftime('%B %d')} - {end.strftime('%B %d')}"
        sheet_range = f"{start.strftime('%d %b')} - {end.strftime('%d %b')}"

    summary_range = f"{_ordinal(start.day)} to {_ordinal(end.day)} {end.strftime('%B')}"

    filename = f"Khotwa_Missed Sessions Report_ ({file_range}).xlsx"
    sheet_name = f"Missed Session ({sheet_range})"
    summary_title = f"Missed Mentor Sessions - {summary_range}"
    return filename, sheet_name, summary_title


def configure_page() -> None:
    st.set_page_config(page_title=PAGE_TITLE, page_icon="🧼", layout="centered")

    st.markdown(
        f"""
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Gopher:wght@400;600;700&display=swap');
            html, body, [data-testid="stAppViewContainer"], [data-testid="block-container"] {{
                font-family: 'Gopher', sans-serif;
                font-size: 10pt;
                color: #E2E8F0;
                background-color: #050505;
            }}
            .app-shell {{
                background-color: transparent;
                border-radius: 0;
                padding: 1.5rem 0;
                margin: 1rem auto;
                max-width: 820px;
                box-shadow: none;
            }}
            .app-subhead {{
                text-align: center;
                color: #94A3B8;
                font-size: 0.95rem;
                margin-bottom: 1rem;
            }}
            .section-title {{
                font-weight: 600;
                color: #F8FAFC;
                margin-top: 1.5rem;
                margin-bottom: 0.5rem;
                text-transform: uppercase;
                letter-spacing: 0.08rem;
                font-size: 0.85rem;
            }}
            .upload-card {{
                background-color: transparent;
                border: none;
                border-radius: 0;
                padding: 0;
                min-height: auto;
                display: flex;
                flex-direction: column;
                gap: 0.5rem;
            }}
            .upload-title {{
                font-weight: 600;
                font-size: 1rem;
                color: #F8FAFC;
            }}
            .upload-body {{
                color: #94A3B8;
                font-size: 0.9rem;
                line-height: 1.4;
            }}
            .summary-box {{
                border: 1px solid rgba(148, 163, 184, 0.25);
                padding: 1rem 1.25rem;
                margin-top: 0.75rem;
                color: #F8FAFC;
                background-color: #0B101A;
                border-radius: 12px;
                text-align: center;
            }}
            .summary-line {{
                margin: 0.2rem 0;
                line-height: 1.45;
            }}
            .notice {{
                border: 1px solid rgba(148, 163, 184, 0.4);
                padding: 0.9rem 1rem;
                margin: 1rem 0;
                color: #F8FAFC;
                background-color: #0B101A;
                border-radius: 12px;
                text-align: center;
            }}
            .footer {{
                text-align: center;
                color: #64748B;
                margin-top: 1.5rem;
                font-size: 10pt;
            }}
            div.stButton > button, div[data-testid="stDownloadButton"] > button {{
                background-color: #1E293B;
                color: #FFFFFF;
                border: none;
                border-radius: 999px;
                padding: 0.65rem 1.75rem;
                font-weight: 600;
                letter-spacing: 0.01rem;
            }}
            div.stButton > button:hover, div[data-testid="stDownloadButton"] > button:hover {{
                background-color: #020617;
            }}
            [data-testid="stFileUploaderDropzone"] {{
                border: 1px dashed rgba(148, 163, 184, 0.25);
                background-color: transparent;
            }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_footer() -> None:
    st.markdown('<div class="footer">Internal Use Only – Mindbase</div>', unsafe_allow_html=True)


def render_summary(summaries: Dict[str, Dict[str, int]], filename: str | None, start: Optional[date], end: Optional[date]) -> None:
    """Render an email template containing the summary counts and the report period.

    Displays a copyable text area with a ready-to-send email and includes the
    total number of rows generated in the cleaned report.
    """

    total = summaries.get("TOTAL", {})
    final_rows = total.get("final_rows", 0)

    # Build a human-friendly period string using the existing ordinal helper.
    if start and end:
        s, e = (start, end) if start <= end else (end, start)
        period = f"{s.strftime('%B')} {_ordinal(s.day)} to {_ordinal(e.day)}"
    else:
        period = "the selected period"

    file_label = filename or "Missed Session Report"
    # Use only the stem of the filename for a cleaner subject line if possible.
    try:
        from pathlib import Path

        file_label = Path(file_label).stem
    except Exception:
        pass

    subject = f"{file_label} ({period}) + Weekly Warning Letter Report"

    email_lines = [
        subject,
        "Hi Jay,",
        "",
        f"Please find attached the Missed Session Report for the period of {period}.",
        f"Rows included: {final_rows}",
        "",
        "I’ve added two additional columns:",
        "Column F: Entry labels received from mentors.",
        "Column E: Standardized labels formatted to align with ADEK’s preferences.",
        "Feel free to use or adjust them as needed.",
        "Also, please find the weekly warning letter submissions attached.",
        "Let me know if you need any further clarification.",
        "Kind regards,",
    ]

    email_text = "\n".join(email_lines)

    st.markdown('<div class="section-title">Email Template</div>', unsafe_allow_html=True)
    # Use a textarea so users can quickly copy the full email body.
    st.text_area("Copy-ready email", value=email_text, height=240)


def main() -> None:
    configure_page()

    if "export_bytes" not in st.session_state:
        st.session_state["export_bytes"] = None
    if "summaries" not in st.session_state:
        st.session_state["summaries"] = None
    if "source_name" not in st.session_state:
        st.session_state["source_name"] = None

    st.markdown('<div class="app-shell">', unsafe_allow_html=True)
    st.markdown(
        '<div class="app-subhead"><strong>MB-Clean</strong> · Generate the weekly missed-session report with one upload.</div>',
        unsafe_allow_html=True,
    )

    st.markdown('<div class="section-title">Report Window</div>', unsafe_allow_html=True)
    date_col_start, date_col_end = st.columns(2, gap="large")
    with date_col_start:
        week_start = st.date_input("Week start", value=None, key="week-start")
    with date_col_end:
        week_end = st.date_input("Week end", value=None, key="week-end")

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

    st.markdown('<div class="section-title">Meeting Report</div>', unsafe_allow_html=True)
    st.markdown(
        "<div class='upload-body'>Upload the meeting report to populate the reason for missed session from meeting notes.</div>",
        unsafe_allow_html=True,
    )
    meeting_file = st.file_uploader(
        "Meeting report",
        type=["xlsx", "xls", "csv"],
        key="meeting-upload",
        label_visibility="collapsed",
    )

    file_name_override, sheet_name_override, summary_title = _format_export_artifacts(week_start, week_end)

    if missed_file is not None and master_file is not None and meeting_file is not None:
        try:
            with st.spinner("Building report…"):
                missed_sheets = load_uploaded_data(missed_file)
                master_sheets = load_uploaded_data(master_file)
                meeting_sheets = load_uploaded_data(meeting_file)
                cleaned_sheets, summaries = clean_workbook(
                    missed_sheets=missed_sheets,
                    master_sheets=master_sheets,
                    meeting_sheets=meeting_sheets,
                    report_sheet_name=sheet_name_override,
                )
                export_bytes = export_cleaned_workbook(cleaned_sheets, summary_title=summary_title)
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
            st.session_state["source_name"] = file_name_override or missed_file.name
    elif missed_file is not None or master_file is not None or meeting_file is not None:
        st.markdown(
            "<div class='notice'>Please upload all three files (missed session export, master roster, meeting report) to generate the report.</div>",
            unsafe_allow_html=True,
        )

    if st.session_state["summaries"]:
        # Display a copy-ready email template instead of the original summary box.
        # Pass the stored source filename and the selected week range to render the period.
        render_summary(
            st.session_state["summaries"],
            st.session_state.get("source_name", file_name_override),
            week_start,
            week_end,
        )
        st.markdown('<div class="section-title">Download</div>', unsafe_allow_html=True)
        stored_name = st.session_state["source_name"] or "cleaned"
        if stored_name.lower().endswith(".xlsx"):
            download_name = stored_name
        else:
            download_name = f"{Path(stored_name).stem}_cleaned.xlsx"
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
