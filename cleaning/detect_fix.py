"""Data loading, cleaning, and export helpers for the MB-Clean app."""

from __future__ import annotations

import io
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Tuple

import numpy as np
import pandas as pd

NA_COLOR = "#0A1D44"
HEADER_FONT_COLOR = "#FFFFFF"
BODY_FONT_COLOR = "#000000"
REPORT_COLUMNS = [
    "Sr. No.",
    "PS Number",
    "Student Full Name",
    "Standardized Entry Label",
    "Entry Label",
    "Date of missed scheduled meeting",
    "Reason for Missed Meeting",
    "Date of rescheduled advising session (if applicable)",
    "Mentor",
    "ADEK Advisor",
]
MISSED_REQUIRED = {
    "Student first": ["Student first"],
    "Student last": ["Student last"],
    "Entry Label": ["Entry Label", "Entry label"],
    "Date of missed scheduled meeting": ["Date of missed scheduled meeting", "Date"],
}
MISSED_OPTIONAL = {
    "Reason for Missed Meeting": ["Reason for Missed Meeting"],
    "Date of rescheduled advising session (if applicable)": [
        "Date of rescheduled advising session (if applicable)"
    ],
}
MASTER_REQUIRED = {
    "Student Full Name": ["Student Name"],
    "PS Number": ["PS Number", "ADEK Applicant ID", "Application Number", "SMS Account"],
    "Mentor": ["Mentor", "Current Mentor"],
    "ADEK Advisor": ["ADEK Advisor"],
}


@dataclass
class SheetSummary:
    """Holds per-sheet cleaning statistics."""

    sheet_name: str
    initial_rows: int
    duplicates_removed: int
    invalid_contacts_cleared: int
    final_rows: int

    @property
    def as_dict(self) -> Dict[str, int]:
        return {
            "initial_rows": self.initial_rows,
            "duplicates_removed": self.duplicates_removed,
            "invalid_contacts_cleared": self.invalid_contacts_cleared,
            "final_rows": self.final_rows,
        }


def load_uploaded_data(uploaded_file) -> Dict[str, pd.DataFrame]:
    """Load a user-uploaded file into a mapping of sheet name to DataFrame."""

    if not uploaded_file or not getattr(uploaded_file, "name", None):
        raise ValueError("A valid file upload is required.")

    filename = Path(uploaded_file.name)
    suffix = filename.suffix.lower()
    uploaded_file.seek(0)

    if suffix in {".xlsx", ".xls"}:
        excel = pd.ExcelFile(uploaded_file)
        data = {sheet: excel.parse(sheet) for sheet in excel.sheet_names}
    elif suffix == ".csv":
        data = {filename.stem or "Cleaned Data": pd.read_csv(uploaded_file)}
    else:
        raise ValueError("Unsupported file type. Please upload .xlsx, .xls, or .csv files.")

    uploaded_file.seek(0)
    return data


def _normalize_label(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value).lower())


def _resolve_column(
    df: pd.DataFrame,
    target_label: str,
    aliases: Iterable[str] | None = None,
    *,
    required: bool = True,
) -> str | None:
    candidates = list(aliases or []) + [target_label]
    normalized_targets = {_normalize_label(label) for label in candidates}
    for column in df.columns:
        if _normalize_label(column) in normalized_targets:
            return column
    if required:
        raise KeyError(f"Missing required column '{target_label}'.")
    return None


def _build_lookup_key(series: pd.Series) -> pd.Series:
    return (
        series.astype("string")
        .fillna("")
        .str.lower()
        .str.strip()
        .str.replace(r"\s+", " ", regex=True)
    )


def _build_full_name(first: pd.Series, last: pd.Series) -> pd.Series:
    first_part = first.astype("string").fillna("").str.strip()
    last_part = last.astype("string").fillna("").str.strip()
    combined = (first_part + " " + last_part).str.replace(r"\s+", " ", regex=True).str.strip()
    formatted = combined.map(_smart_title_case)
    return formatted.replace("", pd.NA)


def _standardize_entry_label(value: str) -> str:
    raw = (value or "").strip()
    lowered = raw.lower()

    if "no" not in lowered and "ns" not in lowered:
        return raw

    number = None
    num_match = re.search(r"(?:ns|no\s*show)\s*([0-9])", lowered)
    if num_match:
        number = num_match.group(1)

    tag = None
    if re.search(r"\btl\b|team\s*lead", lowered):
        tag = "TL"
    elif re.search(r"\baa\b", lowered):
        tag = "AA"

    parts = ["No Show"]
    if number:
        parts.append(number)
    if tag:
        parts.append(tag)

    return " ".join(parts)


def _build_missed_session_report(missed_df: pd.DataFrame, master_df: pd.DataFrame) -> pd.DataFrame:
    if missed_df.empty:
        return pd.DataFrame(columns=REPORT_COLUMNS)

    working = missed_df.copy()
    resolved_missed = {
        label: _resolve_column(working, label, aliases, required=True)
        for label, aliases in MISSED_REQUIRED.items()
    }
    optional_missed = {
        label: _resolve_column(working, label, aliases, required=False)
        for label, aliases in MISSED_OPTIONAL.items()
    }

    working["Student Full Name"] = _build_full_name(
        working[resolved_missed["Student first"]],
        working[resolved_missed["Student last"]],
    )

    entry_series = (
        working[resolved_missed["Entry Label"]]
        .astype("string")
        .fillna("")
        .str.strip()
    )
    include_mask = (
        entry_series.str.contains(r"(?i)\bno\s*show\b", na=False)
        | entry_series.str.contains(r"(?i)\bns\s*\d*\b", na=False)
    )
    working = working.loc[include_mask].reset_index(drop=True)

    if working.empty:
        return pd.DataFrame(columns=REPORT_COLUMNS)

    working["Standardized Entry Label"] = working[resolved_missed["Entry Label"]].map(_standardize_entry_label)

    report_df = pd.DataFrame(index=working.index)
    report_df["Sr. No."] = np.arange(1, len(working) + 1)
    report_df["Student Full Name"] = working["Student Full Name"]
    report_df["Standardized Entry Label"] = working["Standardized Entry Label"]
    report_df["Entry Label"] = working[resolved_missed["Entry Label"]]
    report_df["Date of missed scheduled meeting"] = working[resolved_missed["Date of missed scheduled meeting"]]

    reason_col = optional_missed["Reason for Missed Meeting"]
    if reason_col:
        report_df["Reason for Missed Meeting"] = working[reason_col]
    else:
        report_df["Reason for Missed Meeting"] = pd.Series(pd.NA, index=report_df.index)

    rescheduled_col = optional_missed["Date of rescheduled advising session (if applicable)"]
    if rescheduled_col:
        report_df["Date of rescheduled advising session (if applicable)"] = working[rescheduled_col]
    else:
        report_df["Date of rescheduled advising session (if applicable)"] = pd.Series(pd.NA, index=report_df.index)

    resolved_master = {
        label: _resolve_column(master_df, label, aliases, required=True)
        for label, aliases in MASTER_REQUIRED.items()
    }
    lookup_df = master_df[
        [
            resolved_master["Student Full Name"],
            resolved_master["PS Number"],
            resolved_master["Mentor"],
            resolved_master["ADEK Advisor"],
        ]
    ].copy()

    rename_map = {
        resolved_master["Student Full Name"]: "Student Full Name",
        resolved_master["PS Number"]: "PS Number",
        resolved_master["Mentor"]: "Mentor",
        resolved_master["ADEK Advisor"]: "ADEK Advisor",
    }
    lookup_df = lookup_df.rename(columns=rename_map)
    lookup_df["__lookup_key"] = _build_lookup_key(lookup_df["Student Full Name"])
    lookup_df = lookup_df.drop_duplicates(subset="__lookup_key", keep="first")

    report_df["__lookup_key"] = _build_lookup_key(report_df["Student Full Name"])
    report_df = report_df.merge(
        lookup_df[["__lookup_key", "PS Number", "Mentor", "ADEK Advisor"]],
        on="__lookup_key",
        how="left",
    )
    report_df = report_df.drop(columns="__lookup_key")

    report_df = report_df.reindex(columns=REPORT_COLUMNS)
    return report_df


def _select_sheet_with_columns(
    sheets: Dict[str, pd.DataFrame],
    required_map: Dict[str, Iterable[str]],
) -> Tuple[str, pd.DataFrame]:
    """Return the first sheet that contains all required columns (considering aliases)."""

    last_error: Exception | None = None
    for sheet_name, df in sheets.items():
        try:
            for label, aliases in required_map.items():
                _resolve_column(df, label, aliases, required=True)
        except KeyError as exc:
            last_error = exc
            continue
        return sheet_name, df

    if last_error:
        raise last_error
    raise KeyError("No sheet contains the required columns.")


def clean_workbook(
    missed_sheets: Dict[str, pd.DataFrame],
    master_sheets: Dict[str, pd.DataFrame],
    *,
    report_sheet_name: str | None = None,
) -> Tuple[Dict[str, pd.DataFrame], Dict[str, Dict[str, int]]]:
    """Create the missed session report and retain the original export sheet."""

    if not missed_sheets:
        raise ValueError("Upload for the missed session export is required.")
    if not master_sheets:
        raise ValueError("Upload for the master data is required.")

    missed_name, missed_df = _select_sheet_with_columns(missed_sheets, MISSED_REQUIRED)
    master_name, master_df = _select_sheet_with_columns(master_sheets, MASTER_REQUIRED)

    report_df = _build_missed_session_report(missed_df, master_df)

    sheet_label = report_sheet_name or "Missed Sessions"
    cleaned_sheets: Dict[str, pd.DataFrame] = {
        sheet_label: report_df,
        "Export": missed_df.copy(),
    }

    summary = SheetSummary(
        sheet_name=missed_name or "",
        initial_rows=len(missed_df),
        duplicates_removed=0,
        invalid_contacts_cleared=0,
        final_rows=len(report_df),
    )
    summaries: Dict[str, Dict[str, int]] = {
        "Missed Sessions": summary.as_dict,
        "TOTAL": summary.as_dict,
    }
    return cleaned_sheets, summaries


def export_cleaned_workbook(cleaned_sheets: Dict[str, pd.DataFrame], *, summary_title: str | None = None) -> bytes:
    """Write cleaned sheets to an Excel workbook with the prescribed formatting."""

    if not cleaned_sheets:
        raise ValueError("No cleaned data available to export.")

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        sheet_count = len(cleaned_sheets)
        used_names: set[str] = set()

        if summary_title:
            summary_name = _unique_sheet_name("Summary", used_names)
            worksheet = workbook.add_worksheet(summary_name)
            writer.sheets[summary_name] = worksheet
            _write_summary_sheet(workbook, worksheet, summary_title)

        for original_name, df in cleaned_sheets.items():
            sheet_name = (
                original_name
                if sheet_count > 1
                else "Cleaned Data"
            )
            final_name = _unique_sheet_name(sheet_name, used_names)
            if original_name == "Export":
                df.to_excel(writer, sheet_name=final_name, index=False)
            else:
                worksheet = workbook.add_worksheet(final_name)
                writer.sheets[final_name] = worksheet
                _write_with_formatting(workbook, worksheet, df)

    output.seek(0)
    return output.getvalue()


def _smart_title_case(value: str) -> str:
    """Convert strings like names into title case while handling separators."""

    text = (value or "").strip()
    if not text:
        return ""

    def _title_fragment(fragment: str) -> str:
        if not fragment:
            return fragment
        return fragment[0].upper() + fragment[1:].lower()

    tokens = re.split(r"\s+", text)
    rebuilt_tokens: List[str] = []
    for token in tokens:
        parts = re.split(r"([-'])", token)
        rebuilt = "".join(
            part if part in {"-", "'"} else _title_fragment(part)
            for part in parts
        )
        rebuilt_tokens.append(rebuilt)

    return " ".join(rebuilt_tokens)


def _write_with_formatting(workbook, worksheet, df: pd.DataFrame) -> None:
    """Write data to the worksheet with the required styling."""

    df_to_write = df.copy()
    if df_to_write.empty and len(df_to_write.columns) == 0:
        df_to_write = pd.DataFrame({" ": []})

    headers = list(df_to_write.columns)
    rows = df_to_write.replace({pd.NA: "", np.nan: ""}).values.tolist()

    start_row = 1  # zero-based index => Excel row 2
    start_col = 1  # zero-based index => Excel column B
    total_rows = len(rows)
    total_cols = len(headers)

    format_cache = {}

    def get_format(*, is_header: bool, top=False, bottom=False, left=False, right=False):
        key = (is_header, top, bottom, left, right)
        if key in format_cache:
            return format_cache[key]

        base = {
            "text_wrap": True,
            "border": 1,
            "font_color": HEADER_FONT_COLOR if is_header else BODY_FONT_COLOR,
            "align": "center" if is_header else "left",
            "valign": "vcenter",
            "font_name": "Gopher",
            "font_size": 10,
        }
        if is_header:
            base.update(
                {
                    "bg_color": NA_COLOR,
                    "bold": True,
                }
            )

        if top:
            base["top"] = 5
        if bottom:
            base["bottom"] = 5
        if left:
            base["left"] = 5
        if right:
            base["right"] = 5

        fmt = workbook.add_format(base)
        format_cache[key] = fmt
        return fmt

    # Write header row.
    for col_idx, header in enumerate(headers):
        is_left = col_idx == 0
        is_right = col_idx == total_cols - 1
        fmt = get_format(is_header=True, top=True, left=is_left, right=is_right, bottom=total_rows == 0)
        worksheet.write(start_row, start_col + col_idx, header, fmt)

    # Write data rows.
    for row_offset, row_values in enumerate(rows):
        excel_row = start_row + 1 + row_offset
        is_bottom = row_offset == total_rows - 1
        for col_idx, cell_value in enumerate(row_values):
            is_left = col_idx == 0
            is_right = col_idx == total_cols - 1
            fmt = get_format(
                is_header=False,
                top=False,
                bottom=is_bottom,
                left=is_left,
                right=is_right,
            )
            worksheet.write(excel_row, start_col + col_idx, cell_value, fmt)

    # Auto-fit columns based on content length.
    widths = _estimate_column_widths(headers, rows)
    for idx, width in enumerate(widths):
        worksheet.set_column(start_col + idx, start_col + idx, width)

    if total_cols > 0:
        filter_end_row = start_row if total_rows == 0 else start_row + total_rows
        filter_end_col = start_col + total_cols - 1
        worksheet.autofilter(start_row, start_col, filter_end_row, filter_end_col)


def _estimate_column_widths(headers: Iterable[str], rows: List[List]) -> List[int]:
    """Estimate Excel column widths from data content."""

    widths = [len(str(header)) for header in headers]
    for row in rows:
        for idx, value in enumerate(row):
            widths[idx] = max(widths[idx], len(str(value)))

    return [min(max(width + 2, 12), 60) for width in widths]


def _sanitize_sheet_name(name: str) -> str:
    """Ensure sheet names comply with Excel restrictions."""

    cleaned = re.sub(r"[\[\]\*:/\\\?]", "_", name).strip() or "Cleaned Data"
    return cleaned[:31]


def _unique_sheet_name(name: str, used_names: set[str]) -> str:
    """Generate a unique sheet name within Excel's length constraints."""

    base = _sanitize_sheet_name(name)
    candidate = base
    counter = 1
    while candidate in used_names:
        suffix = f"_{counter}"
        cutoff = 31 - len(suffix)
        candidate = f"{base[:cutoff]}{suffix}"
        counter += 1

    used_names.add(candidate)
    return candidate


def _write_summary_sheet(workbook, worksheet, title: str) -> None:
    """Render a simple summary header at B2:F2."""

    fmt = workbook.add_format(
        {
            "align": "center",
            "valign": "vcenter",
            "font_name": "Gopher",
            "font_size": 16,
            "bold": True,
            "font_color": HEADER_FONT_COLOR,
            "bg_color": NA_COLOR,
        }
    )
    worksheet.merge_range(1, 1, 1, 5, title, fmt)
    worksheet.set_row(1, 28)
    worksheet.set_column(1, 5, 22)
