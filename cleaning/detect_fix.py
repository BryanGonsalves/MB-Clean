"""Data loading, cleaning, and export helpers for the MB-Clean app."""

from __future__ import annotations

import io
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Tuple

import numpy as np
import pandas as pd
from email_validator import EmailNotValidError, validate_email
import phonenumbers
from phonenumbers.phonenumberutil import NumberParseException, PhoneNumberFormat

NA_COLOR = "#0A1D44"
HEADER_FONT_COLOR = "#FFFFFF"
BODY_FONT_COLOR = "#000000"


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


def clean_workbook(
    sheets: Dict[str, pd.DataFrame],
) -> Tuple[Dict[str, pd.DataFrame], Dict[str, Dict[str, int]]]:
    """Clean every sheet in the uploaded workbook."""

    cleaned_sheets: Dict[str, pd.DataFrame] = {}
    summaries: Dict[str, Dict[str, int]] = {}
    totals = SheetSummary(sheet_name="TOTAL", initial_rows=0, duplicates_removed=0, invalid_contacts_cleared=0, final_rows=0)

    for sheet_name, df in sheets.items():
        cleaned_df, summary = _clean_dataframe(
            df,
        )
        cleaned_sheets[sheet_name] = cleaned_df
        summaries[sheet_name] = summary.as_dict

        totals.initial_rows += summary.initial_rows
        totals.duplicates_removed += summary.duplicates_removed
        totals.invalid_contacts_cleared += summary.invalid_contacts_cleared
        totals.final_rows += summary.final_rows

    summaries["TOTAL"] = totals.as_dict
    return cleaned_sheets, summaries


def export_cleaned_workbook(cleaned_sheets: Dict[str, pd.DataFrame]) -> bytes:
    """Write cleaned sheets to an Excel workbook with the prescribed formatting."""

    if not cleaned_sheets:
        raise ValueError("No cleaned data available to export.")

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        sheet_count = len(cleaned_sheets)
        used_names: set[str] = set()

        for original_name, df in cleaned_sheets.items():
            sheet_name = (
                original_name
                if sheet_count > 1
                else "Cleaned Data"
            )
            final_name = _unique_sheet_name(sheet_name, used_names)
            worksheet = workbook.add_worksheet(final_name)
            writer.sheets[final_name] = worksheet
            _write_with_formatting(workbook, worksheet, df)

    output.seek(0)
    return output.getvalue()


def _clean_dataframe(
    df: pd.DataFrame,
) -> Tuple[pd.DataFrame, SheetSummary]:
    """Apply the full cleaning pipeline to a single DataFrame."""

    working_df = df.copy()
    initial_rows = len(working_df)

    # Early exit for empty sheets.
    if initial_rows == 0:
        return working_df, SheetSummary(
            sheet_name="",
            initial_rows=0,
            duplicates_removed=0,
            invalid_contacts_cleared=0,
            final_rows=0,
        )

    working_df = _trim_whitespace(working_df)
    working_df = _normalize_name_columns(working_df)
    working_df = _normalize_date_columns(working_df)

    pre_dedup_rows = len(working_df)
    working_df = working_df.drop_duplicates()
    duplicates_removed = pre_dedup_rows - len(working_df)

    working_df, invalid_contacts_cleared = _validate_contacts(working_df)

    working_df = working_df.reset_index(drop=True)

    summary = SheetSummary(
        sheet_name="",
        initial_rows=initial_rows,
        duplicates_removed=duplicates_removed,
        invalid_contacts_cleared=invalid_contacts_cleared,
        final_rows=len(working_df),
    )
    return working_df, summary


def _trim_whitespace(df: pd.DataFrame) -> pd.DataFrame:
    """Strip leading/trailing whitespace from all string columns."""

    trimmed = df.copy()
    object_cols = trimmed.select_dtypes(include=["object", "string"]).columns

    for col in object_cols:
        trimmed[col] = trimmed[col].astype("string").str.strip()

    return trimmed


def _normalize_name_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize columns that appear to contain name data."""

    normalized = df.copy()
    name_cols = [col for col in normalized.columns if "name" in str(col).lower()]

    if not name_cols:
        return normalized

    for col in name_cols:
        series = normalized[col].astype("string")
        normalized_values = series.fillna("").map(_smart_title_case)
        normalized[col] = normalized_values.replace("", pd.NA)

    return normalized


def _normalize_date_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Convert date-like values to ISO 8601 strings."""

    normalized = df.copy()
    candidate_cols = [
        col
        for col in normalized.columns
        if "date" in str(col).lower()
    ]

    excel_origin = "1899-12-30"

    for col in candidate_cols:
        series = normalized[col]
        parsed = pd.to_datetime(series, errors="coerce", utc=False)

        numeric_mask = parsed.isna() & series.notna()
        if numeric_mask.any():
            numeric_values = pd.to_numeric(series[numeric_mask], errors="coerce")
            numeric_values = numeric_values.where(numeric_values.between(59, 600000))
            excel_dates = pd.to_datetime(
                numeric_values,
                unit="D",
                origin=excel_origin,
                errors="coerce",
            )
            excel_valid = excel_dates.notna()
            if excel_valid.any():
                valid_idx = excel_dates.index[excel_valid]
                parsed.loc[valid_idx] = excel_dates.loc[valid_idx]

        mask_valid = parsed.notna()
        formatted = series.astype("string")

        formatted.loc[mask_valid] = parsed.loc[mask_valid].dt.strftime("%Y-%m-%d")
        formatted = formatted.str.strip()

        normalized[col] = formatted.replace("", pd.NA)

    return normalized


def _validate_contacts(df: pd.DataFrame) -> Tuple[pd.DataFrame, int]:
    """Validate email and phone columns, clearing invalid values without dropping rows."""

    validated = df.copy()
    email_cols = [col for col in validated.columns if "email" in str(col).lower()]
    phone_cols = [
        col
        for col in validated.columns
        if any(token in str(col).lower() for token in ("phone", "mobile", "contact", "tel"))
    ]

    invalid_entries = 0

    for col in email_cols:
        series = validated[col]
        for idx, value in series.items():
            if _is_blank(value):
                validated.loc[idx, col] = pd.NA
                continue
            try:
                normalized_email = validate_email(str(value), check_deliverability=False).email
            except EmailNotValidError:
                validated.loc[idx, col] = pd.NA
                invalid_entries += 1
            else:
                validated.loc[idx, col] = normalized_email

    for col in phone_cols:
        series = validated[col]
        for idx, value in series.items():
            if _is_blank(value):
                validated.loc[idx, col] = pd.NA
                continue

            sanitized = re.sub(r"[^\d+]", "", str(value))

            parsed_number = None
            try:
                parsed_number = phonenumbers.parse(sanitized, None)
            except NumberParseException:
                try:
                    parsed_number = phonenumbers.parse(sanitized, "US")
                except NumberParseException:
                    parsed_number = None

            if not parsed_number or not phonenumbers.is_valid_number(parsed_number):
                validated.loc[idx, col] = pd.NA
                invalid_entries += 1
                continue

            formatted_number = phonenumbers.format_number(parsed_number, PhoneNumberFormat.E164)
            validated.loc[idx, col] = formatted_number

    return validated, invalid_entries


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


def _is_blank(value) -> bool:
    """Check if a value should be considered blank."""

    if pd.isna(value):
        return True
    return str(value).strip() == ""


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
