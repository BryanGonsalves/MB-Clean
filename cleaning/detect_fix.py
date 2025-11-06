"""Data loading, cleaning, and export helpers for the MB-Clean app."""

from __future__ import annotations

import io
import re
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd
from dateutil import parser as date_parser
from email_validator import EmailNotValidError, validate_email
import phonenumbers
from phonenumbers.phonenumberutil import NumberParseException, PhoneNumberFormat

try:
    from .ai_normalize import AINormalizer
except Exception:  # pragma: no cover - optional dependency
    AINormalizer = None

NA_COLOR = "#0A1D44"
HEADER_FONT_COLOR = "#FFFFFF"
BODY_FONT_COLOR = "#000000"


@dataclass
class SheetSummary:
    """Holds per-sheet cleaning statistics."""

    sheet_name: str
    initial_rows: int
    duplicates_removed: int
    invalid_rows_removed: int
    final_rows: int

    @property
    def as_dict(self) -> Dict[str, int]:
        return {
            "initial_rows": self.initial_rows,
            "duplicates_removed": self.duplicates_removed,
            "invalid_rows_removed": self.invalid_rows_removed,
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
    *,
    enable_ai: bool = False,
    ai_normalizer: Optional[AINormalizer] = None,
) -> Tuple[Dict[str, pd.DataFrame], Dict[str, Dict[str, int]]]:
    """Clean every sheet in the uploaded workbook."""

    cleaned_sheets: Dict[str, pd.DataFrame] = {}
    summaries: Dict[str, Dict[str, int]] = {}
    totals = SheetSummary(sheet_name="TOTAL", initial_rows=0, duplicates_removed=0, invalid_rows_removed=0, final_rows=0)

    for sheet_name, df in sheets.items():
        cleaned_df, summary = _clean_dataframe(
            df,
            enable_ai=enable_ai and ai_normalizer is not None,
            ai_normalizer=ai_normalizer,
        )
        cleaned_sheets[sheet_name] = cleaned_df
        summaries[sheet_name] = summary.as_dict

        totals.initial_rows += summary.initial_rows
        totals.duplicates_removed += summary.duplicates_removed
        totals.invalid_rows_removed += summary.invalid_rows_removed
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
    *,
    enable_ai: bool = False,
    ai_normalizer: Optional[AINormalizer] = None,
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
            invalid_rows_removed=0,
            final_rows=0,
        )

    working_df = _trim_whitespace(working_df)
    working_df = _normalize_name_columns(working_df, enable_ai, ai_normalizer)
    working_df = _normalize_date_columns(working_df, enable_ai, ai_normalizer)

    pre_dedup_rows = len(working_df)
    working_df = working_df.drop_duplicates()
    duplicates_removed = pre_dedup_rows - len(working_df)

    working_df, invalid_rows_removed = _validate_contacts(working_df)

    working_df = working_df.reset_index(drop=True)

    summary = SheetSummary(
        sheet_name="",
        initial_rows=initial_rows,
        duplicates_removed=duplicates_removed,
        invalid_rows_removed=invalid_rows_removed,
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


def _normalize_name_columns(
    df: pd.DataFrame,
    enable_ai: bool,
    ai_normalizer: Optional[AINormalizer],
) -> pd.DataFrame:
    """Normalize columns that appear to contain name data."""

    normalized = df.copy()
    name_cols = [col for col in normalized.columns if "name" in str(col).lower()]

    if not name_cols:
        return normalized

    for col in name_cols:
        series = normalized[col].astype("string")
        normalized_values = series.fillna("").map(_smart_title_case)

        if enable_ai and ai_normalizer:
            ai_candidates = {
                idx: value
                for idx, value in series.items()
                if _needs_ai_name_cleanup(value)
            }
            if ai_candidates:
                suggestions = ai_normalizer.normalize_names(list(ai_candidates.values()))
                for (idx, _), suggestion in zip(ai_candidates.items(), suggestions):
                    normalized_values.loc[idx] = suggestion.strip()

        normalized[col] = normalized_values.replace("", pd.NA)

    return normalized


def _normalize_date_columns(
    df: pd.DataFrame,
    enable_ai: bool,
    ai_normalizer: Optional[AINormalizer],
) -> pd.DataFrame:
    """Convert date-like values to ISO 8601 strings."""

    normalized = df.copy()
    candidate_cols = [
        col
        for col in normalized.columns
        if _looks_like_date_column(normalized[col])
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

        if enable_ai and ai_normalizer:
            unresolved_mask = (~mask_valid) & series.notna() & series.astype("string").str.strip().ne("")
            if unresolved_mask.any():
                original_values = series.loc[unresolved_mask].astype("string").tolist()
                ai_suggestions = ai_normalizer.normalize_dates(original_values)
                reparsed = pd.to_datetime(ai_suggestions, errors="coerce", utc=False)
                for idx, suggestion, parsed_value in zip(series.loc[unresolved_mask].index, ai_suggestions, reparsed):
                    if pd.notna(parsed_value):
                        formatted.loc[idx] = parsed_value.strftime("%Y-%m-%d")
                    else:
                        formatted.loc[idx] = suggestion.strip()

        normalized[col] = formatted.replace("", pd.NA)

    return normalized


def _validate_contacts(df: pd.DataFrame) -> Tuple[pd.DataFrame, int]:
    """Validate email and phone columns, removing rows with invalid values."""

    validated = df.copy()
    email_cols = [col for col in validated.columns if "email" in str(col).lower()]
    phone_cols = [
        col
        for col in validated.columns
        if any(token in str(col).lower() for token in ("phone", "mobile", "contact", "tel"))
    ]

    rows_to_drop = set()
    col_updates: Dict[str, Dict[int, str]] = defaultdict(dict)

    for col in email_cols:
        series = validated[col]
        for idx, value in series.items():
            if _is_blank(value):
                rows_to_drop.add(idx)
                continue
            try:
                normalized_email = validate_email(str(value), check_deliverability=False).email
            except EmailNotValidError:
                rows_to_drop.add(idx)
            else:
                col_updates[col][idx] = normalized_email

    for col in phone_cols:
        series = validated[col]
        for idx, value in series.items():
            if _is_blank(value):
                rows_to_drop.add(idx)
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
                rows_to_drop.add(idx)
                continue

            formatted_number = phonenumbers.format_number(parsed_number, PhoneNumberFormat.E164)
            col_updates[col][idx] = formatted_number

    if rows_to_drop:
        validated = validated.drop(index=list(rows_to_drop))

    for col, updates in col_updates.items():
        if not updates:
            continue
        valid_indices = [idx for idx in updates.keys() if idx not in rows_to_drop]
        if valid_indices:
            validated.loc[valid_indices, col] = [updates[idx] for idx in valid_indices]

    return validated, len(rows_to_drop)


def _looks_like_date_column(series: pd.Series) -> bool:
    """Heuristic to determine if a column contains dates."""

    if pd.api.types.is_datetime64_any_dtype(series):
        return True

    non_na_series = series.dropna()

    if non_na_series.empty:
        return False

    numeric = pd.to_numeric(non_na_series, errors="coerce")
    numeric = numeric.dropna()
    if not numeric.empty:
        plausible = numeric.between(59, 600000)  # Excel day numbers roughly through year 2600
        if plausible.mean() >= 0.5:
            return True

    sample = non_na_series.astype("string").str.strip().head(25)
    potential_matches = 0
    for value in sample:
        if not value:
            continue
        if re.search(r"\d{1,4}[^\w]{1}\d{1,2}[^\w]{1}\d{1,4}", value):
            potential_matches += 1
            continue
        try:
            date_parser.parse(value)
        except (ValueError, OverflowError):
            continue
        else:
            potential_matches += 1

    return potential_matches >= max(1, len(sample) // 2)


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


def _needs_ai_name_cleanup(value) -> bool:
    """Determine whether a name value is a candidate for AI refinement."""

    if _is_blank(value):
        return False

    text = str(value)
    stripped = text.strip()
    if not stripped:
        return False

    return stripped.islower() or stripped.isupper()


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
