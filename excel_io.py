"""
Excel import and export utilities.

Import  : read_input_excel  -> pandas DataFrame
Export  : write_results_excel -> bytes (in-memory .xlsx)
"""

import io
import re
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Characters that openpyxl / Excel will reject
_ILLEGAL_CHARS_RE = re.compile(r"[\x00-\x08\x0b-\x0c\x0e-\x1f]")


def _safe(value):
    """Strip illegal XML characters from string values before writing to Excel."""
    if isinstance(value, str):
        return _ILLEGAL_CHARS_RE.sub("", value)
    return value

REQUIRED_COLUMNS = ["sample_id", "line_number", "expected_output", "predicted_output"]

RESULT_COLUMNS = [
    "total characters",
    "total words",
    "total correct predicted characters",
    "total correct predicted characters including spaces",
    "total correct predicted words",
    "total predicted characters",
    "total predicted words",
    "total wrong characters",
    "total wrong words",
    "total deleted characters",
    "notes",
]

SUMMARY_METRICS = [
    "Total characters",
    "Total words",
    "Total correct predicted characters",
    "Total correct predicted characters including spaces",
    "Total correct predicted words",
    "Total predicted characters",
    "Total predicted words",
    "Total wrong characters",
    "Total wrong words",
    "Total deleted characters",
    "Total additional characters",
    "Character Accuracy (%)",
    "Character Error Rate (%)",
    "Word Accuracy (%)",
    "Word Error Rate (%)",
]


# ---------------------------------------------------------------------------
# Import
# ---------------------------------------------------------------------------

def read_input_excel(file_obj) -> pd.DataFrame:
    """
    Read the first sheet of an .xlsx or .xls file and return a DataFrame.
    Raises ValueError if required columns are missing.
    """
    try:
        df = pd.read_excel(file_obj, dtype=str, sheet_name=0)
    except Exception as exc:
        raise ValueError(f"Could not read Excel file: {exc}") from exc

    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            f"Missing required column(s): {', '.join(missing)}. "
            f"Found columns: {', '.join(df.columns)}"
        )

    for col in ["expected_output", "predicted_output"]:
        df[col] = df[col].fillna("").astype(str)

    return df


def read_input_excel_all_sheets(file_obj) -> list[tuple[str, pd.DataFrame, str | None]]:
    """
    Read ALL sheets from an .xlsx or .xls file.

    Each data sheet may optionally contain two extra columns anywhere:
        id                   -> maps to sample_id
        additional_characters -> total additional characters for that sample

    These two columns are extracted as a lookup table, then dropped from the
    main DataFrame before processing. The lookup is merged back by sample_id.

    Returns a list of (sheet_name, DataFrame_or_None, error_or_None).
    """
    try:
        all_sheets = pd.read_excel(file_obj, dtype=str, sheet_name=None)
    except Exception as exc:
        raise ValueError(f"Could not read Excel file: {exc}") from exc

    results = []
    for sheet_name, df in all_sheets.items():
        df.columns = [str(c).strip() for c in df.columns]

        missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
        if missing:
            results.append((sheet_name, None, f"Missing columns: {', '.join(missing)}"))
            continue

        for col in ["expected_output", "predicted_output"]:
            df[col] = df[col].fillna("").astype(str)
            df[col] = df[col].replace("nan", "")

        # Warn if predicted_output is entirely empty
        if df["predicted_output"].str.strip().eq("").all():
            results.append((sheet_name, None, "predicted_output column is empty — sheet skipped."))
            continue

        # Extract additional_characters lookup if both 'id' and
        # 'additional_characters' columns exist on this sheet
        additional_df = None
        if "id" in df.columns and "additional_characters" in df.columns:
            lookup = df[["id", "additional_characters"]].dropna(subset=["id"])
            lookup = lookup.copy()
            lookup["additional_characters"] = pd.to_numeric(
                lookup["additional_characters"], errors="coerce"
            ).fillna(0).astype(int)
            lookup = lookup.rename(columns={"id": "sample_id"})
            # Keep only rows where additional_characters > 0 or explicitly set
            lookup = lookup[lookup["additional_characters"].notna()]
            additional_df = lookup[["sample_id", "additional_characters"]].drop_duplicates("sample_id")

            # Drop these helper columns from the main data
            df = df.drop(columns=["id", "additional_characters"])

        # Merge lookup into main df
        if additional_df is not None:
            df = df.merge(additional_df, on="sample_id", how="left")
            df["additional_characters"] = df["additional_characters"].fillna(0).astype(int)

        results.append((sheet_name, df, None))

    return results


# ---------------------------------------------------------------------------
# Export helpers
# ---------------------------------------------------------------------------

_HEADER_FILL  = PatternFill("solid", fgColor="1F4E79")
_HEADER_FONT  = Font(bold=True, color="FFFFFF", size=11)
_ALT_FILL     = PatternFill("solid", fgColor="D6E4F0")
_ANOMALY_FILL = PatternFill("solid", fgColor="FF4444")
_ANOMALY_FONT = Font(bold=True, color="FFFFFF")
_BORDER_SIDE  = Side(style="thin", color="AAAAAA")
_CELL_BORDER  = Border(
    left=_BORDER_SIDE, right=_BORDER_SIDE,
    top=_BORDER_SIDE, bottom=_BORDER_SIDE,
)


def _style_header_row(ws, row_idx: int, n_cols: int) -> None:
    for col in range(1, n_cols + 1):
        cell = ws.cell(row=row_idx, column=col)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = _CELL_BORDER


def _style_data_row(ws, row_idx: int, n_cols: int, alternate: bool) -> None:
    fill = _ALT_FILL if alternate else PatternFill()
    for col in range(1, n_cols + 1):
        cell = ws.cell(row=row_idx, column=col)
        if alternate:
            cell.fill = fill
        cell.border = _CELL_BORDER
        cell.alignment = Alignment(vertical="center", wrap_text=False)


def _set_column_widths(ws, headers: list[str]) -> None:
    for i, header in enumerate(headers, start=1):
        col_letter = get_column_letter(i)
        # Base width on header length; cap at 40
        width = min(max(len(header) + 4, 12), 40)
        ws.column_dimensions[col_letter].width = width


# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

def _compute_summary(results_df: pd.DataFrame) -> dict:
    """Compute aggregated totals and accuracy from a results DataFrame."""
    numeric_map = {
        "Total characters":                                    "total characters",
        "Total words":                                         "total words",
        "Total correct predicted characters":                  "total correct predicted characters",
        "Total correct predicted characters including spaces": "total correct predicted characters including spaces",
        "Total correct predicted words":                       "total correct predicted words",
        "Total predicted characters":                          "total predicted characters",
        "Total predicted words":                               "total predicted words",
        "Total wrong characters":                              "total wrong characters",
        "Total wrong words":                                   "total wrong words",
        "Total deleted characters":                            "total deleted characters",
        "Total additional characters":                         "total additional characters",
    }
    totals: dict[str, float] = {}
    for label, col in numeric_map.items():
        totals[label] = pd.to_numeric(results_df[col], errors="coerce").sum() if col in results_df.columns else 0

    total_chars      = totals["Total characters"]
    total_words      = totals["Total words"]
    total_correct    = totals["Total correct predicted characters"]
    total_correct_w  = totals["Total correct predicted words"]
    total_wrong      = totals["Total wrong characters"]
    total_deleted    = totals["Total deleted characters"]
    total_additional = totals["Total additional characters"]

    totals["Character Accuracy (%)"] = round((total_correct / total_chars * 100) if total_chars > 0 else 0.0, 2)
    totals["Word Accuracy (%)"]      = round((total_correct_w / total_words * 100) if total_words > 0 else 0.0, 2)

    cer_errors = total_wrong + total_deleted + total_additional
    totals["Character Error Rate (%)"] = round((cer_errors / total_chars * 100) if total_chars > 0 else 0.0, 2)

    wer_errors = totals["Total wrong words"]
    totals["Word Error Rate (%)"] = round((wer_errors / total_words * 100) if total_words > 0 else 0.0, 2)

    return totals


def _write_tab(wb: Workbook, tab_name: str, results_df: pd.DataFrame) -> None:
    """
    Write one tab into wb containing:
      - Detailed results table
      - Blank separator
      - Per-sample summary (one block per unique sample_id)
      - Blank separator
      - Overall summary
    All in a single worksheet named tab_name.
    """
    # Sanitize tab name: Excel sheet names max 31 chars, no special chars
    safe_name = re.sub(r"[\\/*?:\[\]]", "_", tab_name)[:31]

    # Avoid duplicate sheet names
    existing = [s.title for s in wb.worksheets]
    if safe_name in existing:
        suffix = 2
        base = safe_name[:28]
        while f"{base}_{suffix}" in existing:
            suffix += 1
        safe_name = f"{base}_{suffix}"

    ws = wb.create_sheet(safe_name)

    # ---- Section 1: Detailed Results --------------------------------------
    headers = list(results_df.columns)
    n_cols  = len(headers)

    ws.append(headers)
    _style_header_row(ws, 1, n_cols)
    ws.freeze_panes = "A2"

    for r_idx, row in enumerate(results_df.itertuples(index=False), start=2):
        ws.append([_safe(v) for v in row])
        # Highlight anomaly rows in red
        anomaly_val = getattr(row, 'anomaly', '') or ''
        is_anomaly  = bool(str(anomaly_val).strip())
        if is_anomaly:
            for col in range(1, n_cols + 1):
                ws.cell(row=r_idx, column=col).fill   = _ANOMALY_FILL
                ws.cell(row=r_idx, column=col).font   = _ANOMALY_FONT
                ws.cell(row=r_idx, column=col).border = _CELL_BORDER
        else:
            _style_data_row(ws, r_idx, n_cols, alternate=(r_idx % 2 == 0))

    _set_column_widths(ws, headers)

    current_row = len(results_df) + 2  # after detail rows + blank

    # ---- Section 2: Per-sample summaries ----------------------------------
    ws.append([])  # blank separator
    current_row += 1

    # Section header
    sec_cell = ws.cell(row=current_row, column=1, value="Per-Sample Summary")
    sec_cell.font = Font(bold=True, size=13, color="1F4E79")
    current_row += 1

    sample_ids = results_df["sample_id"].unique() if "sample_id" in results_df.columns else []

    for sid in sample_ids:
        sample_df = results_df[results_df["sample_id"] == sid]
        totals    = _compute_summary(sample_df)

        # Sample label row
        label_cell = ws.cell(row=current_row, column=1, value=f"Sample ID: {sid}")
        label_cell.font = Font(bold=True, color="FFFFFF")
        label_cell.fill = PatternFill("solid", fgColor="2E75B6")
        label_cell.border = _CELL_BORDER
        ws.cell(row=current_row, column=2).fill = PatternFill("solid", fgColor="2E75B6")
        ws.cell(row=current_row, column=2).border = _CELL_BORDER
        current_row += 1

        # Header row for this sample
        ws.cell(row=current_row, column=1, value="Metric").font  = _HEADER_FONT
        ws.cell(row=current_row, column=1).fill   = _HEADER_FILL
        ws.cell(row=current_row, column=1).border = _CELL_BORDER
        ws.cell(row=current_row, column=1).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=current_row, column=2, value="Value").font  = _HEADER_FONT
        ws.cell(row=current_row, column=2).fill   = _HEADER_FILL
        ws.cell(row=current_row, column=2).border = _CELL_BORDER
        ws.cell(row=current_row, column=2).alignment = Alignment(horizontal="center", vertical="center")
        current_row += 1

        for i, metric in enumerate(SUMMARY_METRICS):
            value   = totals.get(metric, 0)
            display = f"{value:.2f}%" if metric in ("Character Accuracy (%)", "Word Accuracy (%)") else int(value)
            ws.cell(row=current_row, column=1, value=metric).border  = _CELL_BORDER
            ws.cell(row=current_row, column=2, value=display).border = _CELL_BORDER
            if i % 2 == 0:
                ws.cell(row=current_row, column=1).fill = _ALT_FILL
                ws.cell(row=current_row, column=2).fill = _ALT_FILL
            current_row += 1

        ws.append([])  # blank between samples
        current_row += 1

    # ---- Section 3: Overall Summary ---------------------------------------
    overall_label = ws.cell(row=current_row, column=1, value="Overall Summary")
    overall_label.font = Font(bold=True, size=13, color="1F4E79")
    current_row += 1

    ws.cell(row=current_row, column=1, value="Metric").font  = _HEADER_FONT
    ws.cell(row=current_row, column=1).fill   = _HEADER_FILL
    ws.cell(row=current_row, column=1).border = _CELL_BORDER
    ws.cell(row=current_row, column=1).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=current_row, column=2, value="Value").font  = _HEADER_FONT
    ws.cell(row=current_row, column=2).fill   = _HEADER_FILL
    ws.cell(row=current_row, column=2).border = _CELL_BORDER
    ws.cell(row=current_row, column=2).alignment = Alignment(horizontal="center", vertical="center")
    current_row += 1

    totals = _compute_summary(results_df)
    for i, metric in enumerate(SUMMARY_METRICS):
        value   = totals.get(metric, 0)
        display = f"{value:.2f}%" if metric in ("Character Accuracy (%)", "Word Accuracy (%)") else int(value)
        ws.cell(row=current_row, column=1, value=metric).border  = _CELL_BORDER
        ws.cell(row=current_row, column=2, value=display).border = _CELL_BORDER
        if i % 2 == 0:
            ws.cell(row=current_row, column=1).fill = _ALT_FILL
            ws.cell(row=current_row, column=2).fill = _ALT_FILL
        current_row += 1

    # Ensure summary columns are wide enough
    ws.column_dimensions["A"].width = max(ws.column_dimensions["A"].width, 50)
    ws.column_dimensions["B"].width = max(ws.column_dimensions["B"].width, 20)


def write_results_excel(results_df: pd.DataFrame) -> bytes:
    """
    Single-file export (backwards compat): one tab named 'Results'.
    """
    return write_multi_results_excel([("Results", results_df)])


def write_multi_results_excel(files: list[tuple[str, pd.DataFrame]]) -> bytes:
    """
    Build an in-memory .xlsx where each (name, df) pair becomes one tab.

    Parameters
    ----------
    files : list of (tab_name, results_df) tuples

    Returns raw bytes suitable for st.download_button.
    """
    wb = Workbook()
    # Remove the default empty sheet
    wb.remove(wb.active)

    for tab_name, results_df in files:
        _write_tab(wb, tab_name, results_df)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.read()
