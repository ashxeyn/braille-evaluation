"""
Portrait image export — US Letter (8.5 × 11 in, 150 dpi → 1275 × 1650 px).

Each image has two sections:
  1. Overall Summary  — aggregated metrics table
  2. Detailed Results — row-by-row sentence + metric columns

Per-sample breakdowns and the confusion matrix are excluded.
All drawing is done in figure-fraction coordinates so axes never add
padding or whitespace beyond the actual table rows.
"""

import io
import re

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.transforms as mtransforms

from excel_io import _compute_summary, SUMMARY_METRICS

# ---------------------------------------------------------------------------
# Page constants  (US Letter portrait)
# ---------------------------------------------------------------------------
_PAGE_W = 8.5
_PAGE_H = 11.0
_DPI    = 150          # → 1275 × 1650 px

# Margins in figure-fraction
_ML = 0.03
_MR = 0.97
_MT = 0.955   # top of first section (below title)
_MB = 0.015   # minimum bottom margin

# Row heights in figure-fraction (height of one table row on the page)
_SUMMARY_ROW_H = 0.026
_DETAIL_ROW_H  = 0.018

# Gap between sections and label height
_GAP     = 0.012
_LABEL_H = 0.016

# Font sizes (pt)
_TITLE_FS   = 11
_SECTION_FS = 7.5
_SUMMARY_FS = 7.0
_DETAIL_FS  = 5.0

# Max chars before truncation in detail cells
_CELL_LIMIT = 20

# Colours
_HDR_BG = "#1F4E79"
_HDR_FG = "white"
_ALT_BG = "#D6E4F0"
_TTL_FG = "#1F4E79"

# Short column labels for the detail table
_SHORT = {
    "sample_id":                                          "Sample",
    "line_number":                                        "Line",
    "expected_output":                                    "Expected",
    "predicted_output":                                   "Predicted",
    "total characters":                                   "Tot Char",
    "total words":                                        "Tot Words",
    "total correct predicted characters":                 "Corr Char",
    "total correct predicted characters including spaces":"CorrCh+Sp",
    "total correct predicted words":                      "Corr Words",
    "total predicted characters":                         "Pred Char",
    "total predicted words":                              "Pred Words",
    "total wrong characters":                             "Wrong Char",
    "total wrong words":                                  "Wrong Words",
    "total deleted characters":                           "Del Char",
    "total additional characters":                        "Add Char",
    "notes":                                              "Notes",
    "anomaly":                                            "Anomaly",
}

_RATE_METRICS = frozenset({
    "Character Accuracy (%)", "Word Accuracy (%)",
    "Character Error Rate (%)", "Word Error Rate (%)",
})


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _fmt(metric, value):
    if metric in _RATE_METRICS:
        try:
            return f"{float(value):.2f}%"
        except (TypeError, ValueError):
            return str(value)
    try:
        return str(int(value))
    except (TypeError, ValueError):
        return str(value)


def _trunc(text, limit):
    s = str(text)
    return s if len(s) <= limit else s[:limit - 1] + "…"


def _draw_table_fig(fig, headers, rows, font_size, row_h,
                    x_left, x_right, top_y, cell_limit=None):
    """
    Draw a table directly in figure-fraction space.

    Parameters
    ----------
    fig      : matplotlib Figure
    headers  : list of column header strings
    rows     : list of row lists
    font_size: pt
    row_h    : height of each row in figure-fraction units
    x_left   : left edge (figure-fraction)
    x_right  : right edge (figure-fraction)
    top_y    : top edge of the first (header) row (figure-fraction)
    cell_limit: truncate cell text to this many chars (None = no truncation)

    Returns the y coordinate of the bottom edge of the last drawn row.
    """
    n_cols = len(headers)
    col_w  = (x_right - x_left) / n_cols
    trans  = fig.transFigure

    all_rows = [headers] + rows
    for r_idx, row_data in enumerate(all_rows):
        y_top = top_y - r_idx * row_h
        y_bot = y_top - row_h
        is_header = r_idx == 0
        is_alt    = (not is_header) and (r_idx % 2 == 0)

        for c_idx, cell_val in enumerate(row_data):
            x0 = x_left + c_idx * col_w
            x1 = x0 + col_w

            # Background rectangle
            bg = _HDR_BG if is_header else (_ALT_BG if is_alt else "white")
            rect = mpatches.FancyBboxPatch(
                (x0, y_bot), col_w, row_h,
                boxstyle="square,pad=0",
                facecolor=bg, edgecolor="#BBBBBB", linewidth=0.3,
                transform=trans, clip_on=False,
            )
            fig.add_artist(rect)

            # Cell text
            raw = str(cell_val)
            if cell_limit:
                raw = _trunc(raw, cell_limit)
            color  = _HDR_FG if is_header else "black"
            weight = "bold"  if is_header else "normal"

            fig.text(
                x0 + col_w / 2,
                (y_top + y_bot) / 2,
                raw,
                ha="center", va="center",
                fontsize=font_size,
                color=color, fontweight=weight,
                transform=trans, clip_on=False,
            )

    # Return y of bottom of last row
    return top_y - len(all_rows) * row_h


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def tab_to_portrait_image(results_df, tab_name, dpi=_DPI) -> bytes:
    """
    Render one tab into a fixed US Letter portrait PNG (1275 × 1650 px).
    Sections: Overall Summary + Detailed Results only.
    Returns PNG bytes.
    """
    # ---- Prepare data -------------------------------------------------------
    totals       = _compute_summary(results_df)
    summary_rows = [[m, _fmt(m, totals.get(m, 0))] for m in SUMMARY_METRICS]
    n_sum        = len(summary_rows)   # 15 rows

    detail_headers = [_SHORT.get(str(c), str(c)) for c in results_df.columns]
    detail_rows    = results_df.astype(str).values.tolist()
    n_det          = len(detail_rows)

    # ---- Compute actual heights ---------------------------------------------
    sum_rows_total = n_sum + 1          # +1 header
    det_rows_total = n_det + 1

    avail          = _MT - _MB
    labels_and_gap = 2 * _LABEL_H + _GAP
    table_space    = avail - labels_and_gap

    # Cap both row heights to sensible maximums so neither table looks bloated.
    # Any remaining blank space sits at the bottom — that's fine for a printed page.
    _MAX_SUM_ROW_H = 0.034
    _MAX_DET_ROW_H = 0.034

    # Summary: use up to 48% of table_space, but cap per-row height
    sum_row_h  = min(_MAX_SUM_ROW_H, (table_space * 0.48) / sum_rows_total)
    sum_height = sum_row_h * sum_rows_total

    # Detail: cap per-row height; use whatever rows × height gives
    det_row_h  = min(_MAX_DET_ROW_H, (table_space - sum_height) / det_rows_total)
    det_height = det_row_h * det_rows_total

    # ---- Create figure (fixed US Letter) ------------------------------------
    fig = plt.figure(figsize=(_PAGE_W, _PAGE_H), dpi=dpi)
    fig.patch.set_facecolor("white")
    # Remove default axes
    fig.clf()

    # Title
    fig.text(
        0.5, 0.980,
        f"{_trunc(tab_name, 90)} — Evaluation Results",
        ha="center", va="top",
        fontsize=_TITLE_FS, fontweight="bold", color="black",
    )

    # ---- Overall Summary section --------------------------------------------
    cur_y = _MT

    fig.text(
        _ML, cur_y,
        "Overall Summary",
        ha="left", va="top",
        fontsize=_SECTION_FS, fontweight="bold", color=_TTL_FG,
    )
    cur_y -= _LABEL_H

    cur_y = _draw_table_fig(
        fig, ["Metric", "Value"], summary_rows,
        _SUMMARY_FS, sum_row_h,
        _ML, _MR, cur_y,
    )

    # ---- Detailed Results section -------------------------------------------
    cur_y -= _GAP

    fig.text(
        _ML, cur_y,
        "Detailed Results",
        ha="left", va="top",
        fontsize=_SECTION_FS, fontweight="bold", color=_TTL_FG,
    )
    cur_y -= _LABEL_H

    _draw_table_fig(
        fig, detail_headers, detail_rows,
        _DETAIL_FS, det_row_h,
        _ML, _MR, cur_y,
        cell_limit=_CELL_LIMIT,
    )

    # ---- Render -------------------------------------------------------------
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, facecolor="white",
                bbox_inches=None)
    plt.close(fig)
    buf.seek(0)
    return buf.read()


def sheet_to_image(df, sheet_name, dpi=_DPI) -> bytes:
    """
    Render any arbitrary DataFrame (one Excel sheet) into a fixed US Letter
    portrait PNG using the same visual style as tab_to_portrait_image.

    The image contains a single table of whatever data is in the sheet —
    no evaluation metrics, no summaries, no confusion matrix.

    Returns PNG bytes.
    """
    n_rows = len(df)
    n_rows_total = n_rows + 1   # +1 for header

    headers    = [str(c) for c in df.columns]
    table_rows = df.astype(str).values.tolist()

    avail      = _MT - _MB
    label_h    = _LABEL_H
    # Fixed compact row height — same feel as the summary table in evaluation images.
    # Blank space below is intentional (printed page behaviour).
    row_h      = 0.022

    fig = plt.figure(figsize=(_PAGE_W, _PAGE_H), dpi=dpi)
    fig.patch.set_facecolor("white")
    fig.clf()

    # Title
    fig.text(
        0.5, 0.980,
        f"{_trunc(sheet_name, 90)}",
        ha="center", va="top",
        fontsize=_TITLE_FS, fontweight="bold", color="black",
    )

    # Section label
    cur_y = _MT
    fig.text(
        _ML, cur_y,
        sheet_name,
        ha="left", va="top",
        fontsize=_SECTION_FS, fontweight="bold", color=_TTL_FG,
    )
    cur_y -= label_h

    _draw_table_fig(
        fig, headers, table_rows,
        _DETAIL_FS, row_h,
        _ML, _MR, cur_y,
        cell_limit=_CELL_LIMIT,
    )

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, facecolor="white", bbox_inches=None)
    plt.close(fig)
    buf.seek(0)
    return buf.read()


def safe_filename(name, limit=80) -> str:
    stem = re.sub(r'[\\/*?:"<>|]', "_", str(name)).strip().rstrip(".")
    return (stem or "tab")[:limit]
