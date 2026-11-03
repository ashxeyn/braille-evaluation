"""
Braille Transcription Evaluation System
Streamlit front-end
"""

import os
import streamlit as st
import pandas as pd

from excel_io import read_input_excel_all_sheets, write_multi_results_excel
from batch_processor import process_dataframe

st.set_page_config(
    page_title="Evaluation for Braillelink",
    page_icon="⠿",
    layout="wide",
)

st.title("Evaluation for Braillelink")
st.markdown(
    "Upload one or more Excel files. Each file becomes its own tab in the "
    "downloaded results workbook."
)

# ---------------------------------------------------------------------------
# 1. File upload (multiple)
# ---------------------------------------------------------------------------
uploaded_files = st.file_uploader(
    "Upload Excel file(s) (.xlsx or .xls)",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
    help="Required columns: sample_id, line_number, expected_output, predicted_output",
)

if not uploaded_files:
    st.info("Please upload at least one Excel file to begin.")
    st.stop()

# ---------------------------------------------------------------------------
# 2. Validate and preview each file
# ---------------------------------------------------------------------------
valid_files: list[tuple[str, pd.DataFrame]] = []  # (tab_label, input_df)

for f in uploaded_files:
    file_stem = os.path.splitext(f.name)[0]
    try:
        sheets = read_input_excel_all_sheets(f)
    except ValueError as err:
        st.error(f"{f.name}: {err}")
        continue

    if not sheets:
        st.warning(f"{f.name}: No sheets found.")
        continue

    for sheet_name, df, error in sheets:
        # Tab label: "filename — sheet" if multiple sheets, just filename if one
        tab_label = f"{file_stem} — {sheet_name}" if len(sheets) > 1 else file_stem

        if error:
            st.warning(f"{f.name} / {sheet_name}: {error} — skipped.")
            continue

        valid_files.append((tab_label, df))
        with st.expander(f"Preview: {tab_label} — {len(df):,} row(s)", expanded=False):
            st.dataframe(df.head(20), use_container_width=True)
            if len(df) > 20:
                st.caption(f"Showing first 20 of {len(df):,} rows.")

if not valid_files:
    st.stop()

st.success(f"{len(valid_files)} file(s) loaded successfully.")

# ---------------------------------------------------------------------------
# 3. Evaluate
# ---------------------------------------------------------------------------
if st.button("Evaluate Transcriptions", type="primary"):
    results: list[tuple[str, pd.DataFrame]] = []

    for stem, input_df in valid_files:
        with st.spinner(f"Evaluating {stem}…"):
            results_df = process_dataframe(input_df)
        results.append((stem, results_df))

    st.success("Evaluation complete.")

    # ---- Per-file summary metrics ----------------------------------------
    for stem, results_df in results:
        # Show anomaly warnings at the top
        if "anomaly" in results_df.columns:
            anomaly_rows = results_df[results_df["anomaly"].str.strip().astype(bool)]
            if not anomaly_rows.empty:
                st.warning(f"⚠️ **{stem}** — {len(anomaly_rows)} anomalous line(s) detected (highlighted red in Excel):")
                for _, ar in anomaly_rows.iterrows():
                    st.error(f"Sample {ar.get('sample_id','?')} | Line {ar.get('line_number','?')}: {ar['anomaly']}")

        st.subheader(f"Summary — {stem}")
        numeric_cols = [
            "total characters", "total words",
            "total correct predicted characters",
            "total correct predicted words",
            "total wrong characters", "total wrong words",
            "total deleted characters", "total additional characters",
        ]
        totals = {
            col: pd.to_numeric(results_df[col], errors="coerce").sum()
            for col in numeric_cols if col in results_df.columns
        }

        total_chars      = totals.get("total characters", 0)
        total_words      = totals.get("total words", 0)
        total_correct    = totals.get("total correct predicted characters", 0)
        total_correct_w  = totals.get("total correct predicted words", 0)
        total_wrong      = totals.get("total wrong characters", 0)
        total_deleted    = totals.get("total deleted characters", 0)
        total_additional = totals.get("total additional characters", 0)

        char_acc = (total_correct / total_chars * 100) if total_chars > 0 else 0.0
        word_acc = (total_correct_w / total_words * 100) if total_words > 0 else 0.0
        cer      = ((total_wrong + total_deleted + total_additional) / total_chars * 100) if total_chars > 0 else 0.0

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Characters",   int(total_chars))
        c2.metric("Total Words",         int(total_words))
        c3.metric("Char Accuracy",       f"{char_acc:.2f}%")
        c4.metric("Word Accuracy",       f"{word_acc:.2f}%")

        c5, c6, c7, c8 = st.columns(4)
        c5.metric("Correct Characters",  int(total_correct))
        c6.metric("Correct Words",       int(total_correct_w))
        c7.metric("Wrong Characters",    int(total_wrong))
        c8.metric("Deleted Characters",  int(total_deleted))

        c9, c10, c11 = st.columns(4)[:3]
        c9.metric("Additional Characters", int(total_additional))
        c10.metric("CER",                  f"{cer:.2f}%")
        wer = ((totals.get("total wrong words", 0)) / total_words * 100) if total_words > 0 else 0.0
        c11.metric("WER",                  f"{wer:.2f}%")

        with st.expander(f"Preview results — {stem}", expanded=False):
            st.dataframe(results_df.head(20), use_container_width=True)
            if len(results_df) > 20:
                st.caption(f"Showing first 20 of {len(results_df):,} rows.")

        # Per-sample breakdown
        if "sample_id" in results_df.columns:
            with st.expander(f"Per-sample breakdown — {stem}", expanded=False):
                for sid in results_df["sample_id"].unique():
                    sdf = results_df[results_df["sample_id"] == sid]
                    t_chars      = pd.to_numeric(sdf["total characters"], errors="coerce").sum()
                    t_words      = pd.to_numeric(sdf["total words"], errors="coerce").sum()
                    t_correct    = pd.to_numeric(sdf["total correct predicted characters"], errors="coerce").sum()
                    t_corr_w     = pd.to_numeric(sdf["total correct predicted words"], errors="coerce").sum()
                    t_wrong      = pd.to_numeric(sdf.get("total wrong characters", pd.Series([0])), errors="coerce").sum()
                    t_deleted    = pd.to_numeric(sdf.get("total deleted characters", pd.Series([0])), errors="coerce").sum()
                    t_additional = pd.to_numeric(sdf.get("total additional characters", pd.Series([0])), errors="coerce").sum()
                    c_acc = (t_correct / t_chars * 100) if t_chars > 0 else 0.0
                    w_acc = (t_corr_w  / t_words  * 100) if t_words  > 0 else 0.0
                    cer   = ((t_wrong + t_deleted + t_additional) / t_chars * 100) if t_chars > 0 else 0.0

                    st.markdown(f"**Sample ID: {sid}**")
                    sc1, sc2, sc3, sc4, sc5 = st.columns(5)
                    sc1.metric("Characters",   int(t_chars))
                    sc2.metric("Words",        int(t_words))
                    sc3.metric("Char Acc",     f"{c_acc:.2f}%")
                    sc4.metric("Word Acc",     f"{w_acc:.2f}%")
                    sc5.metric("CER",          f"{cer:.2f}%")
                    st.divider()

        st.divider()

    # ---- Download --------------------------------------------------------
    st.subheader("Download Results")
    excel_bytes = write_multi_results_excel(results)
    st.download_button(
        label="⬇ Download Results Excel",
        data=excel_bytes,
        file_name="braille_evaluation_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
