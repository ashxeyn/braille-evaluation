"""
Braille Transcription Evaluation System
Streamlit front-end
"""

import os
import io
import zipfile
import streamlit as st
import pandas as pd

from excel_io import read_input_excel_all_sheets, write_multi_results_excel
from batch_processor import process_dataframe
from image_export import tab_to_portrait_image, sheet_to_image, safe_filename
from stats import build_ttest_results

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
# Excel → Image converter (independent — placed here so st.stop() calls
# in the evaluation section below never block it)
# ---------------------------------------------------------------------------
with st.expander("📄 Excel to Image Converter", expanded=False):
    st.markdown(
        "Upload any Excel file to export each sheet as a portrait image "
        "(US Letter, same format as the evaluation results images). "
        "This is independent from the evaluation below."
    )

    converter_files = st.file_uploader(
        "Upload Excel file(s) to convert (.xlsx or .xls)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="converter_uploader",
    )

    if converter_files:
        conv_sheets: list[tuple[str, pd.DataFrame]] = []

        for f in converter_files:
            try:
                all_sheets = pd.read_excel(f, dtype=str, sheet_name=None)
            except Exception as exc:
                st.error(f"{f.name}: could not read — {exc}")
                continue

            for sheet_name, df in all_sheets.items():
                df = df.fillna("")
                conv_sheets.append((sheet_name, df))
                with st.expander(
                    f"Preview: {sheet_name} — {len(df):,} row(s)", expanded=False
                ):
                    st.dataframe(df.head(20), use_container_width=True)

        if conv_sheets:
            if st.button("Convert to Images", key="conv_btn", type="primary"):
                conv_zip_buf = io.BytesIO()
                with zipfile.ZipFile(
                    conv_zip_buf, "w", compression=zipfile.ZIP_DEFLATED
                ) as zf:
                    for sheet_name, df in conv_sheets:
                        with st.spinner(f"Rendering {sheet_name}…"):
                            png = sheet_to_image(df, sheet_name)
                        zf.writestr(f"{safe_filename(sheet_name)}.png", png)
                conv_zip_buf.seek(0)

                st.download_button(
                    label="⬇ Download Images (ZIP)",
                    data=conv_zip_buf.read(),
                    file_name="excel_sheets_images.zip",
                    mime="application/zip",
                    type="primary",
                )

st.divider()

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
valid_file_map: list[tuple[str, str, pd.DataFrame]] = []  # (filename, tab_label, input_df)

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
        # Tab label: use the original sheet name from the Excel file
        tab_label = sheet_name

        if error:
            st.warning(f"{f.name} / {sheet_name}: {error} — skipped.")
            continue

        valid_files.append((tab_label, df))
        valid_file_map.append((f.name, tab_label, df))
        with st.expander(f"Preview: {tab_label} — {len(df):,} row(s)", expanded=False):
            st.dataframe(df.head(20), use_container_width=True)
            if len(df) > 20:
                st.caption(f"Showing first 20 of {len(df):,} rows.")

if not valid_files:
    st.stop()

st.success(f"{len(valid_files)} file(s) loaded successfully.")

# ---------------------------------------------------------------------------
# 3. Evaluate (results persist in session state so the Evaluate button is
#    hidden after the first click, leaving only the download buttons).
# ---------------------------------------------------------------------------
if "eval_results" not in st.session_state:
    st.session_state.eval_results = None

# Clear stale results whenever the uploaded file set changes.
_upload_signature = tuple((f.name, getattr(f, "size", None)) for f in uploaded_files)
if st.session_state.get("upload_signature") != _upload_signature:
    st.session_state.eval_results = None
st.session_state.upload_signature = _upload_signature

if st.session_state.eval_results is None:
    if st.button("Evaluate Transcriptions", type="primary"):
        results: list[tuple[str, pd.DataFrame, pd.DataFrame]] = []
        file_tab_map: dict[str, dict[str, pd.DataFrame]] = {}

        for fname, stem, input_df in valid_file_map:
            with st.spinner(f"Evaluating {stem}…"):
                results_df, cm_6x6 = process_dataframe(input_df)
            results.append((stem, results_df, cm_6x6))
            if fname not in file_tab_map:
                file_tab_map[fname] = {}
            file_tab_map[fname][stem] = results_df

        ttest_results = build_ttest_results(file_tab_map)

        st.session_state.eval_results  = results
        st.session_state.ttest_results = ttest_results
        st.rerun()
else:
    results = st.session_state.eval_results
    st.success("Evaluation complete.")

    # ---- Overall summary across ALL tabs ---------------------------------
    if len(results) > 1:
        st.subheader("Overall Summary (All Tabs)")
        all_rows = pd.concat([rdf for _, rdf, _ in results], ignore_index=True)
        o_cols = [
            "total characters", "total words",
            "total correct predicted characters",
            "total correct predicted words",
            "total wrong characters", "total wrong words",
            "total deleted characters", "total additional characters",
        ]
        o = {
            col: pd.to_numeric(all_rows[col], errors="coerce").sum()
            for col in o_cols if col in all_rows.columns
        }
        o_chars      = o.get("total characters", 0)
        o_words      = o.get("total words", 0)
        o_correct    = o.get("total correct predicted characters", 0)
        o_correct_w  = o.get("total correct predicted words", 0)
        o_wrong      = o.get("total wrong characters", 0)
        o_wrong_w    = o.get("total wrong words", 0)
        o_deleted    = o.get("total deleted characters", 0)
        o_additional = o.get("total additional characters", 0)

        o_char_acc = (o_correct / o_chars * 100) if o_chars > 0 else 0.0
        o_word_acc = (o_correct_w / o_words * 100) if o_words > 0 else 0.0
        o_cer      = ((o_wrong + o_deleted + o_additional) / o_chars * 100) if o_chars > 0 else 0.0
        o_wer      = (o_wrong_w / o_words * 100) if o_words > 0 else 0.0

        o1, o2, o3, o4 = st.columns(4)
        o1.metric("Total Characters",   int(o_chars))
        o2.metric("Total Words",         int(o_words))
        o3.metric("Char Accuracy",       f"{o_char_acc:.2f}%")
        o4.metric("Word Accuracy",       f"{o_word_acc:.2f}%")

        o5, o6, o7, o8 = st.columns(4)
        o5.metric("Correct Characters",  int(o_correct))
        o6.metric("Correct Words",       int(o_correct_w))
        o7.metric("Wrong Characters",    int(o_wrong))
        o8.metric("Deleted Characters",  int(o_deleted))

        o9, o10, o11 = st.columns(4)[:3]
        o9.metric("Additional Characters", int(o_additional))
        o10.metric("CER",                  f"{o_cer:.2f}%")
        o11.metric("WER",                  f"{o_wer:.2f}%")

        st.divider()

    # ---- Per-file summary metrics ----------------------------------------
    for stem, results_df, cm_6x6 in results:
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

    # ---- Post-Deployment 6x6 Confusion Matrix ----------------------------
    st.subheader("Post-Deployment 6×6 Confusion Matrix")
    st.markdown(
        "Aggregated across all evaluated files/tabs below. Rows are the "
        "**actual** class group, columns are the **predicted** class group."
    )

    # Build the overall 6x6 matrix by summing each per-file matrix.
    # Aggregation is a linear sum, so summing the 6x6 frames is equivalent
    # to aggregating the summed 31x31 matrix.
    overall_cm: pd.DataFrame | None = None
    for _, _, cm_6x6 in results:
        if overall_cm is None:
            overall_cm = cm_6x6.copy()
        else:
            overall_cm = overall_cm.add(cm_6x6, fill_value=0)
    if overall_cm is not None:
        overall_cm = overall_cm.astype(int)
        st.dataframe(overall_cm, use_container_width=True)

        # Optional: per-file breakdown in expanders
        with st.expander("Per-file confusion matrices", expanded=False):
            for stem, _, cm_6x6 in results:
                st.markdown(f"**{stem}**")
                st.dataframe(cm_6x6.astype(int), use_container_width=True)
                st.divider()

    # ---- Paired t-Test results -------------------------------------------
    ttest_results = st.session_state.get("ttest_results", {})
    if ttest_results:
        st.subheader("Paired t-Test: Raw vs Preprocessed")
        st.markdown(
            "Character Accuracy (%) per sample is used as the paired measurement.  \n"
            "**d = Accuracy(Preprocessed) − Accuracy(Raw)** per sample_id."
        )

        tab_keys = [k for k in ttest_results if k != "Overall"]
        overall  = ttest_results.get("Overall")

        for tab_key in tab_keys:
            res = ttest_results[tab_key]
            with st.expander(f"Tab: {tab_key} — n={res.n}", expanded=False):
                if res.error:
                    st.warning(res.error)
                if not res.pairs.empty:
                    st.dataframe(res.pairs, use_container_width=True)
                if res.n >= 2:
                    tc1, tc2, tc3, tc4 = st.columns(4)
                    tc1.metric("n", res.n)
                    tc2.metric("d̄", f"{res.d_bar:.6f}")
                    tc3.metric("Sd", f"{res.sd:.6f}")
                    tc4.metric("t", f"{res.t_stat:.6f}")

        if overall:
            st.markdown("**Overall (all tabs combined)**")
            if overall.error:
                st.warning(overall.error)
            if not overall.pairs.empty:
                with st.expander("All paired samples (Overall)", expanded=False):
                    st.dataframe(overall.pairs, use_container_width=True)
            if overall.n >= 2:
                oc1, oc2, oc3, oc4 = st.columns(4)
                oc1.metric("n", overall.n)
                oc2.metric("d̄", f"{overall.d_bar:.6f}")
                oc3.metric("Sd", f"{overall.sd:.6f}")
                oc4.metric("t", f"{overall.t_stat:.6f}")

        st.divider()
    elif len(results) >= 2:
        # Files uploaded but neither has "raw" in the name
        st.info(
            "ℹ No paired t-test: to enable it, name one file with **'raw'** in the filename "
            "(e.g. `results_raw.xlsx`) and the other without (e.g. `results_preprocessed.xlsx`)."
        )
        st.divider()

    # ---- Downloads (side by side) ----------------------------------------
    try:
        excel_bytes = write_multi_results_excel(results, ttest_results or None)
    except Exception as _exc:
        st.warning(f"Could not include t-test sheet in Excel: {_exc}")
        excel_bytes = write_multi_results_excel(results, None)

    # Build ZIP of per-tab portrait PNGs only when a single file is uploaded
    # (two files = paired t-test mode; images not needed there)
    single_file_mode = len(set(fname for fname, _, _ in valid_file_map)) == 1

    st.subheader("Download")
    if single_file_mode:
        dl1, dl2, dl3 = st.columns(3)
    else:
        dl1, dl2 = st.columns(2)
        dl3 = None
    with dl1:
        if overall_cm is not None:
            st.download_button(
                label="⬇ Download 6×6 Confusion Matrix (CSV)",
                data=overall_cm.to_csv().encode("utf-8"),
                file_name="post_deployment_confusion_matrix_6x6.csv",
                mime="text/csv",
                use_container_width=True,
            )
    with dl2:
        st.download_button(
            label="⬇ Download Results Excel",
            data=excel_bytes,
            file_name="braille_evaluation_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
    if single_file_mode and dl3 is not None:
        with dl3:
            if st.button("⬇ Generate & Download Images (ZIP)",
                         key="gen_images_btn", use_container_width=True):
                try:
                    zip_buf = io.BytesIO()
                    with st.spinner("Rendering images…"):
                        with zipfile.ZipFile(zip_buf, "w",
                                             compression=zipfile.ZIP_DEFLATED) as zf:
                            for stem, results_df, _ in results:
                                png = tab_to_portrait_image(results_df, stem)
                                zf.writestr(f"{safe_filename(stem)}.png", png)
                    zip_buf.seek(0)
                    st.download_button(
                        label="⬇ Download Images (ZIP)",
                        data=zip_buf.read(),
                        file_name="braille_evaluation_images.zip",
                        mime="application/zip",
                        use_container_width=True,
                        key="images_zip_dl",
                    )
                except Exception as _zip_exc:
                    st.warning(f"Could not build images ZIP: {_zip_exc}")
