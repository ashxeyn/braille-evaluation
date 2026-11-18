"""
Batch processor: applies the evaluator to every row in a DataFrame and
returns a new DataFrame with the evaluation columns appended.

additional_characters:
  Supplied via a separate 'additional_characters' sheet in the Excel file
  (columns: sample_id, additional_characters). The value is merged per
  sample_id before this processor runs, so every row for a given sample
  carries the same additional_characters value.

  To avoid double-counting in summaries, the total per sample is written
  once on the first row of each sample_id; subsequent rows for that sample
  get 0. Summaries can then simply sum the column.
"""

import numpy as np
import pandas as pd
from evaluator import evaluate_line, compute_positional_cm, aggregate_to_6x6
from tokenizer import normalize


def process_dataframe(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Apply the evaluator to every row and return a new DataFrame with the
    evaluation columns appended, PLUS the aggregated 6x6 confusion matrix.

    The existing per-row column additions, anomaly handling, and
    additional_characters deduplication are unchanged. A 31x31 count matrix
    is accumulated alongside (by calling ``compute_positional_cm`` on the
    normalized strings) and then folded into the 6x6 thesis rubric via
    ``aggregate_to_6x6``.

    Returns
    -------
    (results_df, cm_6x6_df)
        results_df  : original columns + 13 metric/anomaly/notes columns.
        cm_6x6_df   : 6x6 labeled confusion-matrix DataFrame.
    """
    seen_samples: set = set()
    records = []

    # Running sum of the 31x31 raw matrix across all rows.
    cm_31 = np.zeros((31, 31), dtype=np.int64)

    for _, row in df.iterrows():
        expected  = str(row.get("expected_output",  ""))
        predicted = str(row.get("predicted_output", ""))

        # --- existing evaluation (unchanged) -------------------------------
        lm = evaluate_line(expected, predicted)

        # --- passive confusion-matrix accumulation -------------------------
        # evaluate_line() normalizes internally; we normalize here too so the
        # CM helpers receive the same normalized form. normalize() is
        # idempotent, so this never alters the evaluation result above.
        exp_norm  = normalize(expected)
        pred_norm = normalize(predicted)
        cm_31 += compute_positional_cm(exp_norm, pred_norm)

        c  = lm.char
        w  = lm.word

        sid = row.get("sample_id", "")

        # Write additional_characters once per sample to avoid double-counting
        if sid not in seen_samples:
            raw_add = pd.to_numeric(row.get("additional_characters", 0), errors="coerce")
            additional = int(raw_add) if pd.notna(raw_add) else 0
            seen_samples.add(sid)
        else:
            additional = 0

        result_row = dict(row)
        # Drop the raw additional_characters column — we'll write our own clean version
        result_row.pop("additional_characters", None)

        result_row["total characters"]                                    = c.total_chars
        result_row["total words"]                                         = w.total_words
        result_row["total correct predicted characters"]                  = c.total_correct
        result_row["total correct predicted characters including spaces"]  = c.total_correct_incl_spaces
        result_row["total correct predicted words"]                       = w.total_correct_words
        result_row["total predicted characters"]                          = c.total_predicted_chars
        result_row["total predicted words"]                               = w.total_predicted_words
        result_row["total wrong characters"]                              = c.total_wrong
        result_row["total wrong words"]                                   = w.total_wrong_words
        result_row["total deleted characters"]                            = c.total_deleted
        result_row["total additional characters"]                         = additional
        result_row["notes"]                                               = lm.notes
        result_row["anomaly"]                                             = lm.anomaly

        records.append(result_row)

    cm_6x6 = aggregate_to_6x6(cm_31)
    return pd.DataFrame(records), cm_6x6
