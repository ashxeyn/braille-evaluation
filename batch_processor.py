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

import pandas as pd
from evaluator import evaluate_line


def process_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    seen_samples: set = set()
    records = []

    for _, row in df.iterrows():
        expected  = str(row.get("expected_output",  ""))
        predicted = str(row.get("predicted_output", ""))

        lm = evaluate_line(expected, predicted)
        c  = lm.char
        w  = lm.word

        sid = row.get("sample_id", "")

        # Write additional_characters once per sample to avoid double-counting
        if sid not in seen_samples:
            additional = int(pd.to_numeric(row.get("additional_characters", 0), errors="coerce") or 0)
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

    return pd.DataFrame(records)
