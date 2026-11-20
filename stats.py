"""
Paired t-test for Braille transcription accuracy.

Compares two sets of evaluation results (raw/baseline vs preprocessed)
using Character Accuracy (%) per sample_id as the paired measurement.

Formula:  t = d̄ / (Sd / √n)

Where:
  d̄  = mean of (preprocessed_acc - raw_acc) per sample
  Sd  = standard deviation of the differences
  n   = number of paired samples
"""

import math
import numpy as np
import pandas as pd
from dataclasses import dataclass, field


@dataclass
class TTestResult:
    """Result of a paired t-test for one tab or the overall dataset."""
    tab_name: str
    n: int                              # number of paired samples
    d_bar: float                        # mean difference d̄
    sd: float                           # std dev of differences Sd
    t_stat: float                       # t = d̄ / (Sd / √n)
    pairs: pd.DataFrame = field(default_factory=pd.DataFrame)
    # columns: sample_id, acc_raw, acc_pre, d
    error: str = ""                     # non-empty if pairing failed


def _sample_char_acc(results_df: pd.DataFrame) -> dict[str, float]:
    """
    Compute Character Accuracy (%) per sample_id from a results DataFrame.

    Accuracy = total_correct_predicted_characters / total_characters * 100
    aggregated across all lines for that sample.
    """
    acc = {}
    if "sample_id" not in results_df.columns:
        return acc
    for sid in results_df["sample_id"].unique():
        sdf = results_df[results_df["sample_id"] == sid]
        total   = pd.to_numeric(sdf["total characters"],                   errors="coerce").sum()
        correct = pd.to_numeric(sdf["total correct predicted characters"], errors="coerce").sum()
        acc[str(sid)] = (correct / total * 100) if total > 0 else 0.0
    return acc


def paired_ttest(
    tab_name: str,
    raw_df: pd.DataFrame,
    pre_df: pd.DataFrame,
) -> TTestResult:
    """
    Run a paired t-test comparing raw vs preprocessed results for one tab.

    Pairs are matched by sample_id.  Only sample_ids present in BOTH
    DataFrames are included (inner join).

    Returns a TTestResult.  If fewer than 2 paired samples exist the
    t-statistic cannot be computed; the error field will explain why.
    """
    acc_raw = _sample_char_acc(raw_df)
    acc_pre = _sample_char_acc(pre_df)

    # Inner join on sample_id
    common = sorted(set(acc_raw) & set(acc_pre))

    if len(common) == 0:
        return TTestResult(tab_name=tab_name, n=0, d_bar=0.0, sd=0.0,
                           t_stat=0.0,
                           error="No matching sample_ids between raw and preprocessed.")

    if len(common) < 2:
        sid = common[0]
        d = acc_pre[sid] - acc_raw[sid]
        pairs = pd.DataFrame([{
            "sample_id": sid,
            "Accuracy Raw (%)": round(acc_raw[sid], 4),
            "Accuracy Pre (%)": round(acc_pre[sid], 4),
            "d (Pre − Raw)":    round(d, 4),
        }])
        return TTestResult(tab_name=tab_name, n=1, d_bar=d, sd=0.0,
                           t_stat=0.0, pairs=pairs,
                           error="Only 1 paired sample — t-statistic requires n ≥ 2.")

    rows = []
    for sid in common:
        d = acc_pre[sid] - acc_raw[sid]
        rows.append({
            "sample_id":           sid,
            "Accuracy Raw (%)":    round(acc_raw[sid], 4),
            "Accuracy Pre (%)":    round(acc_pre[sid], 4),
            "d (Pre − Raw)":       round(d, 4),
        })
    pairs = pd.DataFrame(rows)

    diffs  = pairs["d (Pre − Raw)"].values.astype(float)
    n      = len(diffs)
    d_bar  = float(np.mean(diffs))
    sd     = float(np.std(diffs, ddof=1))          # sample std dev
    t_stat = d_bar / (sd / math.sqrt(n)) if sd > 0 else 0.0

    return TTestResult(
        tab_name=tab_name,
        n=n,
        d_bar=round(d_bar, 6),
        sd=round(sd, 6),
        t_stat=round(t_stat, 6),
        pairs=pairs,
    )


def run_all_ttests(
    results: list[tuple[str, "pd.DataFrame", "pd.DataFrame"]],
    uploaded_files: list,
) -> dict[str, TTestResult] | None:
    """
    Detect which uploaded files are raw vs preprocessed (by filename
    containing "raw", case-insensitive), pair their tabs by sheet name,
    and return a dict of {tab_name: TTestResult} plus an "Overall" key.

    Returns None if exactly 2 files weren't uploaded or one of each type
    can't be identified.
    """
    # Detect raw vs preprocessed files
    raw_stems, pre_stems = [], []
    for f in uploaded_files:
        name = f.name.lower()
        if "raw" in name:
            raw_stems.append(f.name)
        else:
            pre_stems.append(f.name)

    if not raw_stems or not pre_stems:
        return None     # caller decides whether to show a warning

    # Build lookup: tab_name -> results_df, keyed by which file it came from
    raw_file_names  = set(raw_stems)
    pre_file_names  = set(pre_stems)

    # We need to know which result came from which file.  results only carries
    # (tab_name, results_df, cm) so we re-derive file membership from the
    # upload order stored in session_state.
    # Instead we pass the structured dict built in app.py — see _build_file_tab_map.
    return None   # placeholder; actual logic in _build_file_tab_map below


def build_ttest_results(
    file_tab_results: dict[str, dict[str, pd.DataFrame]],
    # {filename: {tab_name: results_df}}
) -> dict[str, TTestResult]:
    """
    Given a mapping of filename → {tab_name → results_df}, identify
    raw vs preprocessed, pair tabs by name, run t-tests, and return
    {tab_name: TTestResult, "Overall": TTestResult}.
    """
    raw_key = None
    pre_key = None
    for fname in file_tab_results:
        if "raw" in fname.lower():
            raw_key = fname
        else:
            pre_key = fname

    if raw_key is None or pre_key is None:
        return {}

    raw_tabs = file_tab_results[raw_key]   # {tab_name: df}
    pre_tabs = file_tab_results[pre_key]   # {tab_name: df}

    ttest_results: dict[str, TTestResult] = {}

    # Per-tab
    common_tabs = sorted(set(raw_tabs) & set(pre_tabs))
    for tab in common_tabs:
        ttest_results[tab] = paired_ttest(tab, raw_tabs[tab], pre_tabs[tab])

    # Overall — concatenate all paired samples across tabs
    if common_tabs:
        all_pairs = pd.concat(
            [ttest_results[t].pairs for t in common_tabs if not ttest_results[t].pairs.empty],
            ignore_index=True,
        )
        if not all_pairs.empty:
            diffs  = all_pairs["d (Pre − Raw)"].values.astype(float)
            n      = len(diffs)
            d_bar  = float(np.mean(diffs))
            sd     = float(np.std(diffs, ddof=1)) if n >= 2 else 0.0
            t_stat = d_bar / (sd / math.sqrt(n)) if sd > 0 and n >= 2 else 0.0
            ttest_results["Overall"] = TTestResult(
                tab_name="Overall",
                n=n,
                d_bar=round(d_bar, 6),
                sd=round(sd, 6),
                t_stat=round(t_stat, 6),
                pairs=all_pairs,
                error="" if n >= 2 else "Need n ≥ 2 paired samples for t-statistic.",
            )

    return ttest_results
