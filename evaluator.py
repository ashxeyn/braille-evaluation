"""
Strict positional character and word evaluator for Braille transcriptions.

Key rules
---------
- Compare positions 1-by-1 without any alignment or shifting.
- Spaces in the expected string are original word separators; they occupy
  positions but are NOT counted in character totals.
- When an expected NON-SPACE position has a predicted SPACE  -> deleted char.
- When an expected NON-SPACE position has a different predicted char -> wrong.
- When positions match (both non-space, same char)             -> correct.
- Positions beyond the expected length are not evaluated positionally.
- For word comparison words are derived from the original (pre-normalised)
  text so that empty slots created by consecutive spaces are excluded.
"""

from dataclasses import dataclass, field
from tokenizer import normalize, CAP_CHAR


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class CharMetrics:
    total_chars: int = 0                        # expected non-space chars
    total_correct: int = 0                      # correct (excl. spaces)
    total_correct_incl_spaces: int = 0          # correct positions incl. spaces
    total_predicted_chars: int = 0              # predicted non-space chars
    total_wrong: int = 0
    total_deleted: int = 0


@dataclass
class WordMetrics:
    total_words: int = 0
    total_correct_words: int = 0
    total_predicted_words: int = 0
    total_wrong_words: int = 0
    incorrect_word_pairs: list = field(default_factory=list)  # [(exp, pred), ...]


@dataclass
class LineMetrics:
    char: CharMetrics = field(default_factory=CharMetrics)
    word: WordMetrics = field(default_factory=WordMetrics)
    notes: str = ""
    anomaly: str = ""   # non-empty when a length mismatch is detected


# ---------------------------------------------------------------------------
# Character-level evaluation
# ---------------------------------------------------------------------------

def _evaluate_chars(exp_norm: str, pred_norm: str) -> CharMetrics:
    m = CharMetrics()

    exp_len = len(exp_norm)
    pred_len = len(pred_norm)
    compare_len = min(exp_len, pred_len)

    # Count total expected non-space characters
    m.total_chars = sum(1 for c in exp_norm if c != " ")

    # Count total predicted non-space characters
    m.total_predicted_chars = sum(1 for c in pred_norm if c != " ")

    for i in range(compare_len):
        e = exp_norm[i]
        p = pred_norm[i]

        if e == " ":
            # Original word separator — occupies a position, no char count
            if p == " ":
                # Both spaces: correct positional match (included in incl_spaces)
                m.total_correct_incl_spaces += 1
            # predicted non-space at an expected space -> not a deletion, not wrong
            # (extra predicted char at a space position; not counted positionally)
            continue

        # Expected is a real character (non-space)
        if p == " ":
            # Missing detection -> deleted
            m.total_deleted += 1
        elif p == e:
            # Correct match
            m.total_correct += 1
            m.total_correct_incl_spaces += 1
        else:
            # Wrong character
            m.total_wrong += 1

    # Positions in expected beyond predicted length: all deleted
    for i in range(compare_len, exp_len):
        if exp_norm[i] != " ":
            m.total_deleted += 1

    return m


# ---------------------------------------------------------------------------
# Word-level evaluation
# ---------------------------------------------------------------------------

def _split_words_with_positions(text: str) -> list[tuple[str, int, int]]:
    """
    Split text into words, returning (word, start, end) tuples.
    start/end are character indices in the original string.
    Empty slots from consecutive spaces are discarded.
    Punctuation-only tokens (without #) are excluded.
    """
    result = []
    i = 0
    n = len(text)
    while i < n:
        if text[i] == " ":
            i += 1
            continue
        # find end of word
        j = i
        while j < n and text[j] != " ":
            j += 1
        word = text[i:j]
        if not _is_punctuation_only(word):
            result.append((word, i, j))
        i = j
    return result


def _split_words(text: str) -> list[str]:
    """Return just the word strings, no positions."""
    return [w for w, _, _ in _split_words_with_positions(text)]


def _is_punctuation_only(token: str) -> bool:
    """
    Return True if every character in the token is punctuation/sentinel,
    with no alphanumeric or # content.
    Tokens like #ba are NOT punctuation-only because # introduces a word.
    """
    if not token:
        return False
    if "#" in token:
        return False
    punct = set(".,!?;:'\"()[]{}\\/-_@$%^&*+=<>|~`" + CAP_CHAR)
    return all(c in punct for c in token)


# Characters stripped from edges when normalising a word for comparison.
# '#' is intentionally excluded so that #ba stays #ba after stripping.
_WORD_STRIP_CHARS = ".,!?;:'\"()[]{}\\/-_@$%^&*+=<>|~`" + CAP_CHAR


def _normalize_word_for_comparison(word: str, expected_has_hash: bool = False) -> str:
    """
    Strip leading/trailing punctuation and CAP markers, then lowercase.

    - CAP sentinel is always stripped (it's a marker, not part of the word).
    - '#' is stripped from the predicted word ONLY when the expected word
      does not start with '#'. If expected starts with '#', both sides must
      keep '#' for the comparison to be meaningful.
    """
    stripped = word.strip(_WORD_STRIP_CHARS)
    if not expected_has_hash:
        stripped = stripped.lstrip("#")
    return stripped.lower()


def _is_merge_char(c: str) -> bool:
    """
    Return True if a character past the expected word boundary indicates a
    real merge (extra letters/digits — but NOT # which starts a new word,
    and NOT punctuation which naturally attaches to words).
    """
    if c == " ":
        return False
    # # always starts a new Braille number word — never a merge continuation
    if c == "#":
        return False
    punct = set(".,!?;:'\"()[]{}\\/-_@$%^&*+=<>|~`" + CAP_CHAR)
    return c not in punct


def _evaluate_words(exp_norm: str, pred_norm: str) -> WordMetrics:
    """
    Word comparison strategy:

    When expected and predicted have the same length (no net length difference),
    use positional span extraction — this correctly handles deletions represented
    as spaces and detects real merges.

    When lengths differ (predicted physically shorter/longer without space
    placeholders), fall back to index-based word list comparison to avoid
    false merge detections from misaligned spans.
    """
    m = WordMetrics()

    exp_words  = _split_words_with_positions(exp_norm)
    pred_words = _split_words_with_positions(pred_norm)

    m.total_words           = len(exp_words)
    m.total_predicted_words = len(pred_words)

    correct         = 0
    wrong           = 0
    incorrect_pairs = []

    same_length = (len(exp_norm) == len(pred_norm))

    if same_length:
        # ---- Span-based comparison (handles space-placeholder deletions) ----
        pred_len       = len(pred_norm)
        consumed_up_to = -1

        for exp_word, start, end in exp_words:
            if start <= consumed_up_to:
                wrong += 1
                incorrect_pairs.append((exp_word, "<consumed by merge>"))
                continue

            pred_span    = pred_norm[start:min(end, pred_len)]
            pred_in_span = pred_span.replace(" ", "")

            char_after = pred_norm[end] if end < pred_len else " "
            has_merge  = _is_merge_char(char_after)

            if has_merge:
                j = end
                while j < pred_len and pred_norm[j] != " ":
                    j += 1
                consumed_up_to = j - 1
                pred_full = pred_in_span + pred_norm[end:j]
                wrong += 1
                incorrect_pairs.append((exp_word, pred_full if pred_full else "<missing>"))
            else:
                has_hash = exp_word.lstrip(_WORD_STRIP_CHARS).startswith("#")
                exp_cmp  = _normalize_word_for_comparison(exp_word, expected_has_hash=has_hash)
                pred_cmp = _normalize_word_for_comparison(pred_in_span, expected_has_hash=has_hash)
                if exp_cmp == pred_cmp:
                    correct += 1
                else:
                    wrong += 1
                    incorrect_pairs.append((exp_word, pred_in_span if pred_in_span else "<missing>"))

    else:
        # ---- Index-based comparison (predicted length differs) --------------
        compare_len = min(len(exp_words), len(pred_words))

        for i in range(compare_len):
            ew, _, _ = exp_words[i]
            pw, _, _ = pred_words[i]
            has_hash = ew.lstrip(_WORD_STRIP_CHARS).startswith("#")
            if _normalize_word_for_comparison(ew, has_hash) == _normalize_word_for_comparison(pw, has_hash):
                correct += 1
            else:
                wrong += 1
                incorrect_pairs.append((ew, pw))

        for i in range(compare_len, len(exp_words)):
            ew, _, _ = exp_words[i]
            wrong += 1
            incorrect_pairs.append((ew, "<missing>"))

    m.total_correct_words  = correct
    m.total_wrong_words    = wrong
    m.incorrect_word_pairs = incorrect_pairs

    return m


# ---------------------------------------------------------------------------
# Notes builder
# ---------------------------------------------------------------------------

def _build_notes(word_metrics: WordMetrics) -> str:
    if not word_metrics.incorrect_word_pairs:
        return "None"

    def _display(w: str) -> str:
        return w.replace(CAP_CHAR, "CAP")

    parts = []
    for exp_w, pred_w in word_metrics.incorrect_word_pairs:
        parts.append(f"[{_display(exp_w)} -> {_display(pred_w)}]")

    return "Misspelled/Invalid words: " + " | ".join(parts)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def evaluate_line(expected_raw: str, predicted_raw: str) -> LineMetrics:
    """
    Evaluate a single line pair.

    Parameters
    ----------
    expected_raw  : original expected string (may contain CAP/COMMA/etc.)
    predicted_raw : original predicted string (may contain CAP/COMMA/etc.)

    Returns
    -------
    LineMetrics with char, word, and notes fields populated.
    """
    exp_norm = normalize(expected_raw)
    pred_norm = normalize(predicted_raw)

    lm = LineMetrics()
    lm.char = _evaluate_chars(exp_norm, pred_norm)
    lm.word = _evaluate_words(exp_norm, pred_norm)
    lm.notes = _build_notes(lm.word)

    # Anomaly detection: length mismatch causes positional cascade.
    # Two non-anomaly cases to ignore:
    # 1. Trailing spaces in predicted = missing detections at end of line (valid)
    # 2. Trailing spaces in expected  = Excel cell artifact (not real content)
    exp_len  = len(exp_norm)
    pred_len = len(pred_norm)

    if exp_len != pred_len:
        exp_trimmed  = len(exp_norm.rstrip(" "))
        pred_trimmed = len(pred_norm.rstrip(" "))
        # Use the smaller of actual vs trimmed for each side
        exp_cmp  = min(exp_len, exp_trimmed)
        pred_cmp = min(pred_len, pred_trimmed)
        if exp_cmp != pred_cmp:
            diff      = pred_cmp - exp_cmp
            direction = "longer" if diff > 0 else "shorter"
            lm.anomaly = (
                f"LENGTH MISMATCH: predicted is {abs(diff)} char(s) {direction} than expected "
                f"(expected={exp_cmp}, predicted={pred_cmp}). "
                f"Positional comparison may be affected from the point of difference."
            )

    return lm
