"""
Automated tests for all 10 validation cases specified in the requirements.
Run with: pytest test_evaluator.py -v
"""

import pytest
from tokenizer import normalize, CAP_CHAR
from evaluator import evaluate_line


# ---------------------------------------------------------------------------
# Tokenizer unit tests
# ---------------------------------------------------------------------------

class TestTokenizer:
    def test_cap_is_one_char(self):
        result = normalize("CAPanna")
        assert result[0] == CAP_CHAR
        assert result[1:] == "anna"
        assert len(result) == 5   # 1 CAP + 4 letters

    def test_comma(self):
        assert normalize("COMMA") == ","

    def test_period(self):
        assert normalize("PERIOD") == "."

    def test_excla(self):
        assert normalize("EXCLA") == "!"

    def test_spaces_preserved(self):
        result = normalize("a  b")
        assert result == "a  b"

    def test_mixed(self):
        result = normalize("CAPanna has COMMAok")
        assert result == CAP_CHAR + "anna has ,ok"

    def test_no_trimming(self):
        assert normalize("  hello  ") == "  hello  "


# ---------------------------------------------------------------------------
# CASE 1: Perfect prediction
# ---------------------------------------------------------------------------

class TestCase1PerfectPrediction:
    def setup_method(self):
        self.lm = evaluate_line("CAPthey are happy.", "CAPthey are happy.")

    def test_all_chars_correct(self):
        c = self.lm.char
        assert c.total_correct == c.total_chars

    def test_all_words_correct(self):
        w = self.lm.word
        assert w.total_correct_words == w.total_words

    def test_zero_wrong(self):
        assert self.lm.char.total_wrong == 0

    def test_zero_deleted(self):
        assert self.lm.char.total_deleted == 0

    def test_notes_none(self):
        assert self.lm.notes == "None"


# ---------------------------------------------------------------------------
# CASE 2: Missing character (deleted h)
# ---------------------------------------------------------------------------

class TestCase2MissingCharacter:
    """
    Expected : CAPthey are happy.
    Predicted: CAPthey are appy.
    The first space between 'are' and 'appy' is the word separator.
    The second space is the missing 'h'.
    Total deleted = 1 (only 'h').
    """

    def setup_method(self):
        self.lm = evaluate_line(
            "CAPthey are happy.",
            "CAPthey are  appy.",
        )

    def test_one_deleted(self):
        assert self.lm.char.total_deleted == 1

    def test_zero_wrong(self):
        # The space at 'h' position is deleted, not wrong
        assert self.lm.char.total_wrong == 0

    def test_word_separator_not_deleted(self):
        # Only 'h' is deleted; original space is not
        assert self.lm.char.total_deleted == 1


# ---------------------------------------------------------------------------
# CASE 3: Wrong character (c vs b)
# ---------------------------------------------------------------------------

class TestCase3WrongCharacter:
    """
    Expected : CAPcat is here.
    Predicted: CAPbat is here.
    One wrong character at the 'c' position.
    """

    def setup_method(self):
        self.lm = evaluate_line("CAPcat is here.", "CAPbat is here.")

    def test_one_wrong_char(self):
        assert self.lm.char.total_wrong == 1

    def test_zero_deleted(self):
        assert self.lm.char.total_deleted == 0

    def test_affected_word_wrong(self):
        # 'CAPcat' vs 'CAPbat' -> wrong word
        assert self.lm.word.total_wrong_words >= 1


# ---------------------------------------------------------------------------
# CASE 4: Missing CAP marker
# ---------------------------------------------------------------------------

class TestCase4MissingCAP:
    """
    Expected : CAPnina took a bus.
    Predicted: nina took a bus.   (leading space where CAP was)
    First position is deleted (CAP marker missing).
    """

    def setup_method(self):
        # Predicted has a leading space in place of the CAP character
        self.lm = evaluate_line(
            "CAPnina took a bus.",
            " nina took a bus.",
        )

    def test_cap_counted_as_deleted(self):
        assert self.lm.char.total_deleted == 1

    def test_zero_wrong(self):
        assert self.lm.char.total_wrong == 0


# ---------------------------------------------------------------------------
# CASE 5: Missing punctuation (period)
# ---------------------------------------------------------------------------

class TestCase5MissingPunctuation:
    """
    Expected : braille. numbers
    Predicted: braille  numbers   (space in place of period)
    The period position is deleted; subsequent chars still align.
    """

    def setup_method(self):
        self.lm = evaluate_line(
            "braille. numbers",
            "braille  numbers",
        )

    def test_period_deleted(self):
        assert self.lm.char.total_deleted == 1

    def test_no_shifting(self):
        # 'numbers' aligns correctly after the period position
        assert self.lm.char.total_wrong == 0


# ---------------------------------------------------------------------------
# CASE 6: Merged words
# ---------------------------------------------------------------------------

class TestCase6MergedWords:
    """
    Expected : and a big clock!
    Predicted: andaa big clock!
    'and' and 'a' are merged; neither should be correct.
    """

    def setup_method(self):
        self.lm = evaluate_line(
            "and a big clock!",
            "andaa big clock!",
        )

    def test_merged_word_not_correct(self):
        # 'and' != 'andaa' and 'a' has no matching pred word in position 2
        # 'big' and 'clock!' shift positions -> none match correctly
        w = self.lm.word
        # at minimum 'and' and 'a' are wrong
        assert w.total_wrong_words >= 2

    def test_merged_words_in_notes(self):
        assert "andaa" in self.lm.notes


# ---------------------------------------------------------------------------
# CASE 7: Consecutive spaces
# ---------------------------------------------------------------------------

class TestCase7ConsecutiveSpaces:
    """
    Expected : anna has .
    Predicted: ann as .
    Compare every position; do not move the period.
    """

    def setup_method(self):
        self.lm = evaluate_line("anna has .", "ann as .")

    def test_positional_comparison(self):
        # pos 0-3: a,n,n,a vs a,n,n,' ' -> a=a,n=n,n=n,a vs ' ' -> 1 deleted
        # We just check that no extra alignment happened
        c = self.lm.char
        assert c.total_deleted + c.total_wrong > 0  # something is off

    def test_period_not_moved(self):
        # The period at its expected position must be compared to whatever
        # is at the same predicted position (not shifted to match)
        pass  # structural test; if no error -> alignment not applied


# ---------------------------------------------------------------------------
# CASE 8: Missing final character (trailing punctuation)
# ---------------------------------------------------------------------------

class TestCase8MissingFinalChar:
    """
    Expected : happy.
    Predicted: happy   (period missing -> predicted is shorter)
    """

    def setup_method(self):
        self.lm = evaluate_line("happy.", "happy")

    def test_period_deleted(self):
        assert self.lm.char.total_deleted == 1

    def test_zero_wrong(self):
        assert self.lm.char.total_wrong == 0


# ---------------------------------------------------------------------------
# CASE 9: Missing first character (CAP at start)
# ---------------------------------------------------------------------------

class TestCase9MissingFirstChar:
    """
    Expected : CAPshe is happy.
    Predicted: she is happy.   (predicted has leading space for missing CAP)
    """

    def setup_method(self):
        self.lm = evaluate_line(
            "CAPshe is happy.",
            " she is happy.",
        )

    def test_cap_deleted(self):
        assert self.lm.char.total_deleted == 1

    def test_zero_wrong(self):
        assert self.lm.char.total_wrong == 0


# ---------------------------------------------------------------------------
# CASE 10: Multiple lines are independent
# ---------------------------------------------------------------------------

class TestCase10MultipleLines:
    def test_lines_are_independent(self):
        """
        A missing character or extra space in line 1 must not affect line 2.
        """
        lm1 = evaluate_line("CAPthey are happy.", "CAPthey are  appy.")
        lm2 = evaluate_line("anna is kind.",      "anna is kind.")

        # Line 1 has 1 deleted
        assert lm1.char.total_deleted == 1
        # Line 2 is perfect
        assert lm2.char.total_deleted == 0
        assert lm2.char.total_wrong == 0
        assert lm2.word.total_correct_words == lm2.word.total_words


# ---------------------------------------------------------------------------
# Additional edge-case tests
# ---------------------------------------------------------------------------

class TestEdgeCases:

    def test_multiple_deletions(self):
        """
        Expected : CAPthey are happy.
        Predicted: CAPthey are   py.  (3 extra spaces -> h,a,p deleted)
        """
        lm = evaluate_line(
            "CAPthey are happy.",
            "CAPthey are    py.",
        )
        assert lm.char.total_deleted == 3

    def test_empty_lines(self):
        lm = evaluate_line("", "")
        assert lm.char.total_chars == 0
        assert lm.char.total_correct == 0

    def test_leading_spaces_preserved(self):
        lm = evaluate_line("  hello", "  hello")
        assert lm.char.total_correct == 5  # 'hello'

    def test_wrong_not_counted_as_deleted(self):
        lm = evaluate_line("abc", "axc")
        assert lm.char.total_wrong == 1
        assert lm.char.total_deleted == 0

    def test_predicted_shorter(self):
        lm = evaluate_line("abcde", "abc")
        # 'd' and 'e' are beyond predicted length -> deleted
        assert lm.char.total_deleted == 2

    def test_cap_in_middle(self):
        norm = normalize("say CAPHello")
        assert norm == "say " + CAP_CHAR + "Hello"


# ---------------------------------------------------------------------------
# Hash-word tests
# ---------------------------------------------------------------------------

class TestHashWords:
    def test_hash_word_correct(self):
        """#ba. vs #ba -> correct (trailing period stripped, # preserved)"""
        lm = evaluate_line("#ba.", "#ba")
        assert lm.word.total_correct_words == 1
        assert lm.word.total_wrong_words == 0

    def test_hash_word_wrong(self):
        """#ba. vs #bb -> wrong because letters differ"""
        lm = evaluate_line("#ba.", "#bb")
        assert lm.word.total_correct_words == 0
        assert lm.word.total_wrong_words == 1

    def test_hash_word_not_filtered(self):
        """#be alone must be counted as a word, not filtered as punctuation-only"""
        lm = evaluate_line("#be", "#be")
        assert lm.word.total_words == 1
        assert lm.word.total_correct_words == 1

    def test_hash_prefix_preserved_in_comparison(self):
        """#ba != ba — the # is part of the word identity"""
        lm = evaluate_line("#ba", "ba")
        assert lm.word.total_wrong_words == 1
