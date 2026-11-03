"""
Braille token normalizer and positional tokenizer.

Special tokens:
  CAP    -> single character '\x00' (CAP marker)
  COMMA  -> ','
  PERIOD -> '.'
  EXCLA  -> '!'

Spaces are preserved exactly as individual positional elements.
"""

CAP_CHAR = "\x00"  # internal sentinel for CAP token

TOKEN_MAP = {
    "CAP": CAP_CHAR,
    "COMMA": ",",
    "PERIOD": ".",
    "EXCLA": "!",
}

# Longest token first so greedy matching works correctly
_TOKENS_SORTED = sorted(TOKEN_MAP.keys(), key=len, reverse=True)


def normalize(text: str) -> str:
    """
    Convert special tokens in *text* to their single-character equivalents.

    Rules
    -----
    - CAP  counts as exactly ONE character (sentinel \\x00).
    - COMMA, PERIOD, EXCLA each become one punctuation character.
    - Every original space is preserved in its exact position.
    - No other normalization is applied.
    """
    if not isinstance(text, str):
        text = "" if text is None else str(text)

    result: list[str] = []
    i = 0
    n = len(text)

    while i < n:
        if text[i] == " ":
            result.append(" ")
            i += 1
            continue

        matched = False
        for token in _TOKENS_SORTED:
            end = i + len(token)
            if text[i:end] == token:
                result.append(TOKEN_MAP[token])
                i = end
                matched = True
                break

        if not matched:
            result.append(text[i])
            i += 1

    return "".join(result)
