"""Shared input validators for live Word tools.

Centralizes safety checks that protect against control-character bytes
which can corrupt Word's Find/Replace engine — most notably the cell
separator (\\x07), which historically deleted entire documents when
passed as find_text.
"""

# Control bytes (U+0000–U+001F) that are unsafe to pass to Word.Range.Find.
# Permitted exceptions: tab (\t), line feed (\n), carriage return (\r) —
# these have well-defined meanings in Word's text model.
# \x0B (vertical tab / Word's manual line break) is NOT exempt: passing it
# raw to Find can match non-breaking line endings unpredictably; callers
# who want a manual line break should use use_wildcards=True with ^l.
_FORBIDDEN_FIND_CHARS = {chr(c) for c in range(0x20)} - {"\t", "\n", "\r"}


def reject_control_chars(label: str, text: str) -> None:
    """Raise ValueError if `text` contains forbidden control characters.

    Args:
        label: Parameter name for the error message (e.g. "find_text").
        text:  The string to validate.

    Raises:
        ValueError: with a message naming the offending bytes (hex) and
            pointing the caller at safer alternatives.
    """
    if not text:
        return
    bad = sorted({hex(ord(c)) for c in text if c in _FORBIDDEN_FIND_CHARS})
    if bad:
        raise ValueError(
            f"{label} contains forbidden control characters {bad}. "
            "Cell separator (\\x07), bell, and other control bytes corrupt "
            "Word Find/Replace and have caused full-document data loss. "
            "If you are trying to remove orphan cell separators left by a "
            "prior delete_table, call word_live_modify_table with "
            "operation='delete_table' and scrub_orphans=True (default), "
            "or use word_live_diagnose_layout to locate them. "
            "For paragraph marks / page breaks, use use_wildcards=True with "
            "^p / ^m."
        )
