"""
shared.py

General-purpose helper functions used throughout the project.
Includes text normalization, fallback selection, user interaction, and data cleaning utilities."""

import re
import pandas as pd

from copy import deepcopy

def normalize_text(text):
    """
    Normalizes a given text by applying character substitutions and standardizations.

    Parameters:
        text (str): The input string.

    Returns:
        str: The normalized string.
    """
    substitutions = {
        'æ': 'ae', 'œ': 'oe',
        'é': 'e', 'è': 'e', 'ë': 'e', 'á': 'a', 'à': 'a',
        'û': 'u', 'î': 'i', 'â': 'a', 'ô': 'o', 'ê': 'e',
        'ü': 'u', 'ö': 'o', 'ä': 'a',
        'ß': 'ss',
        'iu': 'ie', 'üe': 'ue'
    }

    if not text:
        return ""

    text = text.lower()
    for old, new in substitutions.items():
        text = text.replace(old, new)

    text = re.sub(r'\bv\b', 'f', text)
    text = re.sub(r'\s+', ' ', text)

    return text

def get_first_valid_text(*fields):
    """
    Returns the first non-empty string from a list of inputs, skipping over None and NaN.

    Parameters:
        *fields (any): One or more values to evaluate.

    Returns:
        str: The first valid non-empty string, or an empty string if none found.
    """
    for f in fields:
        if isinstance(f, str) and f.strip():
            return f
    return ""

def clean_cell_value(value):
    """
    Returns a normalized string representation of a cell, or an empty string if missing.

    Parameters:
        value (any): The cell content.

    Returns:
        str: A normalized, lowercased string or empty string.
    """
    if pd.isna(value) or value is None:
        return ""
    return normalize_text(str(value).strip())

def sanitize_cell_value(value):
    """
    Cleans a cell value from invisible characters and ensures it is not an artifact like 'NaN'.

    Parameters:
        value (any): The cell content.

    Returns:
        str: A cleaned string or an empty string if not valid.
    """
    if pd.isna(value) or value is None or str(value).lower().strip() in {"", "nan", "na"}:
        return ""

    cleaned = str(value)
    cleaned = re.sub(r'[\u200b\u200c\u200d\uFEFF\xa0]', '', cleaned)
    return cleaned.strip()

def ask_user_choice(prompt: str, valid_options: list[str]) -> str:
    """
    Prompts the user to make a choice from a predefined list of valid options.
    Repeats the prompt until a valid input is received.

    Parameters:
        prompt (str): The message to display to the user.
        valid_options (list[str]): A list of accepted lowercase input values.

    Returns:
        str: The valid input provided by the user.
    """
    valid_options = [opt.lower() for opt in valid_options]
    while True:
        user_input = input(prompt).strip().lower()
        if user_input in valid_options:
            return user_input
        print(f"⚠️ Invalid input. Please select one of the following options: {', '.join(valid_options)}")

def parse_verse_number(value, fallback=-1):
    """
    Converts a given value (string, float, int) into a verse number as float.

    - Handles strings with commas or periods (e.g., "17,02" → 17.02).
    - Returns a float representing the verse number (e.g., "18.7" → 18.7).
    - If the value is invalid or cannot be parsed, returns the fallback (default: -1).

    Parameters:
        value (any): The input to be parsed as a verse number.
        fallback (float|int): Value to return if parsing fails.

    Returns:
        float: Parsed verse number, or fallback on failure.
    """
    try:
        return float(str(value).replace(",", ".").strip())
    except (ValueError, TypeError):
        return fallback

def is_same_verse_number(a, b, tolerance: float = 0.0001) -> bool:
    """
    Compares two verse numbers numerically within a given tolerance.

    - Accepts input as int, float, or string (with "." or "," as decimal separator).
    - Returns True if the absolute numeric difference between a and b is smaller than the tolerance.
    - If parsing fails for either value, returns False.

    Examples:
        is_same_verse_number("18", "18.00001") → True
        is_same_verse_number("18", "18.24")    → False
        is_same_verse_number("foo", 18)        → False

    Parameters:
        a (any): First verse number to compare.
        b (any): Second verse number to compare.
        tolerance (float): Allowed numeric deviation (default: 0.0001).

    Returns:
        bool: True if numbers are equal within tolerance, else False.
    """
    try:
        return abs(float(str(a).replace(",", ".")) - float(str(b).replace(",", "."))) < tolerance
    except (ValueError, TypeError):
        return False

def standardize_verse_number(entry):
    """
    Ensures that the 'Vers' field in a dictionary is stored as a float.

    This function is used to normalize verse values from JSON or Excel
    sources. It ensures that all 'Vers' fields are converted into consistent
    float representations, enabling correct sorting, comparison, and numeric logic.

    Examples:
        {"Vers": "15"}      → {"Vers": 15.0}
        {"Vers": "12,3"}    → {"Vers": 12.3}
        {"Vers": 18.75}     → {"Vers": 18.75}

    Parameters:
        entry (dict): A data dictionary that may contain a 'Vers' field.

    Returns:
        dict: A copy of the original dictionary with 'Vers' normalized as float (if present).
    """
    if isinstance(entry, dict) and "Vers" in entry:
        entry = entry.copy()
        entry["Vers"] = parse_verse_number(entry["Vers"])
    return entry

def sorted_entries(entries: list) -> list:
    """
    Returns a cleaned and consistently sorted list of entry dictionaries.

    Entries are:
    - filtered to include only those with a valid numeric 'Vers' value
    - sorted by:
        (1) verse number (numerically, including decimals),
        (2) the decimal part (e.g. 12.30 > 12.24),
        (3) the first non-empty string among 'Eigennennung', 'Bezeichnung', or 'Erzähler' (case-insensitive)

    Parameters:
        entries (list): A list of dictionaries representing naming or categorization entries.

    Returns:
        list: The cleaned and sorted list of entries.
    """

    def sort_key(entry):
        """
        Sorting key:
        - numerical verse number split into integer and decimal parts
        - alphabetical name resolution fallback
        """
        v = parse_verse_number(entry.get("Vers"))
        return (
            int(v),
            int(round((v % 1) * 100)),
            get_first_valid_text(
                entry.get("Eigennennung"),
                entry.get("Bezeichnung"),
                entry.get("Erzähler")
            ).strip().lower()
        )

    entries_clean = [
        e for e in deepcopy(entries)
        if isinstance(e, dict)
        and parse_verse_number(e.get("Vers")) != -1
        and not math.isnan(parse_verse_number(e.get("Vers")))
    ]

    return sorted(entries_clean, key=sort_key)