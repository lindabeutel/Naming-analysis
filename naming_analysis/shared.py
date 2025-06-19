"""
shared.py

General-purpose helper functions used throughout the project.
Includes text normalization, fallback selection, user interaction, and data cleaning utilities."""

import re
import pandas as pd

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
    Attempts to extract a valid verse number as integer from various formats (float, str, int).
    Example: '17.02' → 17; 18.9 → 18
    Returns fallback if conversion fails.
    """

    try:
        return float(str(value).replace(",", ".").strip())
    except (ValueError, TypeError):
        return fallback

def is_same_verse_number(a, b, tolerance: float = 0.0001) -> bool:
    """
    Compares two verse numbers (e.g. 18, 18.07, "18.24").
    Returns True if they are numerically equal within a small tolerance.
    Falls parsing fails, compares as cleaned strings.
    """
    try:
        return abs(float(str(a).replace(",", ".")) - float(str(b).replace(",", "."))) < tolerance
    except (ValueError, TypeError):
        return False
