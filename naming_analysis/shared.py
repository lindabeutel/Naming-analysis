"""
shared.py

Shared helper utilities for the naming-analysis pipeline.

Scope
-----
This module centralizes small, reusable building blocks that are used across
collection, analysis, exporting, and visualization code paths.

The helpers in this module are intentionally low-level:
they provide deterministic utilities (string/verse handling, parsing, formatting,
color conversions) and lightweight convenience functions used by multiple modules.

Included helper clusters
------------------------
- Text normalization and cleaning
  - token/lemma preprocessing, whitespace normalization, case handling
  - cell-value cleaning for pandas-origin values

- CLI interaction and parsing helpers
  - validated user-choice prompts
  - parsing of compact numeric selections (e.g., "1-3,5")
  - interactive figure-name resolution against categorization entries

- Verse parsing and standardization
  - tolerant parsing of verse identifiers (comma/dot decimals)
  - stable sorting of verse-based entries
  - serialization of verse values for exports

- Naming data preparation (BETA-stage source selection + guards)
  - selection of a preferred naming source (JSON first, Excel fallback)
  - tolerant detection of expected naming columns and trimming to relevant subsets
  - feature-flag reporting and path-specific requirement checks

- Color parsing and accessibility helpers (Plotly-related)
  - hex/rgb/rgba parsing and conversion utilities
  - WCAG-related luminance and contrast helpers
  - contrast-based text-color selection

- KWIC formatting utilities
  - lightweight KWIC splitting based on first variant match (substring search)

- Token extraction utilities
  - extraction/collection of naming-variant and epithet tokens from entries/rows

- Lemma heuristics (visualization support)
  - conservative helpers for treating lemmas as name-mentions
  - optional user confirmation for a suggested proper-name variant (CLI)

- Data discovery / filesystem utilities
  - deriving available reference-book identifiers from expected data-folder structure

Design principles
----------------
- Prefer stateless helpers whenever practical.
- Keep helpers granular and reusable; avoid coupling to specific analysis workflows.
- Provide deterministic behavior (stable ordering, consistent normalization).
- BETA-stage stance: minimal validation by default; hard requirements are enforced
  by explicit guards where needed.
- Interactive behavior (stdin/stdout) is limited to dedicated CLI helper functions
  and kept separate from pure utilities.
"""
# Standard library
import difflib
import math
import re
from copy import deepcopy
from pathlib import Path
from typing import Any

# Third-party libraries
import pandas as pd
from plotly.colors import hex_to_rgb, label_rgb

# Matches CSS-style rgb(...) and rgba(...) color strings
# Example: "rgb(12, 34, 56)" or "rgba(12, 34, 56, 0.5)"
_RGBA_RE = re.compile(
    r"^rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})(?:\s*,\s*([01](?:\.\d+)?))?\s*\)$",
    re.IGNORECASE,
)

# ---------------------------------------------------------------------------
# Text normalization and cleaning utilities
# ---------------------------------------------------------------------------

def normalize_text(text):
    """
    Normalize text for internal matching and comparison.

    Processing steps:
    - Lowercase conversion
    - Replacement of selected diacritics and ligatures
    - Project-specific substitutions (e.g., "iu" → "ie", "üe" → "ue")
    - Replacement of standalone "v" with "f"
    - Whitespace normalization (collapse multiple spaces)

    Parameters:
        text (str | None): Input string.

    Returns:
        str:
            Always returns a string.
            If `text` is falsy (None, empty string, etc.), returns "".
            The function does not raise exceptions for missing input.
    """
    # Project-specific character and grapheme substitutions
    substitutions = {
        'æ': 'ae', 'œ': 'oe',
        'é': 'e', 'è': 'e', 'ë': 'e', 'á': 'a', 'à': 'a',
        'û': 'u', 'î': 'i', 'â': 'a', 'ô': 'o', 'ê': 'e',
        'ü': 'u', 'ö': 'o', 'ä': 'a',
        'ß': 'ss',
        'iu': 'ie', 'üe': 'ue'
    }

    # Return empty string for falsy input
    if not text:
        return ""

    # Case-folding for consistent internal matching
    text = text.lower()

    # Apply configured substitutions sequentially
    for old, new in substitutions.items():
        text = text.replace(old, new)

    # Replace standalone "v" with "f" (editorial normalization rule)
    text = re.sub(r'\bv\b', 'f', text)

    # Collapse multiple whitespace characters into a single space
    text = re.sub(r'\s+', ' ', text)

    return text

def get_first_valid_text(*fields):
    """
    Return the first non-empty string from a sequence of values.

    Only values that are instances of `str` and contain non-whitespace
    characters are considered valid. All other types (including None and NaN)
    are ignored.

    Parameters:
        *fields: One or more values to evaluate.

    Returns:
        str:
            The first valid non-empty string.
            Returns "" if no suitable value is found.
    """
    for f in fields:
        # Accept only non-empty strings (ignore None, NaN, and other types)
        if isinstance(f, str) and f.strip():
            return f

    # No valid string found → return empty string
    return ""

def clean_cell_value(value):
    """
    Normalize a DataFrame cell value for internal processing.

    Behavior:
    - If the value is missing (`None` or `pd.isna(...)`), returns "".
    - Otherwise, converts the value to string, strips surrounding whitespace,
      and applies `normalize_text(...)`.

    Parameters:
        value: Cell content (e.g., from a pandas DataFrame).

    Returns:
        str:
            Always returns a string.
            Returns "" for missing values.
    """
    # Treat pandas NA/NaN and explicit None as missing values
    if pd.isna(value) or value is None:
        return ""

    # Convert to string, strip whitespace, then normalize
    return normalize_text(str(value).strip())

def sanitize_cell_value(value):
    """
    Remove invisible Unicode characters and filter placeholder artifacts.

    Behavior:
    - Returns "" for missing values (`None`, `pd.isna(...)`) or placeholder
      strings such as "", "nan", or "na".
    - Removes zero-width and non-breaking space characters.
    - Preserves original casing (no normalization applied).

    Parameters:
        value: Cell content (e.g., from a pandas DataFrame).

    Returns:
        str:
            Cleaned string with invisible characters removed,
            or "" if the value is considered invalid.
    """
    # Treat pandas NA/NaN, None, and common placeholder strings as invalid
    if pd.isna(value) or value is None or str(value).lower().strip() in {"", "nan", "na"}:
        return ""

    # Convert to string without applying normalization
    cleaned = str(value)

    # Remove zero-width characters and non-breaking spaces
    cleaned = re.sub(r'[\u200b\u200c\u200d\uFEFF\xa0]', '', cleaned)

    # Strip surrounding whitespace and return
    return cleaned.strip()

# ---------------------------------------------------------------------------
# CLI interaction helpers
# ---------------------------------------------------------------------------

def ask_user_choice(prompt: str, valid_options: list[str]) -> str:
    """
    Prompt the user to select one of the allowed options.

    Behavior:
    - Input is stripped and converted to lowercase.
    - Matching is case-insensitive.
    - The prompt repeats until a valid option is entered.

    Parameters:
        prompt (str): Message displayed to the user.
        valid_options (list[str]): Allowed input values.

    Returns:
        str:
            The validated user input (always lowercase).
    """
    # Normalize allowed options for case-insensitive comparison
    valid_options = [opt.lower() for opt in valid_options]
    while True:
        user_input = input(prompt).strip().lower()
        if user_input in valid_options:
            return user_input
        print(f"Invalid input. Please select one of the following options: {', '.join(valid_options)}")

def parse_token_selection(input_str: str, max_value: int) -> list[int] | None:
    """
    Parse a compact numeric selection string into a sorted list of unique indices.

    This helper supports typical CLI selection syntax for choosing items by number.

    Accepted input forms
    --------------------
    - Single indices: "3"
    - Ranges: "1-3"
    - Mixed lists: "1-3,5,7"
    - Unicode en dash "–" is treated like "-" (normalized before parsing).
    - Whitespace is ignored.

    Indexing convention
    -------------------
    Indices are 1-based and inclusive (i.e., "1-3" expands to [1, 2, 3]).

    Validation rules
    ---------------
    - Returns None if the input is empty/whitespace-only.
    - Returns None if any segment is malformed (non-integer tokens).
    - Returns None if any parsed index is out of range (must satisfy 1 <= idx <= max_value).
    - Returns None if a range is inverted (start > end).

    Parameters
    ----------
    input_str : str
        Raw user input string (e.g., from a CLI prompt).
    max_value : int
        Maximum allowed index (typically the number of available options).
        The valid domain is 1..max_value (inclusive).

    Returns
    -------
    list[int] | None
        Sorted list of unique selected indices (1-based), or None if validation fails.
    """
    # Empty input is treated as "no selection" (caller decides what that means).
    if not input_str.strip():
        return None

    # Normalize common human input: en dash → hyphen, remove whitespace for simpler parsing.
    input_str = input_str.replace("–", "-").replace(" ", "")

    # Split by comma to allow mixed selections like "1-3,5,7".
    parts = input_str.split(",")

    # Use a set to de-duplicate indices across overlapping ranges and repeated values.
    result = set()

    for part in parts:
        # Range segment (inclusive), e.g., "2-4".
        if "-" in part:
            try:
                start_str, end_str = part.split("-", 1)
                start = int(start_str)
                end = int(end_str)

                # Reject inverted ranges and out-of-range endpoints.
                if start > end or start < 1 or end > max_value:
                    return None

                # Expand inclusive range and add to the selection set.
                result.update(range(start, end + 1))
            except ValueError:
                # Non-integer range bounds (e.g., "a-3") → invalid input.
                return None
        else:
            # Single index segment, e.g., "5".
            try:
                value = int(part)

                # Reject out-of-range indices.
                if 1 <= value <= max_value:
                    result.add(value)
                else:
                    return None
            except ValueError:
                # Non-integer token (e.g., "x") → invalid input.
                return None

    # Stable output: sorted, unique indices.
    return sorted(result)

def resolve_figure_name(name: str, entries: list[dict]) -> str | None:
    """
    Resolve a user-provided figure name against categorization entries.

    The function attempts to match the input name to existing values in the
    "Benannte Figur" field of the provided entries.

    Resolution strategy
    --------------------
    1. Collect all distinct, non-empty figure names from the entries.
    2. If the input name matches one of them exactly, return it unchanged.
    3. Otherwise, compute the closest match using `difflib.get_close_matches`.
    4. If a suggestion is found, prompt the user for confirmation (y/n).
    5. Return the confirmed suggestion or None if rejected or no suggestion exists.

    Side effects
    ------------
    - Prints diagnostic messages to stdout.
    - Prompts the user for confirmation via `ask_user_choice`.

    Parameters
    ----------
    name : str
        The figure name entered by the user (compared verbatim).
    entries : list[dict]
        Categorization entries containing the key "Benannte Figur".

    Returns
    -------
    str | None
        The resolved canonical figure name if confirmed,
        otherwise None.
    """
    # Build a set of all distinct, non-empty figure names found in the entries.
    # Whitespace is stripped during extraction.
    all_names = {
        str(name).strip()
        for name in [e.get("Benannte Figur") for e in entries]
        if isinstance(name, str) and name.strip()
    }

    # Direct exact match (no fuzzy logic required).
    if name in all_names:
        return name

    # Fuzzy similarity-based suggestion (single best candidate).
    suggestions = difflib.get_close_matches(name, all_names, n=1, cutoff=0.6)

    if suggestions:
        # Inform the user that no exact match was found.
        print(f'Figure "{name}" not found.')
        print(f'Did you mean "{suggestions[0]}"? [y/n]')

        # Ask for confirmation of the suggested match.
        answer = ask_user_choice("> ", ["y", "n"])
        if answer == "y":
            return suggestions[0]
        else:
            # User rejected suggestion.
            print("No valid figure selected.")
            print("Please enter a valid name exactly as it appears in your categorization data.")
            return None

    else:
        # No sufficiently similar candidate found.
        print(f'Figure "{name}" not found and no similar name could be suggested.')
        return None

# ---------------------------------------------------------------------------
# Verse parsing and standardization utilities
# ---------------------------------------------------------------------------

def parse_verse_number(value, fallback=-1):
    """
    Parse a verse identifier into a numeric value (float).

    Behavior:
    - Converts the input to string, strips whitespace, and replaces "," with "."
      to support decimal verse identifiers (e.g., "17,02" → 17.02).
    - Returns a float if parsing succeeds.
    - Returns `fallback` unchanged if parsing fails.

    Parameters:
        value: Raw verse identifier (string/number).
        fallback (float | int): Value returned if parsing fails (default: -1).

    Returns:
        float | int:
            Parsed float if conversion succeeds.
            Returns the provided `fallback` unchanged if conversion fails.

        The function does not raise parsing-related exceptions.
    """
    try:
        # Normalize decimal separator and whitespace, then convert to float
        return float(str(value).replace(",", ".").strip())
    except (ValueError, TypeError):
        # Return fallback if conversion fails
        return fallback

def is_same_verse_number(a, b, tolerance: float = 0.0001) -> bool:
    """
    Compare two verse identifiers numerically within a tolerance.

    Behavior:
    - Converts both inputs to float after normalizing decimal separators
      ("," → ".").
    - Returns True if the absolute difference is smaller than `tolerance`.
    - Returns False if either value cannot be parsed.

    Parameters:
        a: First verse identifier (string/int/float).
        b: Second verse identifier (string/int/float).
        tolerance (float): Maximum allowed numeric deviation (default: 0.0001).

    Returns:
        bool:
            True if both values are numerically equal within tolerance
            (absolute difference strictly smaller than `tolerance`).

            Returns False if parsing fails for either value.
            The function does not raise parsing-related exceptions.
    """
    # Normalize decimal separators, convert both values to float,
    # and compare absolute difference against tolerance
    try:
        return abs(float(str(a).replace(",", ".")) - float(str(b).replace(",", "."))) < tolerance
    except (ValueError, TypeError):
        # Return False if either value cannot be parsed
        return False

def standardize_verse_number(entry):
    """
    Normalize the "Vers" field of a dictionary to a float.

    Behavior:
    - If `entry` is a dict containing key "Vers", a shallow copy is created.
    - The value of "Vers" is converted using `parse_verse_number(...)`.
    - If parsing fails, the fallback value from `parse_verse_number`
      is applied (default: -1).
    - If no "Vers" field is present, the original object is returned unchanged.

    Parameters:
        entry (dict): Data dictionary that may contain a "Vers" field.

    Returns:
        dict:
            If "Vers" is present, returns a new (shallow-copied) dictionary
            with normalized "Vers".

            If "Vers" is not present, returns the original object unchanged.

        The function does not raise parsing-related exceptions.
    """
    # Normalize only dict entries that contain a "Vers" field
    if isinstance(entry, dict) and "Vers" in entry:
        # Avoid mutating the input mapping
        entry = entry.copy()
        # Parse verse identifier into a float (fallback handled by parse_verse_number)
        entry["Vers"] = parse_verse_number(entry["Vers"])
    return entry

def sorted_entries(entries: list) -> list:
    """
    Filter and consistently sort verse-based entry dictionaries.

    Behavior:
    - Creates a deep copy of the input list.
    - Filters out entries where "Vers" cannot be parsed into a valid number
      (i.e., parse result equals -1 or is NaN).
    - Sorts remaining entries by:
        (1) integer part of the verse number,
        (2) decimal part (scaled to two digits),
        (3) first non-empty text among "Eigennennung",
            "Bezeichnung", or "Erzähler" (case-insensitive).

    Parameters:
        entries (list): List of dictionaries that may contain a "Vers" field.

    Returns:
        list:
            A new list containing filtered and sorted entry dictionaries.
            The input list is not modified.

    Notes:
        Sorting uses the parsed numeric verse value, but entries are not mutated.
        Parsing-related errors are handled by `parse_verse_number(...)` (fallback -1).
    """
    def sort_key(entry):
        """
        Sorting key composed of:
        - integer part of the parsed verse number,
        - decimal part scaled to two digits,
        - case-insensitive lexical fallback based on
          "Eigennennung", "Bezeichnung", or "Erzähler".
        """
        # Parse verse identifier once for sorting
        v = parse_verse_number(entry.get("Vers"))

        return (
            # Integer part of verse number
            int(v),

            # Decimal part scaled to two digits (e.g. 12.30 → 30)
            int(round((v % 1) * 100)),

            # Case-insensitive lexical fallback for identical verse numbers
            get_first_valid_text(
                entry.get("Eigennennung"),
                entry.get("Bezeichnung"),
                entry.get("Erzähler")
            ).strip().lower()
        )

    # Deep copy to avoid mutating the input list/dicts
    entries_clean = []
    for e in deepcopy(entries):
        if not isinstance(e, dict):
            continue

        verse_number = parse_verse_number(e.get("Vers"))

        # Keep only dict entries with valid numeric verse numbers
        if verse_number != -1 and not math.isnan(float(verse_number)):
            entries_clean.append(e)

    # Return sorted copy
    return sorted(entries_clean, key=sort_key)

# ---------------------------------------------------------------------------
# Data preparation helpers
# ---------------------------------------------------------------------------

def select_naming_data(
    book_name: str,
    df_json: pd.DataFrame | None,
    df_excel: pd.DataFrame | None,
) -> dict[str, Any]:
    """
    This function selects the preferred naming source (JSON first, Excel fallback),
    detects and normalizes the relevant columns, and returns a compact description
    of what is available for downstream analysis paths.

    Behavior:
    - Prefers JSON if it satisfies the minimum structural requirements.
    - Falls back to Excel if JSON is missing/empty or structurally/content-wise invalid.
    - Detects required columns with tolerant matching (case/whitespace-insensitive).
    - Validates minimum structure:
        - mandatory: 'Benannte Figur', 'Nennende Figur'
        - at least one lemma column: Bezeichnung*/Epitheta*
    - Applies an additional content-level consistency check for JSON sources only.
    - Returns a trimmed DataFrame, a canonical column mapping, explicit feature flags,
      and diagnostic messages (no printing).

    Parameters:
        book_name (str): Label used for error reporting.
        df_json (pd.DataFrame | None): JSON-derived naming data (preferred if valid).
        df_excel (pd.DataFrame | None): Excel-derived naming data (fallback if JSON invalid/missing).

    Returns:
        dict[str, Any]:
            {
                "source": "json" | "excel",
                "df": pd.DataFrame,            # trimmed copy restricted to relevant columns
                "cols": {                      # canonical detected column names
                    "target": <str>,
                    "namer": <str>,
                    "naming_variant_cols": list[str],
                    "epithet_cols": list[str],
                    "verse_col": str | None,
                    "has_unnumbered_naming_variant": bool
                },
                "features": {                  # explicit availability flags
                    "has_target": bool,
                    "has_namer": bool,
                    "has_naming_variants": bool,
                    "has_epithets": bool,
                    "has_content": bool,
                    "has_verse": bool
                },
                "messages": list[str]          # diagnostic notes about selection/fallback
            }

    Raises:
        ValueError:
            If neither JSON nor Excel contains the minimum required structure for `book_name`.

    Notes:
        - The function never returns None. If no valid source can be selected, it raises.
        - The returned DataFrame is a new copy and does not mutate the input DataFrames.
        - Path-specific adequacy checks (hard fail vs. warnings) belong to downstream requirement guards.
    """
    def normalize(colname):
        """
        Normalize column names for tolerant matching.

        Behavior:
        - Converts to lowercase.
        - Strips leading/trailing whitespace.
        - Collapses multiple internal whitespace characters to a single space.
        - Returns "" for non-string inputs.
        """
        return re.sub(r"\s+", " ", colname.strip().lower()) if isinstance(colname, str) else ""

    def detect_columns(df):
        """
        Detect mandatory and dynamic naming-related columns in a DataFrame.

        Behavior:
        - Applies tolerant column matching (case/whitespace-insensitive).
        - Identifies mandatory columns:
            "Benannte Figur" and "Nennende Figur".
        - Detects dynamic lemma columns:
            Bezeichnung*, Epitheta*, and Vers/Vers-ID.
        - Determines whether an unnumbered "Bezeichnung" column exists.
        - Returns a structured dictionary describing detected columns.

        Returns:
            dict: Mapping with detected column names and structural flags.
        """
        # Build a normalized lookup to compare columns case/whitespace-insensitively.
        nmap = {c: normalize(c) for c in df.columns}

        def match(regex):
            """
            Return all column names that fully match the given regex.

            Matching is performed against both:
            - the original column name, and
            - the normalized version (case/whitespace-insensitive).
            """
            rx = re.compile(regex, re.IGNORECASE)

            # Full-match against original and normalized column names
            return [c for c in df.columns if rx.fullmatch(c) or rx.fullmatch(nmap[c])]

        # Detect mandatory structural columns (tolerant to spacing/underscore variants)
        target = None
        namer = None

        for c in df.columns:
            normed = nmap[c]
            # "Benannte Figur" (target) may appear with minor formatting variants.
            if normed in ("benannte figur", "benannte_figur"):
                target = c
            # "Nennende Figur" (namer) may appear with minor formatting variants.
            if normed in ("nennende figur", "nennende_figur"):
                namer = c

        # Dynamic lemma-related columns:
        # - Bezeichnung*, Epitheta* (optionally numbered)
        # - Vers / Vers-ID variants (optional)
        naming_variant_cols = match(r"bezeichnung(\s*\d+)?")
        epithet_cols     = match(r"epitheta(\s*\d+)?")
        verse_cols       = match(r"vers(|-?id)?")

        # Structural flags used for later validation/feature reporting.
        has_unnumbered   = any(normalize(c) == "bezeichnung" for c in naming_variant_cols)

        # True if at least one lemma-related column is present
        has_lex          = (len(naming_variant_cols) + len(epithet_cols)) > 0

        cols = {
            "target": target,
            "namer": namer,
            "naming_variant_cols": naming_variant_cols,
            "epithet_cols": epithet_cols,
            "verse_col": verse_cols[0] if verse_cols else None,
            "has_unnumbered_naming_variant": has_unnumbered,
            "has_lex": has_lex,
        }
        return cols

    def trim(df, cols):
        """
        Return a copy of the DataFrame restricted to relevant columns.

        Behavior:
        - Keeps only detected structural and lemma-related columns.
        - Preserves original column order.
        - Removes duplicate column references while maintaining order.
        - Returns a new DataFrame (does not modify input).

        Parameters:
            df (pd.DataFrame): Source DataFrame.
            cols (dict): Column mapping returned by detect_columns(...).

        Returns:
            pd.DataFrame: Trimmed copy containing only relevant columns.
        """
        keep = []

        # Keep mandatory structural columns first if present.
        if cols["target"]:
            keep.append(cols["target"])
        if cols["namer"]:
            keep.append(cols["namer"])

        # Keep dynamic lemma-related columns next.
        keep += cols["naming_variant_cols"]
        keep += cols["epithet_cols"]

        # Keep verse column last if present.
        if cols["verse_col"]:
            keep.append(cols["verse_col"])

        # Deduplicate while preserving order and ensuring columns exist in df.
        keep = [c for c in dict.fromkeys(keep) if c in df.columns]

        # Return trimmed copy (avoid mutating original DataFrame)
        return df.loc[:, keep].copy()

    def build_features(cols):
        has_target = bool(cols.get("target"))
        has_namer = bool(cols.get("namer"))
        has_naming_variants = bool(cols.get("naming_variant_cols"))
        has_epithets = bool(cols.get("epithet_cols"))
        has_content = has_naming_variants or has_epithets
        has_verse = bool(cols.get("verse_col"))

        return {
            "has_target": has_target,
            "has_namer": has_namer,
            "has_naming_variants": has_naming_variants,
            "has_epithets": has_epithets,
            "has_content": has_content,
            "has_verse": has_verse,
        }

    # Collect diagnostics for CLI/UI; the selector itself does not print.
    messages = []

    # --- Source selection: prefer JSON if structurally adequate, else fall back to Excel. ---

    # 1) JSON path (preferred)
    if df_json is None or df_json.empty:
        # JSON missing/empty is not an error by itself; it triggers Excel fallback.
        messages.append("No JSON data provided or JSON is empty – loading from Excel instead.")
    else:
        # Detect structure and dynamic lemma columns in the JSON-derived DataFrame.
        cj = detect_columns(df_json)

        # Minimum structural requirements for selecting JSON:
        # - target + namer must exist
        # - at least one lemma-related column group must exist (Bezeichnung*/Epitheta*)
        json_ok = cj["target"] and cj["namer"] and cj["has_lex"]

        if not json_ok:
            # Report which structural components are missing to make fallback transparent.
            missing_parts = []
            if not cj["target"]:
                missing_parts.append("'Benannte Figur'")
            if not cj["namer"]:
                missing_parts.append("'Nennende Figur'")
            if not cj["has_lex"]:
                missing_parts.append("Bezeichnung*/Epitheta*")

            messages.append(f"JSON missing {', '.join(missing_parts)} – loading from Excel instead.")
        else:
            # Additional JSON-only adequacy check:
            # If lemma columns are populated, verify that rows provide attribution
            # via at least one of: Eigennennung, Erzähler, or non-empty Nennende Figur.
            lemma_cols = [c for c in (cj["naming_variant_cols"] + cj["epithet_cols"]) if c in df_json.columns]

            # Optional attribution columns (may not exist in every JSON export).
            eigennennung_col = "Eigennennung" if "Eigennennung" in df_json.columns else None
            erzaehler_col = "Erzähler" if "Erzähler" in df_json.columns else None

            if lemma_cols:
                # Helper: detect non-empty cells after trimming
                def _nonempty_series(s):
                    """
                    Return a boolean Series indicating non-empty string values.

                    A value is considered non-empty if:
                    - it is not NaN,
                    - and its string representation (after stripping) is not "".
                    """
                    return (
                        s.fillna("")
                         .astype(str)
                         .str.strip()
                         .ne("")
                    )

                # Rows containing at least one lemma entry (Bezeichnung/Epitheta).
                has_lemma = df_json[lemma_cols].apply(_nonempty_series).any(axis=1)

                # Rows containing Eigennennung (if column exists).
                has_eigennennung = (
                    _nonempty_series(df_json[eigennennung_col])
                    if eigennennung_col
                    else pd.Series(False, index=df_json.index)
                )

                # Rows containing Erzähler (if column exists).
                has_erzaehler = (
                    _nonempty_series(df_json[erzaehler_col])
                    if erzaehler_col
                    else pd.Series(False, index=df_json.index)
                )

                # Rows with a non-empty naming figure attribution ("Nennende Figur").
                namer_col = cj["namer"]
                namer_nonempty = (
                    df_json[namer_col]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .ne("")
                )

                # Rows are considered structurally inconsistent if lemma content exists
                # but no attribution is present in any of the supported fields.
                bad_rows = (
                        has_lemma
                        & ~has_eigennennung
                        & ~has_erzaehler
                        & ~namer_nonempty
                )

                bad_count = int(bad_rows.sum())

                if bad_count > 0:
                    # Reject JSON source for this run; Excel may provide a cleaner structure.
                    messages.append(
                        f"JSON has {bad_count} row(s) with Bezeichnung/Epitheton but empty 'Nennende Figur' — loading Excel instead."
                    )
                else:
                    # Accept JSON: return trimmed df + canonical mapping + features + diagnostics.
                    df_trimmed = trim(df_json, cj)
                    cols_out = {
                        "target": cj["target"],
                        "namer": cj["namer"],
                        "naming_variant_cols": cj["naming_variant_cols"],
                        "epithet_cols": cj["epithet_cols"],
                        "verse_col": cj["verse_col"],
                        "has_unnumbered_naming_variant": cj["has_unnumbered_naming_variant"],
                    }
                    return {
                        "source": "json",
                        "df": df_trimmed,
                        "cols": cols_out,
                        "features": build_features(cols_out),
                        "messages": messages,
                    }
            else:
                # JSON meets minimum structure and has no lemma columns to content-validate.
                df_trimmed = trim(df_json, cj)
                cols_out = {
                    "target": cj["target"],
                    "namer": cj["namer"],
                    "naming_variant_cols": cj["naming_variant_cols"],
                    "epithet_cols": cj["epithet_cols"],
                    "verse_col": cj["verse_col"],
                    "has_unnumbered_naming_variant": cj["has_unnumbered_naming_variant"],
                }
                return {
                    "source": "json",
                    "df": df_trimmed,
                    "cols": cols_out,
                    "features": build_features(cols_out),
                    "messages": messages,
                }

    # 2) Excel fallback (only if JSON was not accepted)
    if df_excel is not None and not df_excel.empty:
        cx = detect_columns(df_excel)

        # Excel must satisfy the same minimum structural requirements (no content-level check).
        x_ok = cx["target"] and cx["namer"] and cx["has_lex"]

        if x_ok:
            messages.append("Using Excel fallback (JSON lacked required structure).")
            df_trimmed = trim(df_excel, cx)
            cols_out = {
                "target": cx["target"],
                "namer": cx["namer"],
                "naming_variant_cols": cx["naming_variant_cols"],
                "epithet_cols": cx["epithet_cols"],
                "verse_col": cx["verse_col"],
                "has_unnumbered_naming_variant": cx["has_unnumbered_naming_variant"],
            }
            return {
                "source": "excel",
                "df": df_trimmed,
                "cols": cols_out,
                "features": build_features(cols_out),
                "messages": messages,
            }

    # No acceptable source: raise an explicit structural error.
    raise ValueError(
        f"[select_naming_data] Missing required columns in JSON/Excel for '{book_name}'. "
        "Expected: 'Benannte Figur', 'Nennende Figur', and at least one of Bezeichnung*/Epitheta*."
    )

def prepare_naming_data(book_name, df_json, df_excel):
    """
    Backward-compatible wrapper around select_naming_data(...).

    This function preserves the historical return signature used across
    analysis paths while delegating all source-selection logic to the
    canonical selector.

    Behavior:
    - Calls select_naming_data(...) to perform source selection,
      structural detection, and feature derivation.
    - Prints diagnostic selection messages to stdout (CLI behavior).
    - Returns only the legacy tuple (source, df_trimmed, cols).

    Parameters:
        book_name (str): Label used for error reporting.
        df_json (pd.DataFrame | None): JSON-derived naming data.
        df_excel (pd.DataFrame | None): Excel-derived naming data.

    Returns:
        tuple[str, pd.DataFrame, dict]:
            (source, df_trimmed, cols)

            - source: "json" or "excel"
            - df_trimmed: DataFrame restricted to relevant naming columns
            - cols: canonical column mapping as defined by select_naming_data(...)

    Raises:
        ValueError:
            Propagated from select_naming_data(...) if no valid source exists.

    Notes:
        - This wrapper exists for API stability within analysis.py.
        - New code should prefer select_naming_data(...) directly if
          feature flags or diagnostic messages are required.
    """
    # Delegate full selection and validation logic to the canonical selector.
    selection = select_naming_data(book_name, df_json, df_excel)


    # Preserve historical CLI behavior: emit selector diagnostics here.
    # The selector itself does not print.
    for msg in selection.get("messages", []):
        print(msg)

    # Return legacy tuple interface for existing analysis functions.
    return selection["source"], selection["df"], selection["cols"]

def check_naming_requirements(selection: dict[str, Any], *, require_target: bool = True,
                              require_namer: bool = True, require_content: bool = True,
                              require_naming_variants: bool = False, require_epithets: bool = False,
                              context: str = "") -> None:
    """
    Path-specific requirement guard for naming analyses.

    This function validates whether the already-selected naming data
    (as returned by select_naming_data(...)) satisfies the minimum
    structural/content requirements of a specific analysis path.

    It does NOT:
    - select a source,
    - modify the DataFrame,
    - print diagnostics.

    It only inspects the descriptive feature flags stored in
    selection["features"] and raises a ValueError if required
    elements are missing.

    Parameters:
        selection (dict[str, Any]):
            Result object returned by select_naming_data(...).
            Must contain a "features" mapping.

        require_target (bool):
            Require presence of a 'Benannte Figur' column.

        require_namer (bool):
            Require presence of a 'Nennende Figur' column.

        require_content (bool):
            Require at least one lemma-related column
            (Bezeichnung* or Epitheta*).

        require_naming_variants (bool):
            Require at least one Bezeichnung* column.

        require_epithets (bool):
            Require at least one Epitheta* column.

        context (str):
            Optional label added to the error message
            to clarify which analysis path triggered the failure.

    Raises:
        ValueError:
            If one or more required structural components are missing.

    Notes:
        - This function is intentionally minimal and declarative.
        - Feature flags are descriptive (what exists), not normative.
        - Normative adequacy is defined here at call-site level.
    """
    # Extract descriptive feature flags from the selector result.
    features = selection["features"]

    # Collect missing requirements to produce a single, clear error message.
    missing = []

    # Mandatory structural elements (if requested).
    if require_target and not features.get("has_target", False):
        missing.append("'Benannte Figur'")
    if require_namer and not features.get("has_namer", False):
        missing.append("'Nennende Figur'")

    # At least one lemma-related column (Bezeichnung*/Epitheta*).
    if require_content and not features.get("has_content", False):
        missing.append("Bezeichnung*/Epitheta*")

    # More specific requirements (path-dependent).
    if require_naming_variants and not features.get("has_naming_variants", False):
        missing.append("Bezeichnung*")
    if require_epithets and not features.get("has_epithets", False):
        missing.append("Epitheta*")

    # If any requirement failed, raise a single explicit structural error.
    if missing:
        prefix = f"[{context}] " if context else ""
        raise ValueError(prefix + "Analysis cannot proceed — missing required data: " + ", ".join(missing))

def serialize_verse_value(value) -> str:
    """
    Convert verse identifiers into a consistent export-friendly string.

    Behavior:
    - Returns "" for None or NaN.
    - Preserves original string formatting where possible.
      (e.g., Excel-style comma decimals such as "17,2".)
    - Removes redundant trailing zeros from dot-based decimals.
    - Ensures numeric values (int/float) are serialized without
      unnecessary decimal artifacts (e.g., 20.0 → "20").

    Parameters:
        value: Verse value from JSON, Excel, or internal processing.

    Returns:
        str:
            Always returns a string representation suitable for export.

            - Returns "" for None or NaN.
            - Returns a cleaned numeric string for int/float values.
            - Falls back to str(value) if numeric conversion fails.

        The function does not raise conversion-related exceptions.
    """
    # Missing values → empty export cell
    if value is None:
        return ""
    try:
        # pandas-aware NaN check (for DataFrame-origin values)
        if pd.isna(value):
            return ""
    except (NameError, AttributeError):
        # Allow use without pandas or with non-pandas inputs
        pass

    # Preserve source formatting for string inputs
    if isinstance(value, str):
        val = value.strip()

        # Keep Excel-style comma decimals unchanged (e.g., "17,2")
        if "," in val:
            return val

        # Trim redundant zeros from dot-decimals (e.g., "17.20" → "17.2")
        if "." in val:
            val = val.rstrip("0").rstrip(".")

        return val

    # Handle numeric values (int/float-like) without trailing ".0" artifacts
    try:
        f = float(value)

        # Serialize integer-like floats as integers (e.g., 20.0 → "20")
        if f.is_integer():
            return str(int(f))

        # Trim redundant zeros from float string representation
        s = str(f)
        if "." in s:
            s = s.rstrip("0").rstrip(".")

        return s

    except (ValueError, TypeError):
        # Last-resort fallback: stringify unknown types
        return str(value)

# ---------------------------------------------------------------------------
# Color and accessibility utilities (Plotly-related)
# ---------------------------------------------------------------------------

def hex_color_to_rgb_tuple(hex_color: str) -> tuple[int, int, int]:
    """
    Convert a hexadecimal color string into an integer RGB tuple (0–255).

    This function delegates parsing to `plotly.colors.hex_to_rgb`
    and wraps its return values into an explicitly typed
    `tuple[int, int, int]`.

    Parameters
    ----------
    hex_color : str
        Hex color in canonical form '#RRGGBB'.

    Returns
    -------
    tuple[int, int, int]
        (r, g, b) as integers in the range 0–255.

    Raises
    ------
    Exception
        Any exception raised by `plotly.colors.hex_to_rgb`
        (e.g. invalid format) is propagated unchanged.

    Notes
    -----
    This utility is Plotly-dependent at parsing level but does not
    encode any visualization semantics. It only standardizes the
    return type for internal use.
    """
    # Delegate hex parsing to Plotly's canonical implementation.
    # Any format validation (e.g. malformed '#RRGGBB') happens there.
    r, g, b = hex_to_rgb(hex_color)

    # Explicit int-cast ensures a stable and predictable return type
    # (tuple[int, int, int]) for downstream WCAG/contrast utilities.
    return int(r), int(g), int(b)

def rgb_tuple_to_plotly_color(rgb: tuple[int, int, int]) -> str:
    """
    Convert an RGB tuple (0–255) to a Plotly color string.

    Parameters
    ----------
    rgb : tuple[int, int, int]
        Integer RGB values in the range 0–255.

    Returns
    -------
    str
        'rgb(r,g,b)' formatted string for Plotly usage.

    Notes
    -----
    Thin adapter around Plotly's `label_rgb`.
    No validation or normalization is performed.
    """
    # Delegate formatting to Plotly's `label_rgb`
    # → returns 'rgb(r,g,b)' string
    return label_rgb(rgb)

def plotly_color_to_rgb_tuple(color: str) -> tuple[int, int, int]:
    """
    Convert a Plotly-compatible color string to an integer RGB tuple (0–255).

    Accepted formats
    ----------------
    - '#RRGGBB'
    - 'rgb(r,g,b)'
    - 'rgba(r,g,b,a)' (alpha ignored)

    Returns
    -------
    tuple[int, int, int]
        (r, g, b) values clamped to the range 0–255.

    Raises
    ------
    ValueError
        If the color string does not match a supported format.
    """
    # Normalize input to string and remove surrounding whitespace
    c = str(color).strip()

    # Hex colors are delegated to the dedicated converter
    if c.startswith("#"):
        return hex_color_to_rgb_tuple(c)

    # Match rgb(...) / rgba(...) using precompiled regex
    m = _RGBA_RE.match(c)
    if not m:
        raise ValueError(f"Unsupported color format: {color!r}")

    # Extract and clamp channel values to valid RGB range (0–255)
    r = max(0, min(255, int(m.group(1))))
    g = max(0, min(255, int(m.group(2))))
    b = max(0, min(255, int(m.group(3))))
    return r, g, b

def hex_color_to_rgba(hex_color: str, alpha: float) -> str:
    """
    Convert a hex color ('#RRGGBB') to a Plotly-compatible 'rgba(...)' string.

    Parameters
    ----------
    hex_color : str
        Hex color in canonical form '#RRGGBB'.
    alpha : float
        Opacity value (expected range 0.0–1.0).
        No validation or clamping is performed.

    Returns
    -------
    str
        Formatted 'rgba(r,g,b,a)' string.

    Notes
    -----
    Delegates RGB parsing to `hex_color_to_rgb_tuple`.
    Pure format conversion without visualization semantics.
    """
    # Delegate hex parsing and validation to existing RGB utility
    r, g, b = hex_color_to_rgb_tuple(hex_color)

    # Format as Plotly-compatible rgba() string (no alpha validation)
    return f"rgba({r},{g},{b},{alpha})"

def srgb_channel_to_linear(c: float) -> float:
    """
    Convert a normalized sRGB channel value to linear RGB.

    Parameters
    ----------
    c : float
        sRGB channel value in the range 0.0–1.0.

    Returns
    -------
    float
        Linearized channel value.

    Notes
    -----
    Implements the standard WCAG sRGB companding inverse.
    No range validation or clamping is performed.
    """
    # Linear segment of the sRGB transfer function
    if c <= 0.04045:
        return c / 12.92

    # Non-linear gamma correction segment
    return ((c + 0.055) / 1.055) ** 2.4

def relative_luminance_from_rgb(rgb: tuple[int, int, int]) -> float:
    """
    Compute WCAG relative luminance from integer RGB values (0–255).

    Parameters
    ----------
    rgb : tuple[int, int, int]
        (r, g, b) channel values in the range 0–255.

    Returns
    -------
    float
        Relative luminance in the range 0.0–1.0.

    Notes
    -----
    Performs sRGB → linear conversion internally.
    Uses WCAG coefficients: 0.2126 (R), 0.7152 (G), 0.0722 (B).
    No input validation or clamping is performed.
    """
    # Convert integer channel (0–255) to normalized sRGB (0.0–1.0)
    def to_linear(v: int) -> float:
        """
        Convert a single 8-bit sRGB channel (0–255) to linear RGB.

        Applies the WCAG sRGB companding inverse.
        No validation is performed.
        """
        s = v / 255.0
        # Apply sRGB companding inverse (linear + gamma segment)
        return s / 12.92 if s <= 0.04045 else ((s + 0.055) / 1.055) ** 2.4

    # Unpack RGB tuple
    r, g, b = rgb

    # Convert each channel to linear RGB
    r_lin, g_lin, b_lin = to_linear(r), to_linear(g), to_linear(b)

    # Weighted luminance according to WCAG definition
    return 0.2126 * r_lin + 0.7152 * g_lin + 0.0722 * b_lin

def contrast_ratio_rgb(fg: tuple[int, int, int], bg: tuple[int, int, int]) -> float:
    """
    Compute the WCAG contrast ratio between two RGB colors.

    Parameters
    ----------
    fg : tuple[int, int, int]
        Foreground color (0–255 per channel).
    bg : tuple[int, int, int]
        Background color (0–255 per channel).

    Returns
    -------
    float
        Contrast ratio in the range 1.0–21.0.

    Notes
    -----
    Uses WCAG formula: (L_lighter + 0.05) / (L_darker + 0.05).
    No input validation or clamping is performed.
    """
    # Compute relative luminance for both colors
    l1 = relative_luminance_from_rgb(fg)
    l2 = relative_luminance_from_rgb(bg)

    # Identify lighter and darker luminance
    lighter, darker = max(l1, l2), min(l1, l2)

    # Apply WCAG contrast ratio formula
    return (lighter + 0.05) / (darker + 0.05)

def pick_accessible_text_color(
    background_color: str,
    *,
    dark_text_hex: str,
    light_text_hex: str = "#F4F6F6",
) -> str:
    """
    Select the higher-contrast text color for a given background.

    Parameters
    ----------
    background_color : str
        Plotly-compatible color string ('#RRGGBB', 'rgb(...)', or 'rgba(...)').
    dark_text_hex : str
        Candidate dark text color (hex or rgb/rgba string).
    light_text_hex : str, default '#F4F6F6'
        Candidate light text color (theme-aligned near-white).

    Returns
    -------
    str
        The candidate color string with the higher WCAG contrast ratio.

    Notes
    -----
    Alpha values in 'rgba(...)' inputs are ignored.
    No validation or WCAG threshold enforcement is performed.
    """
    # Convert all inputs to integer RGB tuples (alpha ignored)
    bg_rgb = plotly_color_to_rgb_tuple(background_color)
    dark_rgb = plotly_color_to_rgb_tuple(dark_text_hex)
    light_rgb = plotly_color_to_rgb_tuple(light_text_hex)

    # Compute WCAG contrast ratios against background
    cr_dark = contrast_ratio_rgb(dark_rgb, bg_rgb)
    cr_light = contrast_ratio_rgb(light_rgb, bg_rgb)

    # Return the candidate with the higher contrast ratio
    # (ties resolved in favor of light text)
    return light_text_hex if cr_light >= cr_dark else dark_text_hex

def apply_accessible_text_colors(
    sunburst_trace,
    segment_colors: list[str],
    dark_text_hex: str,
    light_text_hex: str,
) -> None:
    """
    Apply contrast-aware text colors to a Plotly Sunburst trace.

    For each segment background color, this helper selects the higher-contrast
    candidate text color using pick_accessible_text_color(...), then writes the
    per-segment colors to inside/outside text font settings.

    Parameters
    ----------
    sunburst_trace
        Plotly Sunburst trace (typically fig.data[0]).
    segment_colors : list[str]
        Segment background colors in trace order (hex or rgb/rgba strings).
    dark_text_hex : str
        Candidate dark text color.
    light_text_hex : str
        Candidate light text color.

    Returns
    -------
    None
        Mutates the trace in place.
    """
    text_colors = [
        pick_accessible_text_color(
            bg,
            dark_text_hex=dark_text_hex,
            light_text_hex=light_text_hex,
        )
        for bg in segment_colors
    ]
    sunburst_trace.update(
        insidetextfont=dict(color=text_colors),
        outsidetextfont=dict(color=text_colors),
    )

# ---------------------------------------------------------------------------
# KWIC formatting and highlighting utilities
# ---------------------------------------------------------------------------

def format_kwic(context: str, variants: list[str]) -> tuple[str, str, str]:
    """
    Split a context string into KWIC (Key Word in Context) segments.

    The function searches for the first occurrence of any provided variant
    (case-insensitive) within the given context string and returns a
    three-part tuple:

        (left_context, matched_token, right_context)

    Matching behavior
    ------------------
    - Matching is case-insensitive.
    - Only the first matching variant (in iteration order) is considered.
    - The returned "hit" preserves the original casing from `context`.
    - Surrounding context is stripped of leading/trailing whitespace.
    - No token-boundary or regex-based matching is performed
      (simple substring search via str.find).

    Parameters
    ----------
    context : str
        Full text string in which a keyword should be located
        (e.g., a collocation or verse excerpt).
    variants : list[str]
        List of candidate keyword variants to search for.
        Variants are matched in the given order.

    Returns
    -------
    tuple[str, str, str]
        (left, hit, right)

        - left  : substring before the matched variant (stripped)
        - hit   : matched substring (original casing preserved)
        - right : substring after the matched variant (stripped)

        If no variant matches, returns:
            (context.strip(), "", "")
    """
    # Lowercase copy used exclusively for case-insensitive search.
    # The original `context` string remains unchanged for slicing.
    context_lower = context.lower()

    # Iterate over variants in the provided order.
    # The first successful match determines the KWIC split.
    for variant in variants:
        # Case-insensitive substring search.
        index = context_lower.find(variant.lower())

        # If a match is found (index >= 0), compute KWIC segments.
        if index != -1:
            # Left context: everything before the match.
            left = context[:index].strip()

            # Hit: preserve original casing from `context`.
            hit = context[index:index + len(variant)]

            # Right context: everything after the match.
            right = context[index + len(variant):].strip()

            return left, hit, right

    # No variant matched → return full context as left segment,
    # and empty hit/right placeholders.
    return context.strip(), "", ""

# ---------------------------------------------------------------------------
# Token extraction utilities
# ---------------------------------------------------------------------------

def extract_tokens(entries: list[dict], unit: str) -> list[str]:
    """
    Extract lemma-level tokens from categorized naming entries.

    This helper collects values from structured categorization dictionaries
    that contain numbered lemma columns for naming variants and epithets.

    Column conventions (assumed structure)
    --------------------------------------
    - Naming variants:  "Bezeichnung 1" ... "Bezeichnung 4"
    - Epithets:         "Epitheta 1" ... "Epitheta 5"

    Extraction behavior
    -------------------
    - Only string values are considered.
    - Empty or whitespace-only strings are ignored.
    - Leading/trailing whitespace is stripped.
    - Order of tokens follows input order:
        entries order → column order (1..n).
    - No de-duplication is performed.
    - No normalization is applied (caller decides whether to normalize).

    Parameters
    ----------
    entries : list[dict]
        List of categorization entries (e.g., JSON-derived rows),
        each represented as a dictionary.
    unit : str
        Token category selector:
            - "bezeichnung" → only naming variants
            - "epitheta"    → only epithets
            - "combined"    → both naming variants and epithets

        Any other value results in an empty list.

    Returns
    -------
    list[str]
        Flat list of extracted token strings.
        May contain duplicates if present across entries.
    """
    tokens = []

    for entry in entries:
        # Extract naming variants if requested.
        if unit in ("bezeichnung", "combined"):
            # Fixed column range (1–4) according to project schema.
            for i in range(1, 5):
                val = entry.get(f"Bezeichnung {i}")

                # Accept only non-empty string values.
                if isinstance(val, str) and val.strip():
                    tokens.append(val.strip())

        # Extract epithets if requested.
        if unit in ("epitheta", "combined"):
            # Fixed column range (1–5) according to project schema.
            for i in range(1, 6):
                val = entry.get(f"Epitheta {i}")

                # Accept only non-empty string values.
                if isinstance(val, str) and val.strip():
                    tokens.append(val.strip())

    return tokens

def collect_tokens_for_cooccurrence(row: dict, include_naming_variants: bool, include_epithets: bool) -> list[str]:
    """
    Collect lemma tokens from a single categorization entry for co-occurrence analysis.

    This helper extracts tokens from the fixed-schema lemma columns of one
    entry (row) and prepares them for intra-entry co-occurrence computation.

    Column conventions (assumed schema)
    -----------------------------------
    - Naming variants:  "Bezeichnung 1" ... "Bezeichnung 4"
    - Epithets:         "Epitheta 1" ... "Epitheta 5"

    Extraction behavior
    -------------------
    - Only string values are considered.
    - Empty or whitespace-only values are ignored.
    - Leading/trailing whitespace is stripped.
    - Tokens are de-duplicated within the entry.
    - Output is alphabetically sorted to ensure deterministic order
      (important for stable co-occurrence pairing).

    Parameters
    ----------
    row : dict
        Single categorization entry (e.g., one JSON-derived record).
    include_naming_variants : bool
        If True, extract tokens from "Bezeichnung 1–4".
    include_epithets : bool
        If True, extract tokens from "Epitheta 1–5".

    Returns
    -------
    list[str]
        Sorted list of unique tokens extracted from the row.
        Returns an empty list if no valid tokens are found or both
        include-flags are False.
    """
    tokens: list[str] = []

    # Extract naming variants if requested.
    if include_naming_variants:
        # Fixed project schema: up to four Bezeichnung columns.
        for i in range(1, 5):
            v = row.get(f"Bezeichnung {i}", "")

            # Accept only non-empty string values.
            if isinstance(v, str) and v.strip():
                tokens.append(v.strip())

    # Extract epithets if requested.
    if include_epithets:
        # Fixed project schema: up to five Epitheta columns.
        for i in range(1, 6):
            v = row.get(f"Epitheta {i}", "")

            # Accept only non-empty string values.
            if isinstance(v, str) and v.strip():
                tokens.append(v.strip())

    # Remove duplicates within this entry and sort alphabetically
    # to ensure deterministic and reproducible co-occurrence pairing.
    return sorted(set(tokens))

# ---------------------------------------------------------------------------
# Lemma resolution and heuristic matching utilities
# ---------------------------------------------------------------------------

def match_name_to_lemma(target_figure, lemma, aliases=None):
    """
    Determine whether a lemma qualifies as a proper-name mention of a target figure.

    This helper performs a lightweight, case-insensitive string comparison
    between a canonical target figure name and a candidate lemma.

    Matching behavior
    ------------------
    - Comparison is case-insensitive.
    - Leading/trailing whitespace is ignored.
    - A match is accepted if:
        (1) the normalized lemma equals the normalized target name, or
        (2) the normalized lemma equals any normalized alias (if provided).
    - No advanced normalization (e.g., diacritic handling, fuzzy matching,
      edit distance) is performed.
    - Non-string or empty lemma values are rejected.

    Parameters
    ----------
    target_figure : str
        Canonical form of the target figure name.
        Assumed to be a non-empty string.
    lemma : str
        Lemma/designation candidate to evaluate.
    aliases : list[str] | None, optional
        Optional list of alternative spellings or known name variants.
        Comparison follows the same normalization logic as for the target.

    Returns
    -------
    bool
        True if the lemma counts as a name-based mention of the target figure,
        False otherwise.

    Notes
    -----
    - This function is intentionally conservative and deterministic.
    - It is designed for controlled naming-analysis contexts, not for
      general fuzzy name resolution.
    """
    # Reject non-string or empty lemma candidates early.
    if not isinstance(lemma, str) or lemma.strip() == "":
        return False

    # Normalize both target and lemma for case-insensitive comparison.
    # Only lowercase + strip; no diacritic or grapheme normalization here.
    norm_target = target_figure.lower().strip()
    norm_lemma  = lemma.lower().strip()

    # Direct (exact) match.
    if norm_lemma == norm_target:
        return True

    # Alias-based match (if alias list provided).
    if aliases:
        for a in aliases:
            # Only consider valid string aliases.
            if isinstance(a, str) and a.lower().strip() == norm_lemma:
                return True

    # No match found.
    return False

def resolve_name_lemmas_for_figure(df, cols, figure_name):
    """
    Resolve which lemmas should be treated as the proper name (Eigenname) of a figure
    for a single visualization run.

    This helper scans the naming data for a given `figure_name` and determines which
    lemma strings (from Bezeichnung* and optionally Epitheta*) should be interpreted
    as proper-name mentions of that figure. The result is used only to label/aggregate
    items as "Eigenname" in the current visualization context; it does not modify
    the underlying data and is not persisted.

    Decision procedure
    ------------------
    1) Filter rows where the target figure column equals `figure_name` (string-stripped).
    2) Collect all lemma strings from:
       - naming-variant columns (Bezeichnung*, excluding an unnumbered "Bezeichnung" column)
       - epithet columns (Epitheta*)
    3) Attempt direct name matching using `match_name_to_lemma(...)` (case-insensitive,
       whitespace-trimmed exact match; aliases are currently not used here).
       - If at least one lemma matches directly, return the set of matched lemmas.
    4) If no direct matches exist but lemma candidates were collected:
       - Use `difflib.get_close_matches` to suggest the closest lemma candidate
         to `figure_name` (string similarity heuristic).
       - Prompt the user (CLI) to confirm whether the suggestion should be treated
         as a proper-name variant for this visualization run.
       - If confirmed, return a singleton set containing the suggested lemma.
    5) Otherwise return an empty set.

    Parameters
    ----------
    df : pandas.DataFrame
        Naming data (typically trimmed to relevant columns). Must contain the column
        referenced by `cols["target"]` and may contain columns listed in
        `cols["naming_variant_cols"]` and `cols["epithet_cols"]`.
    cols : dict
        Canonical column mapping, expected keys include:
        - "target" (str): column name for the named figure (Benannte Figur)
        - "naming_variant_cols" (list[str]): Bezeichnung* columns (may include "Bezeichnung")
        - "epithet_cols" (list[str]): Epitheta* columns
    figure_name : str
        Canonical figure name to resolve against lemma candidates.

    Returns
    -------
    set[str]
        Set of lemma strings that should be treated as proper-name mentions of
        `figure_name` for this visualization run. Empty set if no decision can be made
        (missing target column, no lemma candidates, no similarity suggestion, or user rejects).

    Notes
    -----
    - This function is intentionally interactive (CLI prompt) in the fallback path.
    - It uses a conservative direct-match first, then a similarity-based suggestion.
    - It does not apply project-wide text normalization beyond `.lower().strip()`.
      (If diacritic/orthography normalization is required, it should be handled upstream.)
    """
    # Column names are provided via `cols` to keep the function independent of
    # the concrete DataFrame schema (JSON vs. Excel derived).
    target_col = cols.get("target")

    # Naming-variant columns may include an unnumbered base column "Bezeichnung".
    # For visualization aggregation, we exclude that base column here to focus on
    # the numbered lemma slots (avoids ambiguous/legacy schema artifacts).
    naming_variant_cols_all = cols.get("naming_variant_cols", [])
    naming_variant_cols = [
        c for c in naming_variant_cols_all
        if str(c).strip().lower() != "bezeichnung"
    ]

    # Epithets are optional; if present, they can contribute candidate lemmas.
    epithet_cols = cols.get("epithet_cols", [])

    # Without a target column, there ist no filter to the figure
    # and anything cannot be resolved.
    if not target_col:
        return set()

    # Restrict to rows where the target figure equals `figure_name` (string-based match).
    # This is a narrow filter: it does not perform fuzzy matching for the target column.
    dff = df[df[target_col].astype(str).str.strip() == str(figure_name).strip()].copy()
    dff = dff.reset_index(drop=True)

    # Collect all lemma candidates encountered for the figure, and those that match
    # the figure name as a proper name under the direct-match heuristic.
    lemmas_all = set()
    lemmas_matched_as_proper_name = set()

    # 1) Direct name matching over lemma candidates in the figure-restricted subset.
    for _, row in dff.iterrows():
        # Naming variant (Bezeichnung* columns).
        for col in naming_variant_cols:
            val = row.get(col)
            if not isinstance(val, str):
                continue

            lemma = val.strip()
            if not lemma:
                continue

            # Track all observed lemma candidates for later fallback suggestion.
            lemmas_all.add(lemma)

            # Direct-match heuristic: case-insensitive exact match (aliases not used here).
            try:
                if match_name_to_lemma(figure_name, lemma, aliases=None):
                    lemmas_matched_as_proper_name.add(lemma)
            except (TypeError, ValueError, AttributeError):
                # Defensive: ignore unexpected type/attribute issues in matching.
                pass

        # Epithets (Epitheta* columns). These are optional and treated similarly.
        for col in epithet_cols:
            val = row.get(col)
            if not isinstance(val, str):
                continue

            lemma = val.strip()
            if not lemma:
                continue

            lemmas_all.add(lemma)

            try:
                if match_name_to_lemma(figure_name, lemma, aliases=None):
                    lemmas_matched_as_proper_name.add(lemma)
            except (TypeError, ValueError, AttributeError):
                pass

    # 2) If any direct name matches exist, return them without invoking the interactive fallback.
    if lemmas_matched_as_proper_name:
        return lemmas_matched_as_proper_name

    # 3) Fallback: if no lemma candidates exist, nothing can be suggested.
    if not lemmas_all:
        return set()

    # Build a case-insensitive lookup for difflib suggestions while preserving original casing.
    lowered = {lm.lower(): lm for lm in lemmas_all}

    # Similarity heuristic: propose the closest lemma candidate to `figure_name`.
    # cutoff=0.6 is a pragmatic threshold; it may be tuned later if needed.
    best = difflib.get_close_matches(
        figure_name.lower(),
        list(lowered.keys()),
        n=1,
        cutoff=0.6
    )

    # If difflib cannot suggest anything above the threshold, return empty set.
    if not best:
        return set()

    suggested = lowered[best[0]]

    # Interactive confirmation: only applied in the absence of direct matches.
    print(f'{figure_name} could not be matched to any lemma as a proper name.')
    yn = ask_user_choice(
        f'Could "{suggested}" be a variant of the proper name? (y/n)',
        ["y", "n"]
    )

    if yn == "y":
        return {suggested}

    return set()

# ---------------------------------------------------------------------------
# Data discovery / filesystem utilities
# ---------------------------------------------------------------------------

def list_available_reference_books(*, data_dir: str | Path = "data", exclude: str | None = None) -> list[str]:
    """
    Discover available reference books based on the expected folder structure.

    The function scans a base data directory and derives book identifiers
    from subdirectories that contain a categorization JSON file following
    the naming convention:

        data/<book_name>/categorization_<book_name>.json

    Only directories matching this structure are considered valid books.

    Parameters
    ----------
    data_dir : str | Path, optional
        Base directory to scan (default: "data").
        Must contain subdirectories named after book identifiers.
    exclude : str | None, optional
        Optional book identifier to exclude from the result list
        (e.g., the currently analyzed book).

    Returns
    -------
    list[str]
        Alphabetically sorted list (case-insensitive) of discovered
        book identifiers.

        Returns an empty list if:
        - the base directory does not exist,
        - it is not a directory,
        - or no valid categorization files are found.

    Notes
    -----
    - No file content validation is performed.
    - The function only checks for file existence and naming conformity.
    - Sorting is case-insensitive to ensure deterministic ordering.
    """
    # Normalize base path and verify existence.
    base = Path(data_dir)
    if not base.exists() or not base.is_dir():
        return []

    books: list[str] = []

    # Iterate over immediate subdirectories in the base directory.
    for p in base.iterdir():
        # Only consider directories.
        if not p.is_dir():
            continue

        book = p.name

        # Optionally exclude a specific book identifier.
        if exclude and book == exclude:
            continue

        # Check for expected categorization file naming pattern.
        cat_path = p / f"categorization_{book}.json"

        # Accept only if the file exists and is a regular file.
        if cat_path.exists() and cat_path.is_file():
            books.append(book)

    # Return case-insensitive sorted list for stable CLI presentation.
    return sorted(books, key=lambda s: s.lower())
