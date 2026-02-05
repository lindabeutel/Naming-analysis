"""
shared.py

General-purpose helper functions used throughout the project.
Includes text normalization, fallback selection, user interaction, and data cleaning utilities."""

import math
import re
import pandas as pd

from copy import deepcopy
from plotly.colors import hex_to_rgb, label_rgb

_RGBA_RE = re.compile(
    r"^rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})(?:\s*,\s*([01](?:\.\d+)?))?\s*\)$",
    re.IGNORECASE,
)

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

def prepare_naming_data(book_name, df_json, df_excel):
    """
    Pick a source (prefer JSON, fallback to Excel) and detect relevant columns
    for the 'Naming figure analysis'.

    Returns a tuple (source, df_trimmed, cols):
      - source: "json" or "excel"
      - df_trimmed: DataFrame with only relevant columns
      - cols: dictionary with
          {
            "target": <column name for 'Benannte Figur'>,
            "namer": <column name for 'Nennende Figur'>,
            "designation_cols": [...],       # Bezeichnung*, includes unnumbered 'Bezeichnung' if present
            "epithet_cols": [...],           # Epitheta*
            "verse_col": <column name or None>,
            "has_unnumbered_designation": True/False
          }

    Raises:
      ValueError if neither JSON nor Excel satisfies the minimum requirements.
    """

    def normalize(colname):
        """Normalize column names: lowercase, strip, collapse spaces."""
        return re.sub(r"\s+", " ", colname.strip().lower()) if isinstance(colname, str) else ""

    def detect_columns(df):
        """Detect mandatory and dynamic columns using tolerant matching rules."""
        nmap = {c: normalize(c) for c in df.columns}

        def match(regex):
            rx = re.compile(regex, re.IGNORECASE)
            return [c for c in df.columns if rx.fullmatch(c) or rx.fullmatch(nmap[c])]

        # mandatory columns
        target = None
        namer = None
        for c in df.columns:
            normed = nmap[c]
            if normed in ("benannte figur", "benannte_figur"):
                target = c
            if normed in ("nennende figur", "nennende_figur"):
                namer = c

        # dynamic lemma columns
        designation_cols = match(r"bezeichnung(\s*\d+)?")
        epithet_cols     = match(r"epitheta(\s*\d+)?")
        verse_cols       = match(r"vers(|-?id)?")

        has_unnumbered   = any(normalize(c) == "bezeichnung" for c in designation_cols)
        has_lex          = (len(designation_cols) + len(epithet_cols)) > 0

        cols = {
            "target": target,
            "namer": namer,
            "designation_cols": designation_cols,
            "epithet_cols": epithet_cols,
            "verse_col": verse_cols[0] if verse_cols else None,
            "has_unnumbered_designation": has_unnumbered,
            "has_lex": has_lex,
        }
        return cols

    def trim(df, cols):
        """Keep only the relevant columns (remove duplicates, preserve order)."""
        keep = []
        if cols["target"]:
            keep.append(cols["target"])
        if cols["namer"]:
            keep.append(cols["namer"])
        keep += cols["designation_cols"]
        keep += cols["epithet_cols"]
        if cols["verse_col"]:
            keep.append(cols["verse_col"])
        # deduplicate but preserve order
        keep = [c for c in dict.fromkeys(keep) if c in df.columns]
        return df.loc[:, keep].copy()

    # 1) try JSON first
    if df_json is None or df_json.empty:
        print("⚠️ No JSON data provided or JSON is empty – loading from Excel instead.")
    else:
        cj = detect_columns(df_json)

        # structural check: columns present?
        json_ok = cj["target"] and cj["namer"] and cj["has_lex"]
        if not json_ok:
            missing_parts = []
            if not cj["target"]:
                missing_parts.append("'Benannte Figur'")
            if not cj["namer"]:
                missing_parts.append("'Nennende Figur'")
            if not cj["has_lex"]:
                missing_parts.append("Bezeichnung*/Epitheta*")
            print(f"⚠️ JSON missing {', '.join(missing_parts)} – loading from Excel instead.")
        else:
            # content-level check: 'Nennende Figur' required only if a Bezeichnung/Epitheton is present,
            # not if the row has an Eigennennung or Erzähler instead
            lemma_cols = [c for c in (cj["designation_cols"] + cj["epithet_cols"]) if c in df_json.columns]
            eigennennung_col = "Eigennennung" if "Eigennennung" in df_json.columns else None
            erzaehler_col = "Erzähler" if "Erzähler" in df_json.columns else None

            if lemma_cols:
                def _nonempty_series(s):
                    return s.dropna().astype(str).str.strip().ne("")

                has_lemma = df_json[lemma_cols].apply(_nonempty_series).any(axis=1)

                has_eigennennung = (
                    _nonempty_series(df_json[eigennennung_col]) if eigennennung_col else pd.Series(False,
                                                                                                   index=df_json.index)
                )
                has_erzaehler = (
                    _nonempty_series(df_json[erzaehler_col]) if erzaehler_col else pd.Series(False, index=df_json.index)
                )

                namer_col = cj["namer"]
                namer_nonempty = df_json[namer_col].fillna("").astype(str).str.strip().ne("")

                bad_rows = has_lemma & ~has_eigennennung & ~has_erzaehler & ~namer_nonempty
                bad_count = int(bad_rows.sum())

                if bad_count > 0:
                    print(
                        f"⚠️ JSON has {bad_count} row(s) with Bezeichnung/Epitheton but empty 'Nennende Figur' — loading Excel instead.")
                else:
                    # JSON passes both structural and content checks → accept JSON
                    return "json", trim(df_json, cj), {
                        "target": cj["target"],
                        "namer": cj["namer"],
                        "designation_cols": cj["designation_cols"],
                        "epithet_cols": cj["epithet_cols"],
                        "verse_col": cj["verse_col"],
                        "has_unnumbered_designation": cj["has_unnumbered_designation"],
                    }

    # 2) fallback to Excel
    if df_excel is not None and not df_excel.empty:
        cx = detect_columns(df_excel)
        x_ok = cx["target"] and cx["namer"] and cx["has_lex"]
        if x_ok:
            print("ℹ️ Using Excel fallback (JSON lacked required structure).")
            return "excel", trim(df_excel, cx), {
                "target": cx["target"],
                "namer": cx["namer"],
                "designation_cols": cx["designation_cols"],
                "epithet_cols": cx["epithet_cols"],
                "verse_col": cx["verse_col"],
                "has_unnumbered_designation": cx["has_unnumbered_designation"],
            }

    # 3) no valid data found
    raise ValueError(
        f"[prepare_naming_data] Missing required columns in JSON/Excel for '{book_name}'. "
        "Expected: 'Benannte Figur', 'Nennende Figur', and at least one of Bezeichnung*/Epitheta*."
    )

def serialize_verse_value(value) -> str:
    """
    Serialize numeric or textual verse information for export
    without altering the original formatting semantics.

    - Keeps Excel-style comma decimals as-is (e.g. '17,2')
    - Keeps integers as '20' (not '20.0')
    - Converts JSON floats (17.20) to '17.2'
    - Returns an empty string for None or NaN
    """
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (NameError, AttributeError):
        pass

    # preserve source formatting
    if isinstance(value, str):
        val = value.strip()
        # keep comma decimals (Excel format)
        if "," in val:
            return val
        # trim redundant zeros from dot-decimals
        if "." in val:
            val = val.rstrip("0").rstrip(".")
        return val

    # handle numeric values from JSON (int/float)
    try:
        f = float(value)
        if f.is_integer():
            return str(int(f))
        s = str(f)
        if "." in s:
            s = s.rstrip("0").rstrip(".")
        return s
    except (ValueError, TypeError):
        return str(value)

def hex_color_to_rgb_tuple(hex_color: str) -> tuple[int, int, int]:
    """
    Convert a hexadecimal color code to an RGB tuple.

    Parameters
    ----------
    hex_color : str
        Hexadecimal color code in the form '#RRGGBB'.

    Returns
    -------
    tuple[int, int, int]
        RGB color as a tuple of integers in the range 0–255.

    Notes
    -----
    This utility is format-agnostic and can be reused across different
    visualization backends. It does not encode any visualization semantics.
    """

    r, g, b = hex_to_rgb(hex_color)
    return int(r), int(g), int(b)

def rgb_tuple_to_plotly_color(rgb: tuple[int, int, int]) -> str:
    """
    Convert an RGB tuple to a Plotly-compatible color string.

    Parameters
    ----------
    rgb : tuple[int, int, int]
        RGB color as a tuple of integers in the range 0–255.

    Returns
    -------
    str
        Color string in the form 'rgb(r,g,b)', suitable for Plotly colorscales.

    Notes
    -----
    This function provides a minimal adapter layer between numeric RGB values
    and Plotly's expected color string format. It contains no plot-specific
    logic beyond output formatting.
    """

    # returns 'rgb(r,g,b)' which px.imshow accepts in a scale list
    return label_rgb(rgb)

def plotly_color_to_rgb_tuple(color: str) -> tuple[int, int, int]:
    """
    Convert Plotly color strings (hex, rgb(...), rgba(...)) to an RGB tuple (0–255).
    Accepts '#RRGGBB', 'rgb(r,g,b)', 'rgba(r,g,b,a)'.
    """
    c = str(color).strip()

    if c.startswith("#"):
        return hex_color_to_rgb_tuple(c)

    m = _RGBA_RE.match(c)
    if not m:
        raise ValueError(f"Unsupported color format: {color!r}")

    r = max(0, min(255, int(m.group(1))))
    g = max(0, min(255, int(m.group(2))))
    b = max(0, min(255, int(m.group(3))))
    return r, g, b

def hex_color_to_rgba(hex_color: str, alpha: float) -> str:
    """
    Convert a hex color to an rgba() string with the given alpha value.

    Parameters
    ----------
    hex_color : str
        Hexadecimal color code '#RRGGBB'.
    alpha : float
        Opacity value in the range 0.0–1.0.

    Returns
    -------
    str
        Plotly-compatible rgba() color string.

    Notes
    -----
    This function is purely technical and does not encode visualization
    semantics. It is safe to reuse across different plot types.
    """
    r, g, b = hex_color_to_rgb_tuple(hex_color)
    return f"rgba({r},{g},{b},{alpha})"

def srgb_channel_to_linear(c: float) -> float:
    if c <= 0.04045:
        return c / 12.92
    return ((c + 0.055) / 1.055) ** 2.4

def relative_luminance_from_rgb(rgb: tuple[int, int, int]) -> float:
    """
    WCAG relative luminance from integer RGB (0–255).
    """
    def to_linear(v: int) -> float:
        s = v / 255.0
        return s / 12.92 if s <= 0.04045 else ((s + 0.055) / 1.055) ** 2.4

    r, g, b = rgb
    r_lin, g_lin, b_lin = to_linear(r), to_linear(g), to_linear(b)
    return 0.2126 * r_lin + 0.7152 * g_lin + 0.0722 * b_lin

def contrast_ratio_rgb(fg: tuple[int, int, int], bg: tuple[int, int, int]) -> float:
    """
    WCAG contrast ratio for two RGB tuples.
    """
    l1 = relative_luminance_from_rgb(fg)
    l2 = relative_luminance_from_rgb(bg)
    lighter, darker = max(l1, l2), min(l1, l2)
    return (lighter + 0.05) / (darker + 0.05)

def pick_accessible_text_color(
    background_color: str,
    *,
    dark_text_hex: str,
    light_text_hex: str = "#FFFFFF",
) -> str:
    """
    Pick the better (higher-contrast) text color for a given background.
    Background can be '#RRGGBB', 'rgb(...)', or 'rgba(...)'.
    """
    bg_rgb = plotly_color_to_rgb_tuple(background_color)
    dark_rgb = plotly_color_to_rgb_tuple(dark_text_hex)
    light_rgb = plotly_color_to_rgb_tuple(light_text_hex)

    cr_dark = contrast_ratio_rgb(dark_rgb, bg_rgb)
    cr_light = contrast_ratio_rgb(light_rgb, bg_rgb)

    return light_text_hex if cr_light >= cr_dark else dark_text_hex
