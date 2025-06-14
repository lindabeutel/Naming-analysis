"""
tei_utils.py

Utility functions for working with TEI-XML documents.
Includes normalization of <seg> elements and retrieval of verse context.
"""

from naming_analysis.shared import normalize_text

tei_ns = {'tei': 'http://www.tei-c.org/ns/1.0'}

def get_valid_verse_number(value, fallback=-1):
    """
    Tries to parse a verse number as integer.
    Returns fallback if conversion fails (e.g., empty, NaN, wrong type).

    Parameters:
        value (any): Input value to parse.
        fallback (int): Value to return on failure.

    Returns:
        int: Validated verse number or fallback.
    """
    try:
        return int(str(value).strip())
    except (ValueError, TypeError):
        return fallback

def normalize_tei_text(root):
    """
    Normalizes the text content of all <seg> elements in a TEI document.

    Parameters:
        root (Element): Root of the TEI XML tree.

    Returns:
        Element: Modified TEI root with normalized text.
    """
    if root is None:
        return None

    for seg in root.findall('.//tei:seg', tei_ns):
        if seg.text:
            seg.text = normalize_text(seg.text)

    print("✓ TEI text has been normalized.")
    return root

def get_verse_context(verse_number, root_tei):
    """
    Retrieves the surrounding 6 verses from the TEI file, numbered 1–13.

    Parameters:
        verse_number (int): Central verse number.
        root_tei (Element): Parsed TEI root element.

    Returns:
        list[tuple[int, str]]: List of numbered context lines.
    """
    context = []
    verse_list = []

    for i in range(-6, 7):
        verse_id = str(verse_number + i)
        line = root_tei.find(f'.//tei:l[@n="{verse_id}"]', tei_ns)

        if line is not None:
            text = normalize_text(' '.join([
                seg.text for seg in line.findall('.//tei:seg', tei_ns) if seg.text
            ]))
            verse_list.append(text)

    for i, verse in enumerate(verse_list, start=1):
        context.append((i, verse))

    return context
