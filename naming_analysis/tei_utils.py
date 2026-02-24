"""
tei_utils.py

Utilities for TEI-XML handling in the naming-analysis pipeline.

Scope:
- Provide the shared TEI namespace mapping (`tei_ns`) used across XML queries.
- Normalize textual content of TEI <seg> elements for downstream matching.
- Parse and standardize verse identifiers (e.g., TEI <l @n="n"> and Excel "Vers").
- Retrieve a local verse context window around a given verse number.

Notes:
- `normalize_tei_text(...)` mutates the passed TEI root element in-place and prints a short status line.
- `get_verse_context(...)` returns only verses that exist in the TEI tree; numbering is 1-based for display.
"""
from naming_analysis.shared import normalize_text, parse_verse_number

# Global TEI namespace mapping (project-wide standard for XPath queries)
# Official TEI-P5 namespace (must remain HTTP, not HTTPS — defined by TEI standard)
tei_ns = {'tei': 'http://www.tei-c.org/ns/1.0'}

def get_valid_verse_number(value, fallback=-1):
    """
    Normalize and validate a verse identifier as float.

    Thin wrapper around `parse_verse_number(...)` to enforce a consistent
    verse-number interpretation across the pipeline (TEI @n attributes,
    Excel "Vers" columns, JSON imports).

    Accepted input formats:
    - strings using comma or dot as decimal separator (e.g. "15,2", "15.2")
    - integers or floats (e.g. 17, 18.0)

    Normalization rules:
    - Decimal separators are unified.
    - Return type is always float if parsing succeeds.

    Failure behavior:
    - If parsing fails, the provided `fallback` value is returned unchanged.
    - No exception is raised.

    Parameters:
        value (any): Raw verse identifier.
        fallback (float | int): Value returned on parsing failure (default: -1).

    Returns:
        float | int: Parsed float on success, otherwise the fallback value.
    """
    return parse_verse_number(value, fallback=fallback)

def get_verse_context(verse_number, root_tei):
    """
    Return a normalized local verse context window (±6) around a given TEI verse number.

    The function attempts to fetch verses `verse_number - 6` through `verse_number + 6`
    via TEI lines `<l @n="n">`. Matching is based on exact string equality of the
    `@n` attribute. Missing verses are skipped, so the returned list may
    contain fewer than 13 entries.

    Text extraction is limited to descendant `<seg>` elements; their texts are joined
    with spaces and passed through `normalize_text(...)`.

    The returned numbering (1nnN) is a local index for the context window and does not
    correspond to the original TEI verse numbers.

    Parameters:
        verse_number (int): Central TEI verse number used to build the context window.
        root_tei (Element): Parsed TEI root element.

    Returns:
        list[tuple[int, str]]:
            A list of (local_index, normalized_text) tuples.

            - local_index starts at 1 and increases consecutively.
            - Only verses that exist in the TEI tree are included.
            - The index reflects the position within the returned window,
              not the original TEI verse number.
    """
    # Stores final (local_index, normalized_text) tuples
    context = []
    # Temporary list of extracted verse texts (order preserved)
    verse_list = []

    # Build a ±6 verse window around the central verse (inclusive)
    for i in range(-6, 7):
        verse_id = str(verse_number + i)
        # Query TEI <l> elements by exact @n attribute using the project namespace
        line = root_tei.find(f'.//tei:l[@n="{verse_id}"]', tei_ns)

        if line is not None:
            # Extract text only from descendant <seg> elements,
            # join them with spaces, then normalize the result
            text = normalize_text(' '.join([
                seg.text for seg in line.findall('.//tei:seg', tei_ns) if seg.text
            ]))
            verse_list.append(text)

    # Assign consecutive local indices starting at 1
    for i, verse in enumerate(verse_list, start=1):
        context.append((i, verse))

    return context

def normalize_tei_text(root):
    """
    Normalize the text content of all TEI <seg> elements in-place.

    The function iterates over all descendant <seg> elements (TEI namespace)
    and replaces their `.text` value with the result of `normalize_text(...)`.

    Important:
    - The passed XML tree is modified in-place.
    - If `root` is None, the function returns None without raising an exception.
    - A short status message is printed to stdout.

    Parameters:
        root (Element | None): Root element of the parsed TEI XML tree.

    Returns:
        Element | None:
            The identical root element (mutated in-place).
            No new XML tree is created.
            Returns None only if the input was None.
    """
    # Exit early if root is None
    if root is None:
        return None

    # Iterate over all descendant <seg> elements using the TEI namespace
    for seg in root.findall('.//tei:seg', tei_ns):
        if seg.text:
            # Replace original text with normalized version (in-place mutation)
            seg.text = normalize_text(seg.text)

    # Informational CLI status output
    print("TEI text has been normalized.")

    # Return the same (mutated) root element for API consistency
    return root