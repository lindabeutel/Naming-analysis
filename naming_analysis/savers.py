"""
savers.py

This module handles all write operations for persistent data during
the naming analysis process. It includes:

- Saving of progress and checkpoints (verses, variants, categories)
- Export of normalization and categorization data
- Writing configuration and annotation files

All JSON write operations rely on the unified `safe_write_json()` utility
from `io_utils.py`.
"""
from naming_analysis.io_utils import safe_write_json, safe_read_json
from naming_analysis.shared import get_first_valid_text
from naming_analysis.tei_utils import get_valid_verse_number
from copy import deepcopy

def save_progress(
    missing_naming_variants,
    last_processed_verse,
    paths,
    previous_verse=None,
    previous_naming_variants=None,
    collocation_data=None,
    previous_collocations=None,
    categorized_entries=None,
    previous_categorized_entries=None,
    check_naming_variants=False,
    perform_collocations=False,
    perform_categorization=False
):

    """
    Saves the current progress of the data collection process.

    Depending on the enabled flags, the function updates the last processed
    verse number for naming variants, collocations, or categorizations, and
    persists the corresponding data only if changes have occurred.

    Parameters:
        missing_naming_variants (list): New or updated list of naming variants.
        last_processed_verse (int): The most recently processed verse number.
        paths (dict): Dictionary of file paths used for saving.
        previous_verse (int, optional): The last saved verse number (for comparison).
        previous_naming_variants (list, optional): Previously saved variant list.
        collocation_data (list, optional): Updated collocation entries.
        previous_collocations (list, optional): Previously saved collocations.
        categorized_entries (list, optional): Newly categorized entries.
        previous_categorized_entries (list, optional): Previously saved categorized entries.
        check_naming_variants (bool): Whether naming variants should be saved.
        perform_collocations (bool): Whether collocations should be saved.
        perform_categorization (bool): Whether categorizations should be saved.
    """
    # Load existing progress file (if available)
    progress_data = safe_read_json(paths["progress_json"], default={})

    # Update the respective last-verse value only if it changed
    if previous_verse is None or last_processed_verse != previous_verse:
        if check_naming_variants:
            progress_data["naming_variants_last_verse"] = last_processed_verse
        if perform_collocations:
            progress_data["collocations_last_verse"] = last_processed_verse
        if perform_categorization:
            progress_data["categorization_last_verse"] = last_processed_verse

        safe_write_json(progress_data, paths["progress_json"])

    if previous_naming_variants is None or sorted_entries(missing_naming_variants) != sorted_entries(previous_naming_variants):
        safe_write_json(missing_naming_variants, paths["missing_naming_variants_json"], merge=True)

    if collocation_data is not None:
        if previous_collocations is None or collocation_data != previous_collocations:
            safe_write_json(collocation_data, paths["collocations_json"], merge=True)

    if categorized_entries is not None:
        if previous_categorized_entries is None or sorted_entries(categorized_entries) != sorted_entries(
                previous_categorized_entries):
            safe_write_json(categorized_entries, paths["categorization_json"], merge=True)

def save_lemma_normalization(data, path="lemma_normalization.json"):
    """
    Writes lemma normalization rules to a JSON file.

    The data is sorted alphabetically by lemma, and all variant lists
    are de-duplicated and sorted internally.

    Parameters:
        data (dict): Mapping of lemmas to variant lists.
        path (str): Destination file path. Defaults to 'lemma_normalization.json'.
    """
    sorted_data = {
        lemma: sorted(set(variants))
        for lemma, variants in sorted(data.items(), key=lambda x: x[0].lower())
    }
    safe_write_json(sorted_data, path, merge=False)

def save_ignored_lemmas(data, path="ignored_lemmas.json"):
    """
    Saves the list of ignored lemmas to a JSON file.

    The list is sorted alphabetically and merged with existing content if present.

    Parameters:
        data (set | list): Lemmas to be ignored in future processing.
        path (str): Destination file path. Defaults to 'ignored_lemmas.json'.
    """
    safe_write_json(data, path, sort_keys=True, merge=True)

def save_lemma_categories(data, path="data/lemma_categories.json"):
    """
    Writes categorized lemma labels (e.g. 'a' or 'e') to a JSON file.

    Existing entries are preserved and updated; the resulting dictionary
    is sorted alphabetically by key.

    Parameters:
        data (dict): Mapping from lemma to category label.
        path (str): Destination file path. Defaults to 'data/lemma_categories.json'.
    """
    existing = safe_read_json(path, default={})
    existing.update(data)
    sorted_data = dict(sorted(existing.items()))
    safe_write_json(sorted_data, path, merge=False)

def save_json_annotations(path, annotations):
    """
    Writes annotated entries to a JSON file using a merge strategy.

    Parameters:
        path (str): Target file path.
        annotations (list): List of new annotations to be written.
    """
    safe_write_json(annotations, path, merge=True)

def sorted_entries(entries: list) -> list:
    """
    Returns a cleaned and consistently sorted list of entry dictionaries.

    Entries are:
    - filtered to include only those with a valid integer 'Vers' value
    - sorted by:
        (1) verse number (numerical),
        (2) the first non-empty string among 'Eigennennung', 'Bezeichnung', or 'Erzähler' (case-insensitive)

    Used to ensure consistent ordering when comparing or saving naming variants or categorized entries.

    Parameters:
        entries (list): A list of dictionaries representing naming or categorization entries.

    Returns:
        list: The cleaned and sorted list of entries.
    """
    # Only keep entries with a valid numeric "Vers" value
    entries_clean = [
        e for e in deepcopy(entries)
        if isinstance(e, dict)
        and str(e.get("Vers", "")).strip().isdigit()
    ]

    return sorted(
        entries_clean,
        key=lambda x: (
            get_valid_verse_number(x.get("Vers")),
            get_first_valid_text(
                x.get("Eigennennung"),
                x.get("Bezeichnung"),
                x.get("Erzähler")
            ).strip().lower()
        )
    )