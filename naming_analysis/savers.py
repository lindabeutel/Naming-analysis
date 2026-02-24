"""
savers.py

Persistence layer for the Naming-analysis pipeline.

This module centralizes all write operations to disk, including:

- Saving workflow progress and checkpoints (e.g. verses, variants, categories)
- Exporting normalization and categorization data
- Writing configuration and annotation files

Scope (BETA):
- Functions in this module perform direct file-system side effects.
- No schema validation or structural verification is performed before writing.
- Existing files may be overwritten unless explicitly guarded in the calling context.
- No atomic write guarantees or rollback mechanisms are implemented.

All JSON write operations delegate to the unified `safe_write_json()` utility
defined in `io_utils.py`.
"""
# Local application imports
from naming_analysis.io_utils import safe_write_json, safe_read_json
from naming_analysis.shared import sorted_entries

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
    Persist workflow progress and (optionally) updated data artifacts to disk.

    This function updates the progress JSON (last processed verse) for the enabled
    pipeline steps and writes the corresponding JSON data outputs only if changes
    are detected compared to the provided "previous_*" snapshots.

    Behavior (BETA):
    - Reads the existing progress file via `safe_read_json(..., default={})`.
    - Updates `progress_data` keys depending on the active flags:
      - "naming_variants_last_verse" (if `check_naming_variants` is True)
      - "collocations_last_verse" (if `perform_collocations` is True)
      - "categorization_last_verse" (if `perform_categorization` is True)
    - Writes JSON outputs as side effects via `safe_write_json`.
    - Comparison semantics:
      - naming variants and categorized entries are compared using `sorted_entries(...)`
        to reduce ordering-related diffs.
      - collocations are compared directly (`!=`) without sorting.

    Parameters:
        missing_naming_variants (list): Current list of naming variants to persist.
        last_processed_verse (int): Most recently processed verse number in the current run.
        paths (dict): Path mapping providing JSON targets (e.g. progress, variants, collocations, categorization).
        previous_verse (int, optional): Previously saved last verse value (used for change detection).
        previous_naming_variants (list, optional): Previously saved naming variants (used for change detection).
        collocation_data (list, optional): Current collocation data to persist (if provided).
        previous_collocations (list, optional): Previously saved collocation data (used for change detection).
        categorized_entries (list, optional): Current categorization entries to persist (if provided).
        previous_categorized_entries (list, optional): Previously saved categorization entries (used for change detection).
        check_naming_variants (bool): If True, update naming-variants progress key and persist variants when changed.
        perform_collocations (bool): If True, update collocations progress key and persist collocations when changed.
        perform_categorization (bool): If True, update categorization progress key and persist categorizations when changed.

    Returns:
        None: Writes files only.
    """
    # Load existing progress JSON.
    # If the file does not exist, an empty dict is returned (BETA fallback behavior).
    progress_data = safe_read_json(paths["progress_json"], default={})

    # Only update the stored "last processed verse" if:
    # - no previous value was provided, or
    # - the verse number changed.
    # This prevents unnecessary writes.
    if previous_verse is None or last_processed_verse != previous_verse:

        # Update the progress key depending on which pipeline stage is active.
        # Multiple flags may be True simultaneously.
        if check_naming_variants:
            progress_data["naming_variants_last_verse"] = last_processed_verse

        if perform_collocations:
            progress_data["collocations_last_verse"] = last_processed_verse

        if perform_categorization:
            progress_data["categorization_last_verse"] = last_processed_verse

        # Persist updated progress metadata (overwrites existing progress file).
        safe_write_json(progress_data, paths["progress_json"])

    # Persist naming variants only if:
    # - no previous snapshot exists, or
    # - the content changed (order-insensitive comparison via sorted_entries).
    if previous_naming_variants is None or sorted_entries(missing_naming_variants) != sorted_entries(previous_naming_variants):

        # merge=True indicates that new data is merged with existing JSON content
        # rather than fully replacing it (implementation delegated to safe_write_json).
        safe_write_json(
            missing_naming_variants,
            paths["missing_naming_variants_json"],
            merge=True
        )

    # Collocations are only considered if new data was provided.
    # Comparison is order-sensitive (direct != comparison).
    if collocation_data is not None:
        if previous_collocations is None or collocation_data != previous_collocations:

            safe_write_json(
                collocation_data,
                paths["collocations_json"],
                merge=True
            )

    # Categorized entries use order-insensitive comparison (sorted_entries),
    # mirroring naming-variant behavior.
    if categorized_entries is not None:
        if previous_categorized_entries is None or sorted_entries(categorized_entries) != sorted_entries(
                previous_categorized_entries):

            safe_write_json(
                categorized_entries,
                paths["categorization_json"],
                merge=True
            )

def save_lemma_normalization(data, path="lemma_normalization.json"):
    """
    Persist lemma-normalization rules to a JSON file.

    Behavior:
    - Lemmas are sorted alphabetically (case-insensitive).
    - Each variant list is de-duplicated via `set()` and sorted alphabetically.
    - The resulting structure is written to disk using `safe_write_json`
      with `merge=False` (full overwrite semantics).

    Scope (BETA):
    - No validation of input structure or data types is performed.
    - Assumes `data` is a mapping of lemma (str) → iterable of variant strings.
    - This function performs file-system side effects and does not return a value.

    Parameters:
        data (dict): Mapping of lemma → iterable of variant strings.
        path (str): Destination file path. Defaults to "lemma_normalization.json".

    Returns:
        None: Writes JSON file only.
    """
    # Build a new normalized mapping:
    # - Lemmas are sorted alphabetically (case-insensitive) to ensure stable output.
    # - Each variant list is:
    #     1) de-duplicated via set()
    #     2) sorted alphabetically for deterministic ordering.
    # This guarantees reproducible JSON output across runs.
    sorted_data = {
        lemma: sorted(set(variants))
        for lemma, variants in sorted(data.items(), key=lambda x: x[0].lower())
    }

    # Persist full normalization structure.
    # merge=False enforces full overwrite semantics
    # (existing file content is replaced, not merged).
    safe_write_json(sorted_data, path, merge=False)

def save_ignored_lemmas(data, path="ignored_lemmas.json"):
    """
    Persist ignored lemmas to a JSON file.

    Behavior:
    - The provided lemmas are written via `safe_write_json` with `merge=True`.
    - If the target file already exists, new lemmas are merged into the existing content.
    - `sort_keys=True` is passed to ensure deterministic JSON key ordering.

    Scope (BETA):
    - No validation of input type or structure is performed.
    - Assumes `data` is an iterable (set or list) of lemma strings.
    - This function performs file-system side effects and does not return a value.

    Parameters:
        data (set | list): Lemmas to be ignored in future processing.
        path (str): Destination file path. Defaults to "ignored_lemmas.json".

    Returns:
        None: Writes JSON file only.
    """
    safe_write_json(data, path, sort_keys=True, merge=True)

def save_lemma_categories(data, path="data/lemma_categories.json"):
    """
    Persist lemma-category mappings to a JSON file.

    Behavior:
    - Existing content is loaded via `safe_read_json(..., default={})`.
    - New entries in `data` update or overwrite existing keys.
    - The resulting mapping is sorted alphabetically by lemma (default string ordering).
    - The final structure is written using `safe_write_json` with `merge=False`
      (full overwrite of the file).

    Scope (BETA):
    - No validation of input structure or label format is performed.
    - Assumes `data` is a mapping of lemma (str) → category label (str).
    - This function performs file-system side effects and does not return a value.

    Parameters:
        data (dict): Mapping from lemma to category label.
        path (str): Destination file path. Defaults to "data/lemma_categories.json".

    Returns:
        None: Writes JSON file only.
    """
    # Load existing category mappings.
    # If the file does not exist, an empty dict is returned (BETA fallback).
    existing = safe_read_json(path, default={})

    # Update existing entries with new data.
    # If a lemma already exists, its category label is overwritten.
    existing.update(data)

    # Sort the combined mapping alphabetically by lemma key
    # (default Python string ordering, case-sensitive).
    # This ensures deterministic output across runs.
    sorted_data = dict(sorted(existing.items()))

    # Persist the fully updated and sorted mapping.
    # merge=False enforces full overwrite semantics.
    safe_write_json(sorted_data, path, merge=False)

def save_json_annotations(path, annotations):
    """
    Persist annotation entries to a JSON file using merge semantics.

    Behavior:
    - Delegates writing to `safe_write_json` with `merge=True`.
    - If the target file already exists, new annotations are merged
      into the existing JSON content.
    - Conflict resolution behavior is determined by `safe_write_json`.

    Scope (BETA):
    - No validation of annotation structure or schema is performed.
    - Assumes `annotations` is a JSON-serializable list.
    - This function performs file-system side effects and does not return a value.

    Parameters:
        path (str): Target file path.
        annotations (list): List of annotation entries to persist.

    Returns:
        None: Writes JSON file only.
    """
    safe_write_json(annotations, path, merge=True)
