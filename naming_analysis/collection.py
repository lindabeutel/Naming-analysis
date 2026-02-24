"""
collection.py

Core data-collection module of the naming-analysis pipeline.

Responsibilities:
- TEI-based detection and extension of naming variants
- Interactive collection and completion of collocation data
- Lemma-based normalization and categorization of naming expressions
  into naming variants and epithets

Design characteristics:
- Designed exclusively for interactive CLI-driven collection sessions
- Operates on in-memory data structures (DataFrame, XML root, JSON containers)
- May trigger persistence via savers.py (progress updates, JSON writes)
- Does not perform analytical visualization (handled in analysis.py)
- Does not perform export formatting (handled in exporter.py)

Semantics:
- Functions may mutate passed-in containers (e.g., dictionaries/lists)
  depending on workflow stage.
- No strict schema validation is performed (BETA state).
- Fallback behavior (e.g., TEI iteration start, missing resources)
  is handled defensively but non-strictly.

Scope:
This module implements the interactive collection logic only.
It is invoked from controller.py during "collect" sessions.
"""
# Standard library
import math
import re
from xml.etree.ElementTree import Element

# Third-party libraries
import pandas as pd

# Project modules
from naming_analysis.io_utils import safe_write_json
from naming_analysis.loaders import (
    load_ignored_lemmas,
    load_lemma_categories,
    load_lemma_normalization,
)
from naming_analysis.savers import (
    save_ignored_lemmas,
    save_lemma_categories,
    save_lemma_normalization,
    save_progress,
)
from naming_analysis.shared import (
    ask_user_choice,
    clean_cell_value,
    get_first_valid_text,
    is_same_verse_number,
    normalize_text,
    parse_verse_number,
    sanitize_cell_value,
)
from naming_analysis.tei_utils import get_valid_verse_number, get_verse_context, tei_ns

# ============================================================================
# Top-level orchestration
# ============================================================================
# Central interactive entry point for collection sessions.
# Coordinates TEI-driven processing (naming detection + collocations)
# and optional lemma-based categorization with resume checkpoints.

def run_data_collection(
    df,
    root,
    naming_variants_dict,
    last_verse,
    paths,
    missing_naming_variants,
    collocation_data,
    check_naming_variants=True,
    perform_collocations=False,
    perform_categorization=False,
    lemma_normalization=None,
    ignored_lemmas=None,
    lemma_categories=None,
    categorized_entries=None
):
    """
    Run the interactive data-collection workflow for the enabled mode(s).

    This function is the central collection loop for "collect" sessions and may
    combine multiple sub-workflows in a single run:

    - Naming variants:
        If `check_naming_variants` is True, iterates over TEI verses (XML root)
        starting after `last_verse` and interactively extends `missing_naming_variants`
        based on matches against `naming_variants_dict` and Excel context where needed.

    - Collocations:
        If `perform_collocations` is True, fills collocation data for verse rows
        using TEI-based iteration (requires `root`).

    - Categorization:
        If `perform_categorization` is True, lemmatizes and categorizes naming
        expressions into lemma slots (Bezeichnung*/Epitheta*) using the optional
        resources `lemma_normalization`, `ignored_lemmas`, and `lemma_categories`.

    Control flow:
        - TEI-based processing is used for naming-variant checks and collocation collection.
          If `root` is missing, TEI-dependent steps return early with diagnostic output.
        - If `check_naming_variants` is False but `perform_categorization` is True,
          an Excel-based loop is used to determine the next starting point for filling
          categorization-related columns (i.e., a starting-index mechanism, not a full
          substitute for TEI-based collection).

    Operational effects:
        - Performs interactive CLI prompts (blocking).
        - Calls `save_progress(...)` repeatedly (typically once per processed verse),
          which writes progress metadata and JSON outputs via `safe_write_json(...)`.

    Mutation vs. copy:
        - The passed-in containers `missing_naming_variants`, `collocation_data`, and
          `categorized_entries` are treated as mutable session state and may be modified
          in-place depending on code paths. The function returns the (potentially mutated)
          containers for explicit downstream use.

    Notes (BETA state):
        - No strict schema validation is performed for input DataFrames, JSON containers,
          or the TEI structure beyond best-effort checks.
        - Fallback behavior is defensive (e.g., early returns if TEI root/verses are missing).

    Returns:
        tuple:
            missing_naming_variants
            collocation_data
            categorized_entries
    """
    # TEI-driven processing block.
    # Handles naming-variant detection and may also trigger collocation and categorization
    # steps depending on active flags. Requires a valid TEI root (`root`).
    if check_naming_variants:
        if root is None:
            print("No TEI root found – cannot perform TEI-based iteration.")
            return missing_naming_variants, collocation_data, categorized_entries

        verse = root.findall('.//tei:l', tei_ns)
        if not verse:
            print("No verses found in TEI.")
            return missing_naming_variants, collocation_data, categorized_entries

        # Determine the first TEI verse whose numeric identifier exceeds `last_verse`.
        # This enables resume behavior across sessions without reprocessing already
        # confirmed verses.
        start_index = next(
            (i for i, line in enumerate(verse) if get_valid_verse_number(line.get("n")) > last_verse),
            0
        )

        print(f"Starting TEI iteration from verse {verse[start_index].get('n')} (Index {start_index})")

        for line in verse[start_index:]:
            verse_number = get_valid_verse_number(line.get("n"))

            # TEI text extraction ('MHDBDB' convention):
            # Aggregate verse text from descendant <seg> elements only (uses seg.text).
            # This assumes tokenized verses as provided by the 'MHDBDB' export format.
            # Text in .tail or outside <seg> elements is not reconstructed.
            # Other TEI encodings may require adapted extraction logic.
            verse_text = ' '.join([seg.text for seg in line.findall(".//tei:seg", tei_ns) if seg.text])
            normalized_verse = normalize_text(verse_text)

            # Interactive naming-variant detection and extension.
            # May modify `missing_naming_variants` in-place depending on user input.
            missing_naming_variants = check_and_extend_namings(
                int(verse_number),
                verse_text,
                normalized_verse,
                df,
                naming_variants_dict,
                missing_naming_variants,
                root,
                paths,
                perform_categorization,
                lemma_normalization,
                ignored_lemmas,
                lemma_categories,
                categorized_entries
            )

            # Collocation collection (TEI-dependent).
            # Operates verse-wise within the TEI loop and updates `collocation_data`.
            # Requires TEI input; cannot run without `root`.
            if perform_collocations:
                rows = df[df["Vers"] == verse_number]

                for _, row in rows.iterrows():
                    check_and_add_collocations(
                        verse_number, collocation_data, root, paths, row=row
                    )

            # Lemma normalization and categorization step.
            # Processes candidate entries for the current verse and updates
            # `categorized_entries` in-place.
            if perform_categorization:
                df_verse = df[(df["Vers"] >= verse_number) & (df["Vers"] < verse_number + 1)]
                entries = df_verse.to_dict(orient="records")

                for entry in entries:
                    source_text = normalize_text(get_first_valid_text(
                        entry.get("Erzähler"),
                        entry.get("Bezeichnung"),
                        entry.get("Eigennennung")
                    ))
                    if not source_text:
                        continue

                    skip = False
                    for e in categorized_entries:
                        if not is_same_verse_number(e.get("Vers", -1), verse_number):
                            continue

                        target_text = normalize_text(get_first_valid_text(
                            e.get("Erzähler"),
                            e.get("Bezeichnung"),
                            e.get("Eigennennung")
                        ))

                        if source_text == target_text and normalize_text(e.get("Benannte Figur", "")) == normalize_text(
                            entry.get("Benannte Figur", "")):
                            if any(
                                str(e.get(k, "")).strip()
                                for k in e.keys()
                                if k.startswith("Bezeichnung") or k.startswith("Epitheta")
                            ):
                                skip = True
                                break

                    if skip:
                        continue

                    annotated = lemmatize_and_categorize_entry(
                        entry, lemma_normalization, paths, ignored_lemmas, lemma_categories
                    )
                    if annotated:
                        categorized_entries.append(annotated)

            # Persist resume state after each processed verse (progress tracking).
            save_progress(
                missing_naming_variants=missing_naming_variants,
                last_processed_verse=int(verse_number),
                paths=paths,
                check_naming_variants=check_naming_variants,
                perform_collocations=perform_collocations,
                perform_categorization=perform_categorization
            )

    # Excel-based resume mechanism.
    # Used to determine the next potential starting point for categorization when naming
    # detection is disabled. This does not replace TEI-driven collection logic.
    elif perform_collocations or perform_categorization:
        print("Starting EXCEL-based iteration over 'Vers' list.")

        # Extract and sort valid verse numbers from Excel (resume window after `last_verse`).
        vers_list = sorted(set(
            v for v in (parse_verse_number(v) for v in df["Vers"])
            if v != -1 and not math.isnan(v) and v > last_verse
        ))

        print(f"Resuming from last edited verse: {last_verse}")

        for verse_number in vers_list:
            verse_number = parse_verse_number(verse_number)

            # Collocations (TEI-dependent; requires `root` for verse context).
            if perform_collocations:
                rows = df[df["Vers"] == verse_number]

                for _, row in rows.iterrows():
                    check_and_add_collocations(
                        verse_number, collocation_data, root, paths, row=row
                    )
            # Categorization resume loop (Excel-based entry selection).
            if perform_categorization:
                df_verse = df[df["Vers"].apply(lambda v: is_same_verse_number(v, verse_number))]
                entries = df_verse.to_dict(orient="records")

                for entry in entries:
                    source_text = normalize_text(get_first_valid_text(
                        entry.get("Erzähler"),
                        entry.get("Bezeichnung"),
                        entry.get("Eigennennung")
                    ))
                    if not source_text:
                        continue

                    skip = False
                    for e in categorized_entries:

                        if not is_same_verse_number(e.get("Vers", -1), verse_number):
                            continue

                        target_text = normalize_text(get_first_valid_text(
                            e.get("Erzähler"),
                            e.get("Bezeichnung"),
                            e.get("Eigennennung")
                        ))

                        if source_text == target_text and normalize_text(e.get("Benannte Figur", "")) == normalize_text(
                            entry.get("Benannte Figur", "")):
                            if any(
                                str(e.get(k, "")).strip()
                                for k in e.keys()
                                if k.startswith("Bezeichnung") or k.startswith("Epitheta")
                            ):
                                skip = True
                                break

                    if skip:
                        continue

                    annotated = lemmatize_and_categorize_entry(
                        entry, lemma_normalization, paths, ignored_lemmas, lemma_categories
                    )
                    if annotated:
                        categorized_entries.append(annotated)

            # Persist resume state after each processed verse (progress tracking).
            save_progress(
                missing_naming_variants=missing_naming_variants,
                last_processed_verse = int(verse_number),
                paths=paths,
                check_naming_variants=check_naming_variants,
                perform_collocations=perform_collocations,
                perform_categorization=perform_categorization
            )

    # Return updated data
    return missing_naming_variants, collocation_data, categorized_entries

# ============================================================================
# Naming variants (TEI-driven detection and extension)
# ============================================================================
# Logic for detecting naming-relevant passages in TEI verses and interactively
# extending the missing_naming_variants container based on the reference dict.

def check_and_extend_namings(
    verse_number: int,
    verse_text: str,
    normalized_verse: str,
    df: pd.DataFrame,
    naming_variants_dict: dict,
    missing_naming_variants: list,
    root: Element,
    paths: dict,
    perform_categorization: bool,
    lemma_normalization: dict,
    ignored_lemmas: set,
    lemma_categories: dict,
    categorized_entries: list
) -> list:
    """
    Detect and optionally confirm naming variants for a TEI verse.

    This function compares:
    (1) naming-related surface strings already present in the Excel data for the current verse
        (columns: "Eigennennung", "Bezeichnung", "Erzähler")
    against
    (2) the reference naming-variant inventory from `naming_variants_dict["Namings"]`.

    For each reference naming variant:
    - Skip if it is already represented in the Excel verse row(s) (substring/token-set heuristics).
    - Skip if it was already handled for this verse in `missing_naming_variants` (JSON container).
    - Detect in the TEI verse using a word-boundary regex against `normalized_verse`.
    - If detected, prompt the user to confirm/reject and, if confirmed, assign the appropriate
      role ("Eigennennung" / "Bezeichnung" / "Erzähler") and collect required metadata.

    Optional workflow integration:
    - If `perform_categorization` is True and the user confirms the entry, the function may
      immediately call `lemmatize_and_categorize_entry(...)` and append the result to
      `categorized_entries`.

    Operational effects:
    - Interactive CLI prompts (blocking).
    - Writes progress checkpoints via `save_progress(...)` when entries are confirmed or rejected.

    Mutation vs. copy:
    - `missing_naming_variants` is treated as mutable session state and is appended to in-place.
      The returned list is the same container reference.
    - If categorization is enabled, `categorized_entries` may also be appended to in-place.

    Notes (BETA state):
    - No strict schema validation is performed for `df`, `naming_variants_dict`, or JSON entries.
    - TEI context display assumes 'MHDBDB'-style tokenization (verse text inside descendant <seg>).
      Text in `.tail` or outside <seg> elements is not reconstructed.

    Returns:
        list: The (potentially extended) `missing_naming_variants` container.
    """
    # -------------------------------------------------------------------------
    # 1) Collect naming-relevant surface strings already present in Excel
    # -------------------------------------------------------------------------
    existing_naming_variants = set()
    if "Vers" in df.columns:
        df_verse = df[df["Vers"] == verse_number]
        for column in ["Eigennennung", "Bezeichnung", "Erzähler"]:
            if column in df_verse.columns:
                values = df_verse[column].dropna().tolist()
                existing_naming_variants.update(
                    normalize_text(str(value).strip()) for value in values if str(value).strip()
                )


    # -------------------------------------------------------------------------
    # 2) Build reference naming-variant set from naming_variants_dict
    # -------------------------------------------------------------------------
    dict_naming_variants = set()
    for book_list in naming_variants_dict.get("Namings", {}).values():
        dict_naming_variants.update(
            normalize_text(name.strip()) for name in book_list if name.strip()
        )

    # -------------------------------------------------------------------------
    # 3) Detection loop: skip heuristics → TEI match → interactive confirmation
    # -------------------------------------------------------------------------
    for naming_variant in dict_naming_variants:
        if not naming_variant:
            continue

        # --- Skip if already represented in Excel (substring + token-set heuristics) ---
        naming_variant_tokens = set(naming_variant.split())

        skip_existing = False
        for entry in existing_naming_variants:
            entry_tokens = set(entry.split())
            if naming_variant in entry or entry in naming_variant:
                skip_existing = True
                break
            if naming_variant_tokens <= entry_tokens or entry_tokens <= naming_variant_tokens:
                skip_existing = True
                break

        if skip_existing:
            continue

        # --- Skip if already handled in JSON container for this verse ---
        skip = False
        for entry in missing_naming_variants:
            if entry.get("Vers") == verse_number:
                values = [
                    entry.get("Eigennennung", ""),
                    entry.get("Bezeichnung", ""),
                    entry.get("Erzähler", "")
                ]
                if normalize_text(naming_variant) in map(normalize_text, values):
                    skip = True
                    break
        if skip:
            continue

        # --- Detection: word-boundary match against normalized verse string ---
        if not re.search(rf'\b{re.escape(naming_variant)}\b', normalized_verse):
            continue


        # ---------------------------------------------------------------------
        # 4) Context display (TEI-based; 'MHDBDB' convention: descendant <seg>.text)
        # ---------------------------------------------------------------------
        print("\n" + "-" * 60)
        print(f"New naming variant found that is not listed in the Excel file!")
        print(f"Detected naming variant: \"{naming_variant}\"")

        # ---------------------------------------------------------------------
        # TEI CONTEXT (decision support only)
        # ---------------------------------------------------------------------
        # This context display is purely informational.
        # It helps the user evaluate whether the detected surface string
        # truly functions as a naming variant in this passage.
        #
        # IMPORTANT:
        # - This block does NOT influence detection logic.
        # - Naming detection is based solely on normalized_verse + regex.
        #
        # TEI assumption ('MHDBDB' convention):
        # - Verse text is reconstructed from descendant <seg>.text elements.
        # - .tail text and non-<seg> content are not reconstructed.
        # - This is suitable for 'MHDBDB' TEI exports and may require adaptation
        #   for other TEI encodings.
        prev_line = root.find(f'.//tei:l[@n="{verse_number - 1}"]', tei_ns)
        if prev_line is not None:
            prev_text = ' '.join([seg.text for seg in prev_line.findall('.//tei:seg', tei_ns) if seg.text])
            print(f"Previous verse ({verse_number - 1}): {prev_text}")

        highlighted = verse_text.replace(naming_variant, f"\033[1m\033[93m{naming_variant}\033[0m")
        print(f"Verse ({verse_number}): {highlighted}")

        next_line = root.find(f'.//tei:l[@n="{verse_number + 1}"]', tei_ns)
        if next_line is not None:
            next_text = ' '.join([seg.text for seg in next_line.findall('.//tei:seg', tei_ns) if seg.text])
            print(f"Next verse ({verse_number + 1}): {next_text}")

        # ---------------------------------------------------------------------
        # 5) Interactive decision: confirm / reject
        # ---------------------------------------------------------------------
        confirm = ask_user_choice("Is this a missing naming variant? (y/n): ", ["y", "n"])
        if confirm == "n":
            missing_naming_variants.append({
                "Vers": verse_number,
                "Eigennennung": naming_variant,
                "Nennende Figur": "",
                "Bezeichnung": "",
                "Erzähler": "",
                "Status": "rejected"
            })
            save_progress(missing_naming_variants, verse_number, paths)
            print("Rejection saved.")
            continue

        # ---------------------------------------------------------------------
        # 6) Optional adjustment: shorten/lengthen surface string
        # ---------------------------------------------------------------------
        extend = ask_user_choice("Would you like to shorten or lengthen the naming variant (y/n): ", ["y", "n"])
        if extend == "y":
            naming_variant = input("Enter the adapted naming variant: ").strip()

        # ---------------------------------------------------------------------
        # 7) Classification + required metadata (Benannte Figur / optional Nennende Figur)
        # ---------------------------------------------------------------------
        print("Please choose the correct category:")
        print("[1] Eigennennung")
        print("[2] Bezeichnung")
        print("[3] Erzähler")
        print("[4] Skip")

        choice = input("Your selection: ").strip()
        if choice == "4":
            continue

        named_entity = input("Enter the \"Benannte Figur\": ").strip()
        naming_entity = ""
        if choice == "2":
            naming_entity = input("Enter the \"Nennende Figur\": ").strip()

        entry = {
            "Benannte Figur": named_entity,
            "Vers": verse_number,
            "Eigennennung": naming_variant if choice == "1" else "",
            "Nennende Figur": naming_entity,
            "Bezeichnung": naming_variant if choice == "2" else "",
            "Erzähler": naming_variant if choice == "3" else "",
            "Status": "confirmed"
        }

        # ---------------------------------------------------------------------
        # 8) Optional: attach collocation context (interactive selection)
        # ---------------------------------------------------------------------
        wants_collocation = ask_user_choice("Do you want to add a collocation (context lines)? (y/n): ", ["y", "n"])

        if wants_collocation == "y":
            # ---------------------------------------------------------------------
            # EXTENDED TEI CONTEXT (collocation selection only)
            # ---------------------------------------------------------------------
            # This block is used exclusively for defining collocation spans.
            # The numbered verse window (-6 … +6) allows the user to select
            # exactly which verses should be stored as contextual collocation.
            #
            # Unlike the previous context block:
            # - This step is data-generating (adds "Kollokation" to the entry).
            # - It is fully optional and user-driven.
            #
            # The same 'MHDBDB'-based <seg>.text extraction logic is used here.
            print("\nExtended context (1–13):")
            context_lines = {}
            number = 1

            for i in range(6, 0, -1):
                line = root.find(f'.//tei:l[@n="{verse_number - i}"]', tei_ns)
                if line is not None:
                    text = ' '.join([seg.text for seg in line.findall('.//tei:seg', tei_ns) if seg.text])
                    context_lines[number] = text
                    print(f"[{number}] {text}")
                    number += 1

            context_lines[number] = verse_text
            print(f"[{number}] {verse_text}")
            number += 1

            for i in range(1, 7):
                line = root.find(f'.//tei:l[@n="{verse_number + i}"]', tei_ns)
                if line is not None:
                    text = ' '.join([seg.text for seg in line.findall('.//tei:seg', tei_ns) if seg.text])
                    context_lines[number] = text
                    print(f"[{number}] {text}")
                    number += 1

            selection = input("\nPlease enter the line number(s) (e.g., '5-7' or '6'): ").strip()
            selected = []

            try:
                if "-" in selection:
                    start, end = map(int, selection.split("-"))
                    selected = [context_lines[i] for i in range(start, end + 1) if i in context_lines]
                else:
                    idx = int(selection)
                    selected = [context_lines[idx]]
            except (ValueError, KeyError):
                print("Invalid input – no collocation saved.")

            if selected:
                entry["Kollokation"] = ' / '.join(selected)

        # ---------------------------------------------------------------------
        # 9) Persist: store confirmed entry and checkpoint progress
        # ---------------------------------------------------------------------
        missing_naming_variants.append(entry)
        save_progress(missing_naming_variants, verse_number, paths)
        print("Entry saved.")


        # ---------------------------------------------------------------------
        # 10) Optional: immediate categorization hook (if enabled)
        # ---------------------------------------------------------------------
        if perform_categorization and entry["Status"] == "confirmed":
            annotated = lemmatize_and_categorize_entry(
                entry,
                lemma_normalization,
                paths,
                ignored_lemmas,
                lemma_categories
            )
            if annotated:
                categorized_entries.append(annotated)

    return missing_naming_variants

# ============================================================================
# Collocations (interactive collection; TEI-dependent context)
# ============================================================================
# Utilities to gather and persist collocations for a given verse row.
# Collocation prompting is interactive and uses TEI context reconstruction.

def check_and_add_collocations(verse_number, collocation_data, root, paths, row):
    """
    Collect collocations for a single verse row if they are missing in Excel and JSON.

    This function is invoked during TEI-driven collection runs. It checks whether
    the Excel column "Kollokationen" is empty. If so, it reconstructs a TEI-based
    verse context and prompts the user to select relevant lines.

    Duplicate handling:
    - If a matching entry (same verse, same named entity, same naming surface form)
      already exists in `collocation_data` and contains non-empty collocations,
      no new entry is added.
    - Duplicate detection uses `normalize_text` for robust comparison, but
      stored values remain unchanged.

    Operational effects:
    - Interactive CLI prompts (blocking).
    - Immediate persistence of `collocation_data` via `safe_write_json`.

    Mutation vs. copy:
    - `collocation_data` is mutated in-place when a new entry is added.

    Notes (BETA state):
    - TEI context reconstruction is delegated to `get_verse_context(...)`
      and assumes 'MHDBDB'-style tokenized <l>/<seg> structure.
    - No strict schema validation of Excel or JSON structure is performed.

    Returns:
        bool | None:
            True  -> a new collocation entry was appended and persisted.
            None  -> no new entry was created (Excel already filled,
                      or JSON already contains a non-empty entry).
    """
    # -------------------------------------------------------------------------
    # 1) Skip if Excel already contains collocation data
    # -------------------------------------------------------------------------
    if sanitize_cell_value(row.get("Kollokationen")) != "":
        return None

    # -------------------------------------------------------------------------
    # 2) Normalize verse number using project parser (supports decimal verses)
    # -------------------------------------------------------------------------
    # Important:
    # - Sub-verses (e.g., 8158.15) are preserved.
    # - TEI lookup is strict and requires an exact verse identifier.
    # - No fallback to the base verse (e.g., 8158) is performed,
    #   in order to avoid displaying semantically incorrect context.
    verse_number = parse_verse_number(verse_number)
    if verse_number == -1:
        return None

    # -------------------------------------------------------------------------
    # 3) Extract naming surface form and named entity from Excel row
    # -------------------------------------------------------------------------
    naming_variant = clean_cell_value(row.get("Eigennennung")) \
             or clean_cell_value(row.get("Bezeichnung")) \
             or clean_cell_value(row.get("Erzähler"))

    named_entity = clean_cell_value(row.get("Benannte Figur"))

    # -------------------------------------------------------------------------
    # 4) Duplicate detection (normalized comparison only)
    # -------------------------------------------------------------------------
    named_entity_normalized = normalize_text(named_entity or "")
    naming_variant_normalized = normalize_text(naming_variant or "")

    if any(
        is_same_verse_number(entry.get("Vers", -1), verse_number)
        and normalize_text(entry.get("Benannte Figur", "")) == named_entity_normalized
        and normalize_text(entry.get("Naming", "")) == naming_variant_normalized
        and str(entry.get("Kollokationen", "")).strip()
        for entry in collocation_data
    ):
        return None

    # -------------------------------------------------------------------------
    # 5) Reconstruct TEI context (decision & selection support only)
    # -------------------------------------------------------------------------
    # Context is retrieved strictly for the normalized verse identifier.
    # If no matching TEI <l @n="..."> exists (e.g., due to Excel/TEI divergence),
    # no context will be displayed. This is intentional and methodologically strict.
    context = get_verse_context(verse_number, root)

    # -------------------------------------------------------------------------
    # 6) Interactive collocation selection
    # -------------------------------------------------------------------------
    collocations = ask_for_collocations(
        verse_number,
        named_entity,
        naming_variant,
        context
    )

    # Guard: avoid persisting empty collocation selections
    if not isinstance(collocations, str) or not collocations.strip():
        return None

    # -------------------------------------------------------------------------
    # 7) Append new entry and persist immediately
    # -------------------------------------------------------------------------
    collocation_data.append({
        "Vers": verse_number,
        "Benannte Figur": named_entity,
        "Naming": naming_variant,
        "Kollokationen": collocations
    })

    safe_write_json(paths["collocations_json"], collocation_data)

    return True

def ask_for_collocations(verse_number, named_entity, naming_variant, context):
    """
    Interactively collects a collocation span from a TEI-derived verse context.

    The function displays a numbered verse window (typically ±6 lines around
    the current verse) and allows the user to select either:
        - a single line number (e.g., "5"),
        - a contiguous range (e.g., "5-7"),
        - or press Enter to explicitly skip collocation assignment
          (requires confirmation).

    Input validation:
        - Only numbers in the displayed range are accepted.
        - Only formats "N" or "N-N" are permitted.
        - Invalid or out-of-range selections trigger a repeat prompt.
        - An empty input triggers a confirmation step before skipping.

    Highlighting:
        The provided naming_variant is highlighted in the displayed context
        using substring-based replacement. This is intentionally tolerant to
        orthographic variation across editions, as verse matching relies on
        verse_number rather than strict token equality.

    Parameters:
        verse_number: Canonically parsed verse identifier.
        named_entity: The referenced figure (display only).
        naming_variant: The naming surface form (used for highlighting).
        context: Iterable of (number, text) tuples representing the
                 TEI-derived context window.

    Returns:
        str:
            - A " / "-joined string of the selected context lines.
            - An empty string if the user confirms skipping.
    """
    print(f"\nEmpty collocation field detected in verse {verse_number}!")
    if named_entity or naming_variant:
        print(f"{named_entity}: {naming_variant}\n")

    for number, text in context:
        if naming_variant:
            # Highlight naming_variant using substring replacement.
            # This is intentionally tolerant to orthographic variation
            # across editions. Verse identification relies on verse_number,
            # not strict token equality.
            highlighted = text.replace(str(naming_variant), f"\033[1;33m{naming_variant}\033[0m")
        else:
            highlighted = text
        print(f"{number}. {highlighted}")

    # Ensure selected is defined across all control paths
    selected = []

    while True:
        user_input = input(
            "\nPlease enter the number(s) of the relevant lines (e.g., '5' or '5-7') "
            "or press Enter to skip: "
        ).strip()

        # Explicit skip path: user must confirm leaving collocation empty.
        # Prevents accidental data loss due to unintended Enter key.
        if user_input == "":
            confirm_skip = ask_user_choice(
                "Are you sure you want to skip and leave the collocation empty? (y/n): ",
                ["y", "n"]
            )
            if confirm_skip == "y":
                print("Skipped: collocation will remain empty.")
                return ""
            continue

        # Accept only a single number or a numeric range (N or N-N).
        # Structural validation is performed via regex before parsing.
        if not re.fullmatch(r"\d{1,2}(\s*-\s*\d{1,2})?", user_input):
            print("Invalid input. Use a single number (1–13), a range (e.g., 5-7), or press Enter to skip.")
            continue

        try:
            if "-" in user_input:
                start, end = map(int, user_input.split("-"))
                selected = [text for number, text in context if start <= number <= end]
            else:
                number = int(user_input)
                selected = [text for num, text in context if num == number]

            # Prevent silent acceptance of out-of-range selections.
            # If no context lines match the numeric input, re-prompt.
            if not selected:
                print("No matching context lines found. Please try again.")
                continue

            break

        except ValueError:
            print("Invalid input. Please enter a single number or a range.")

    return " / ".join(selected)

# ============================================================================
# Categorization (lemmatization + assignment into Bezeichnung*/Epitheta* slots)
# ============================================================================
# Interactive lemmatization/categorization for individual entries, plus the
# higher-level categorization driver used by the collection workflow.

def lemmatize_and_categorize_entry(entry, lemma_normalization, paths, ignored_lemmas=None, lemma_categories=None):
    """
    Annotate a single naming entry by resolving its tokens to lemmas
    and interactively assigning each lemma to a category.

    Workflow:
    1. Determine the first available textual field (Erzähler / Bezeichnung / Eigennennung)
       to use as annotation basis.
    2. Tokenize the text (lowercased, alphabetic tokens only).
    3. Ensure lemma normalization coverage:
       - Unknown tokens are collected.
       - The user is prompted to provide corresponding lemma(s).
       - The lemma_normalization mapping is updated and written immediately
         to disk (write-through persistence).
    4. Resolve tokens to lemmas.
    5. Run interactive categorization for each lemma via `run_categorization`,
       assigning:
           - "a" → naming variant
           - "e" → epithet
       The user may step back ("<") within that process.
    6. If no categories are assigned, the user must explicitly confirm skipping
       the entry; otherwise, categorization restarts for the same lemma sequence.
    7. Build a flat annotated entry (Bezeichnung 1–4, Epitheta 1–5) and append it
       to the categorization JSON file using merge semantics.

    Persistence model:
    - lemma_normalization, ignored_lemmas, and lemma_categories are loaded once
      (if not provided) and modified in memory.
    - lemma_normalization is saved immediately when extended.
    - lemma_categories and ignored_lemmas are saved within
      `run_categorization` (write-through).
    - The final annotated entry is appended via `safe_write_json(..., merge=True)`,
      which reloads and merges with existing file content.

    Parameters:
        entry (dict):
            Naming entry containing at least keys such as
            "Vers", "Benannte Figur", and one of
            "Erzähler", "Bezeichnung", or "Eigennennung".
        lemma_normalization (dict | None):
            Mapping of lemma → list of normalized surface forms.
            If None, loaded from paths["lemma_normalization_json"].
        paths (dict):
            Dictionary containing required file paths.
        ignored_lemmas (set | None):
            Set of lemmas to skip during categorization.
            If None, loaded from paths["ignored_lemmas_json"].
        lemma_categories (dict | None):
            Mapping of lemma → category label ("a" or "e").
            If None, loaded from paths["lemma_categories_json"].

    Returns:
        dict | None:
            The annotated entry (including Bezeichnung/Epitheta slots),
            or None if the user confirms skipping the entry.
    """
    # Load normalization and categorization resources if not provided externally.
    # These objects are kept in memory during the function call.
    if lemma_normalization is None:
        lemma_normalization = load_lemma_normalization(paths["lemma_normalization_json"])

    if ignored_lemmas is None:
        ignored_lemmas = load_ignored_lemmas(paths["ignored_lemmas_json"])

    if lemma_categories is None:
        lemma_categories = load_lemma_categories(paths["lemma_categories_json"])

    # Determine the first available textual field to annotate.
    # Priority order differs from type display below.
    text = get_first_valid_text(
        entry.get("Erzähler"),
        entry.get("Bezeichnung"),
        entry.get("Eigennennung")
    )

    # Skip entries without usable text.
    if not text:
        print("No text to annotate – entry skipped.\n")
        return None

    print("\n" + "=" * 60)
    print(f"Verse: {entry.get('Vers')}")
    print(f"Named Entity: {entry.get('Benannte Figur')}")

    # Determine display type based on first non-empty naming field.
    # This is purely informational (CLI output only).
    first_text = get_first_valid_text(
        entry.get("Eigennennung"),
        entry.get("Bezeichnung"),
        entry.get("Erzähler")
    )

    typ = "(unbestimmt)"

    if first_text == entry.get("Eigennennung"):
        typ = "Eigennennung"
    elif first_text == entry.get("Bezeichnung"):
        typ = "Bezeichnung"
    elif first_text == entry.get("Erzähler"):
        typ = "Erzähler"

    print(f"Type: {typ}")
    print(f"\nOriginal text: {text}")

    # Tokenize lowercased text and keep alphabetic tokens only.
    tokens = [t for t in tokenize(text.lower()) if t.isalpha()]

    # Identify tokens that are not yet covered by lemma_normalization.
    # Matching is exact (no fuzzy/substring logic).
    missing = [
        t for t in tokens
        if t.isalpha() and not any(t in v or t == k for k, v in lemma_normalization.items())
    ]

    # If uncovered tokens exist, require explicit lemma assignment.
    if missing:
        while True:
            print(f"\nPlease add lemma(s) for {', '.join(missing)} (comma-separated):")
            user_input = input("> ").strip()
            new_lemmas = [l.strip() for l in user_input.split(",") if l.strip()]

            # Enforce one-to-one mapping between tokens and provided lemmas.
            if len(new_lemmas) == len(missing):
                break

            print(
                f"Number of lemmas ({len(new_lemmas)}) doesn't match number of tokens ({len(missing)})."
                f"Please try again."
            )

        # Extend normalization mapping.
        for token, lemma in zip(missing, new_lemmas):
            lemma_normalization.setdefault(lemma, [])
            if token not in lemma_normalization[lemma]:
                lemma_normalization[lemma].append(token)

        # Normalize structure: unique surface forms per lemma, sorted alphabetically.
        for lemma in lemma_normalization:
            lemma_normalization[lemma] = sorted(set(lemma_normalization[lemma]))

        # Persist updated normalization immediately (write-through).
        save_lemma_normalization(lemma_normalization, path=paths["lemma_normalization_json"])

    # Resolve each token to its lemma representation.
    lemmas = [resolve_lemma(t, lemma_normalization) for t in tokens]
    print(f"\nLemma: {', '.join(lemmas)}\n")

    # Interactive categorization loop.
    # Ensures that empty categorization must be explicitly confirmed as skip.
    while True:
        naming_variants, epithets = run_categorization(
            lemmas, lemma_categories, ignored_lemmas, paths
        )

        # Prevent accidental skip via empty input.
        if not naming_variants and not epithets:
            print("No entry – please review and confirm again.")
            confirm = ask_user_choice("Really skip this entry? [y = yes / n = no]: ", ["y", "n"])

            if confirm == "y":
                print("Entry skipped.\n")
                return None
            else:
                # Treat empty output as accidental input unless explicitly confirmed as skip.
                # Restart categorization for the same entry.
                continue

        else:
            break

    # Build flattened export structure with fixed slot schema.
    # Empty strings fill unused slots to maintain Excel compatibility.
    annotated_entry = {
        **entry,
        "Bezeichnung 1": naming_variants[0] if len(naming_variants) > 0 else "",
        "Bezeichnung 2": naming_variants[1] if len(naming_variants) > 1 else "",
        "Bezeichnung 3": naming_variants[2] if len(naming_variants) > 2 else "",
        "Bezeichnung 4": naming_variants[3] if len(naming_variants) > 3 else "",
        "Epitheta 1": epithets[0] if len(epithets) > 0 else "",
        "Epitheta 2": epithets[1] if len(epithets) > 1 else "",
        "Epitheta 3": epithets[2] if len(epithets) > 2 else "",
        "Epitheta 4": epithets[3] if len(epithets) > 3 else "",
        "Epitheta 5": epithets[4] if len(epithets) > 4 else ""
    }

    # Append entry to categorization JSON.
    # safe_write_json handles reload, merge, de-duplication and verse normalization.
    safe_write_json([annotated_entry], paths["categorization_json"], merge=True)

    print("Entry saved.\n")
    return annotated_entry

def run_categorization(lemmas, lemma_categories, ignored_lemmas, paths):
    """
    Interactive helper for assigning lemma categories within a single entry.

    The function iterates sequentially over the provided `lemmas` list and
    prompts the user to assign each lemma to one of the following categories:

        - "a" → naming variant
        - "e" → epithet
        - ignore (with confirmation)
        - "<" → step back to the previous decision

    Behavior and semantics:

    - Processing is strictly lemma-by-lemma (one lemma per iteration).
    - If a lemma already has a stored category in `lemma_categories`,
      pressing Enter accepts this default.
    - Ignored lemmas are skipped automatically in subsequent runs.
    - The "<" command reverts the most recent action recorded in `history`:
        * For "a"/"e": removes the lemma from the corresponding local result list.
        * For "ignore": removes the lemma from `ignored_lemmas`.
        * For "override": deletes the category entry from `lemma_categories`
          (no historical category restoration is performed).
    - Persistent structures (`lemma_categories`, `ignored_lemmas`) are written
      immediately after each modifying action (write-through persistence).
      This ensures crash tolerance and allows manual inspection/editing
      during runtime.

    Local vs. persistent state:

    - `naming_variants` and `epithets` are local result containers for the
      current categorization call and are not written to disk here.
    - `lemma_categories` and `ignored_lemmas` represent persistent,
      cross-session knowledge and are updated incrementally.

    Notes (BETA state):

    - No schema validation is performed for input arguments.
    - No locking mechanism is used for concurrent manual edits of JSON files.
    - Undo operations are not fully reversible in a historical sense;
      overridden categories are deleted rather than restored.

    Returns:
        tuple[list[str], list[str]]:
            Two lists containing naming variants and epithets
            in the order they were confirmed during this session.
    """

    naming_variants = []
    epithets = []
    history = []
    i = 0

    # Main interactive loop: strictly one lemma per iteration.
    # `i` moves forward after a confirmed decision and backward on "<".
    while i < len(lemmas):
        lemma = lemmas[i]

        # Automatically skip lemmas that are globally marked as ignored.
        # These are not shown again to the user.
        if lemma in ignored_lemmas:
            i += 1
            continue

        # UI-only default representation.
        # If the lemma already has a stored category ("a"/"e"),
        # it is displayed in brackets and can be accepted via Enter.
        default = f"[{lemma_categories.get(lemma, '')}]" if lemma in lemma_categories else ""
        print(f"{lemma:<12} → {default} ", end="")
        user_input = input().strip()

        # Step-back command.
        # Reverts exactly one previously recorded action from `history`.
        if user_input == "<":
            if i == 0 or not history:
                print("Already at beginning – can't step back.")
                continue

            # Move index back to the previous lemma.
            i -= 1
            last_action = history.pop()

            # Undo local result assignment.
            if last_action["type"] == "a":
                naming_variants.pop()
            elif last_action["type"] == "e":
                epithets.pop()

            # Undo global ignore state (write-through).
            elif last_action["type"] == "ignore":
                ignored_lemmas.discard(last_action["lemma"])
                save_ignored_lemmas(ignored_lemmas, path=paths["ignored_lemmas_json"])

            # Undo override by deleting the stored category.
            # No historical category restoration is performed.
            elif last_action["type"] == "override":
                del lemma_categories[last_action["lemma"]]
                save_lemma_categories(lemma_categories, path=paths["lemma_categories_json"])

            continue

        # Accept stored default category via Enter.
        # Uses UI representation "[a]" / "[e]" for comparison.
        if user_input == "" and default:
            if default == "[a]":
                naming_variants.append(lemma)
                history.append({"type": "a", "lemma": lemma})
            elif default == "[e]":
                epithets.append(lemma)
                history.append({"type": "e", "lemma": lemma})

            i += 1
            continue

        # Empty input without default → interpret as "ignore" candidate.
        # Confirmation required to prevent accidental ignores.
        if user_input == "":
            confirm_ignore = ask_user_choice(f"Really ignore lemma “{lemma}”? [y/n]: ", ["y", "n"])
            if confirm_ignore == "y":
                ignored_lemmas.add(lemma)
                save_ignored_lemmas(ignored_lemmas, path=paths["ignored_lemmas_json"])
                print(f"Lemma “{lemma}” added to ignore list.")
                history.append({"type": "ignore", "lemma": lemma})
                i += 1
                continue
            else:
                print("Skipped ignoring – please choose a category or go back.\n")
                continue

        # Direct category assignment ("a" or "e").
        # Updates both local result list and persistent category memory.
        if user_input in ("a", "e"):
            if user_input == "a":
                naming_variants.append(lemma)
            else:
                epithets.append(lemma)

            lemma_categories[lemma] = user_input
            save_lemma_categories(
                lemma_categories,
                path=paths["lemma_categories_json"]
            )

            history.append({"type": user_input, "lemma": lemma})
            i += 1
            continue

        # Any other input is interpreted as an override:
        # user provides a corrected lemma string and must define its category.
        correction = user_input
        cat = ""

        # Force valid category input for override.
        while cat not in ("a", "e"):
            cat = input(f'Define category for “{correction}” [a/e]: ').strip().lower()

        # Assign override lemma locally.
        if cat == "a":
            naming_variants.append(correction)
        else:
            epithets.append(correction)

        # Persist override category (write-through).
        lemma_categories[correction] = cat
        save_lemma_categories(
            lemma_categories,
            path=paths["lemma_categories_json"]
        )

        # Record override in history for single-step undo.
        history.append({"type": "override", "lemma": correction})
        i += 1

    # Return local categorization results for the current entry.
    return naming_variants, epithets

# ============================================================================
# Tokenization and lemma resolution utilities
# ============================================================================
# Small text helpers used by categorization:
# - tokenize: split surface strings into candidate tokens
# - resolve_lemma: map token -> lemma using normalization dictionary logic

def tokenize(text):
    """
    Split a string into discrete tokens using a simple regular expression–based segmentation.

    This function performs a technical tokenization intended for internal
    lemma categorization workflows. It does not implement linguistic
    tokenization and does not apply any normalization (e.g., no lowercasing,
    no diacritic harmonization, no stemming).

    Token definition:
    - Alphanumeric sequences matched by \\w+ (Unicode-aware)
    - Individual punctuation characters matched by [^\\w\\s]

    Whitespace is treated purely as a separator and is not preserved.
    The function is deterministic and returns tokens derived solely from
    the provided input string.

    Parameters:
        text (str): The input string to tokenize.

    Returns:
        list[str]: A list of tokens extracted from the input text.
    """
    return re.findall(r'\w+|[^\w\s]', text, re.UNICODE)

def resolve_lemma(token: str, lemma_dict: dict[str, list[str]]) -> str:
    """
    Resolve a token to its corresponding lemma using an explicit variant mapping.

    The mapping dictionary must follow the structure:
        {lemma: [variant1, variant2, ...]}

    Resolution is based on exact string comparison. If the provided token
    matches any listed variant (using direct equality), the corresponding
    lemma is returned. No normalization is applied (e.g., no lowercasing,
    no diacritic harmonization, no fuzzy matching).

    If no variant matches, the original token is returned unchanged
    (fallback behaviour).

    If a token appears in multiple variant lists, the first matching lemma
    encountered during dictionary iteration is returned.

    Parameters:
        token (str): The word form to resolve.
        lemma_dict (dict[str, list[str]]): Mapping of lemma → variant list.

    Returns:
        str: The resolved lemma or the original token if no match is found.
    """
    for lemma, variants in lemma_dict.items():
        if token in variants:
            return lemma
    return token