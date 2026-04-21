"""
analysis.py

Core analysis module of the naming-analysis pipeline.

Responsibilities:
- Interactive CLI dispatchers for selecting analysis workflows
- Frequency-based wordlist generation from categorized naming data
- Naming figure analyses (overview tables and profile exports)
- Keyword analysis using log-likelihood (G²)
- Collocation analysis and KWIC display
- Plotly-based visualization routines

Design characteristics:
- Designed for interactive CLI-driven analysis sessions
- Operates on categorized JSON data (primary) with optional Excel fallback
- Uses centralized source-selection and requirement guards from shared helpers
- Writes CSV exports to the analysis directory of the active book
- Delegates visualization output handling to io_utils.py
- Does not perform data collection (handled in collection.py)
- Does not perform low-level persistence logic (handled in savers.py)

Semantics:
- Analytical functions operate on filtered DataFrame views derived from
  categorized data.
- Canonical figure names are expected to be validated interactively
  before analytical computation.
- Matching logic (e.g., lemma matching, name detection) relies on
  explicit helper functions defined in shared.py.
- No strict schema validation or automatic repair is performed (BETA state).
- Structural requirements are enforced via explicit requirement guards.

Scope:
This module implements analytical logic only.
It is invoked from controller.py during "analysis" sessions.
"""
# Standard library
import os
import math
import difflib
from collections import Counter
from datetime import datetime
from itertools import combinations
from pathlib import Path
from typing import Any, cast

# Third-party libraries
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
from plotly.colors import n_colors

# Internal project imports
from naming_analysis.shared import (
    ask_user_choice,
    get_first_valid_text,
    hex_color_to_rgba,
    hex_color_to_rgb_tuple,
    apply_accessible_text_colors,
    resolve_figure_name,
    rgb_tuple_to_plotly_color,
    format_kwic,
    list_available_reference_books,
    prepare_naming_data,
    serialize_verse_value,
    select_naming_data,
    check_naming_requirements,
    match_name_to_lemma,
    parse_token_selection,
    extract_tokens,
    collect_tokens_for_cooccurrence,
    resolve_name_lemmas_for_figure,
)
from naming_analysis.io_utils import (
    safe_read_json,
    write_csv_table,
    export_visualization_output,
)
from naming_analysis.loaders import (
    load_collocation_sheet,
    build_fallback_collocation_df_from_tei,
    load_naming_sources_with_excel_fallback,
)

# =============================================================================
# MAIN ANALYSIS ENTRY (central dispatcher)
# =============================================================================

def run_analysis_menu(config_data, paths, data, book_name):
    """
    Central interactive dispatcher for analysis workflows.

    This function provides the top-level CLI menu for selecting one of the
    available analysis paths (wordlists, naming figure analysis, keywords,
    collocations, or visualization). It delegates execution to the respective
    submenu functions based on validated user input.

    The function runs in a blocking loop until the user selects the exit option.
    It does not perform any data transformation itself but orchestrates calls
    to downstream analysis components.

    Parameters:
        config_data (dict):
            Loaded configuration settings required by certain analysis paths
            (e.g., keyword and collocation analysis).

        paths (dict):
            Dictionary containing resolved project paths (e.g., data directories,
            export locations).

        data (dict):
            Loaded project data (e.g., TEI structures and Excel-derived tables)
            required by analysis submodules.

        book_name (str):
            Short identifier of the currently active book; passed to submenus
            for context-sensitive analysis and export naming.

    Returns:
        None

    Behavior:
        Prints interactive prompts to stdout and waits for user input.
        Delegates control flow to subordinate analysis menu functions.
    """
    while True:
        print("Which type of analysis do you want to run?")
        print("[1] Wordlist")
        print("[2] Naming figure analysis")
        print("[3] Keywords")
        print("[4] Collocations")
        print("[5] Visualization")
        print("[6] Exit analysis menu")

        choice = ask_user_choice("> ", ["1", "2", "3", "4", "5","6"])

        if choice == "1":
            run_wordlist_menu(paths, book_name)
        elif choice == "2":
            run_naming_figure_analysis(config_data, paths, data, book_name)
        elif choice == "3":
            run_keyword_menu(config_data, paths, data, book_name)
        elif choice == "4":
            run_collocation_menu(config_data, paths, data, book_name)
        elif choice == "5":
            run_visualization_menu(paths, book_name, data)
        elif choice == "6":
            print("Exiting analysis.")
            break

# =============================================================================
# WORDLIST ANALYSIS (menu + generators)
# =============================================================================

def run_wordlist_menu(paths, book_name):
    """
    Interactive CLI menu for generating wordlist exports from categorization data.

    This function provides an interactive selection interface for producing
    frequency-based wordlists derived from the categorization JSON file of the
    currently active work. Depending on the user’s choice, it delegates the
    generation task to the corresponding wordlist function.

    Available wordlist types:
        1. All values from a selected categorical column
           (e.g., "Benannte Figur", "Bezeichnung", "Epitheta").
        2. All naming variants for a selected figure.
        3. All epithets for a selected figure.
        4. A combined list of naming variants and epithets for a selected figure.

    Output files are written as CSV to:
        data/<book_name>/analysis/

    The function runs in a blocking loop until the user selects the option
    to return to the main analysis menu.

    Parameters:
        paths (dict):
            Dictionary containing resolved project paths. Must include
            the key "categorization_json".

        book_name (str):
            Short identifier of the currently active work. Used for
            output directory and filename construction.

    Returns:
        None
    """
    json_path = paths["categorization_json"]
    output_dir = os.path.join("data", book_name, "analysis")
    os.makedirs(output_dir, exist_ok=True)

    while True:
        print("\nWhat kind of wordlist do you want to generate?")
        print("[1] All values from a column (e.g., 'Benannte Figur')")
        print("[2] All naming variants for a specific figure")
        print("[3] All epithets for a specific figure")
        print("[4] Combined naming variants and epithets")
        print("[5] Back to main analysis menu")

        menu_choice = ask_user_choice("> ", ["1", "2", "3", "4", "5"])

        if menu_choice == "1":
            print("\nAvailable columns:")
            print("[1] Benannte Figur")
            print("[2] Bezeichnung (Naming variants)")
            print("[3] Epitheta (Epithets)")

            column_choice = ask_user_choice("> ", ["1", "2", "3"])

            column_input = {
                "1": "Benannte Figur",
                "2": "Bezeichnung",
                "3": "Epitheta",
            }[column_choice]

            filename = f"wordlist_{column_input}_{book_name}.csv".replace(" ", "_")
            output_path = os.path.join(output_dir, filename)
            generate_wordlist_by_column(column_input, json_path, output_path)

        elif menu_choice == "2":
            figure = ask_valid_figure_name(json_path)
            if figure is None:
                return
            filename = f"wordlist_naming_variants_{figure}.csv".replace(" ", "_")
            output_path = os.path.join(output_dir, filename)
            generate_naming_variants_for_figure(figure, json_path, output_path)

        elif menu_choice == "3":
            figure = ask_valid_figure_name(json_path)
            if figure is None:
                return
            filename = f"wordlist_epithets_{figure}.csv".replace(" ", "_")
            output_path = os.path.join(output_dir, filename)
            generate_epithets_for_figure(figure, json_path, output_path)

        elif menu_choice == "4":
            figure = ask_valid_figure_name(json_path)
            if figure is None:
                return
            filename = f"wordlist_combined_{figure}.csv".replace(" ", "_")
            output_path = os.path.join(output_dir, filename)
            generate_combined_naming_variants_epithets(figure, json_path, output_path)

        elif menu_choice == "5":
            print("Returning to analysis menu.")
            return

def generate_wordlist_by_column(column_name: str, json_path: str, output_path: str):
    """
    Generate a frequency-based wordlist from a selected categorical column group.

    This function reads the categorization JSON file of a work and extracts
    all non-empty string values belonging to the specified logical column group.
    Depending on the column name, the function expands grouped fields
    (e.g., "Bezeichnung 1–4" or "Epitheta 1–5") and aggregates their values
    into a unified frequency list.

    The resulting token counts are sorted in descending order of frequency
    and written to a CSV file with the columns "Token" and "Count".

    Parameters:
        column_name (str):
            Logical column group to analyze. Supported special groups are
            "Bezeichnung" and "Epitheta", which are internally expanded to
            their numbered column variants. Any other value is treated as
            a direct JSON key.

        json_path (str):
            Path to the categorization JSON file.

        output_path (str):
            Destination path for the generated CSV file.

    Returns:
        None
    """
    entries = safe_read_json(json_path, default=[])

    normalized_column = column_name.strip().lower()

    if normalized_column == "bezeichnung":
        columns = [f"Bezeichnung {i}" for i in range(1, 5)]
    elif normalized_column == "epitheta":
        columns = [f"Epitheta {i}" for i in range(1, 6)]
    else:
        columns = [column_name]

    all_values = []
    for entry in entries:
        for col in columns:
            value = entry.get(col)
            if isinstance(value, str) and value.strip():
                all_values.append(value.strip())

    counts = Counter(all_values)
    most_common = counts.most_common()

    write_csv_table(
        output_path,
        header=["Token", "Count"],
        rows=[(val, count) for val, count in most_common],
    )

    print(f"Wordlist written to: {output_path}")

def generate_naming_variants_for_figure(named_figure: str, json_path: str, output_path: str):
    """
    Generate a frequency-based wordlist of naming variants for a specific figure.

    This function reads the categorization JSON file of a work, filters all
    entries in which the field "Benannte Figur" matches the given figure,
    and extracts all non-empty values from the grouped fields
    "Bezeichnung 1–4".

    The extracted naming variants are aggregated into a frequency list,
    sorted in descending order by occurrence, and written to a CSV file
    with the columns "Token" and "Count".

    Parameters:
        named_figure (str):
            Name of the figure to analyze. The value is expected to be
            validated prior to calling this function.

        json_path (str):
            Path to the categorization JSON file.

        output_path (str):
            Destination path for the generated CSV file.

    Returns:
        None
    """
    entries = safe_read_json(json_path, default=[])

    filtered = [e for e in entries if e.get("Benannte Figur") == named_figure]
    values = []
    for entry in filtered:
        for i in range(1, 5):
            val = entry.get(f"Bezeichnung {i}")
            if isinstance(val, str) and val.strip():
                values.append(val.strip())

    counts = Counter(values)
    most_common = counts.most_common()

    write_csv_table(
        output_path,
        header=["Token", "Count"],
        rows=[(val, count) for val, count in most_common],
    )

    print(f"Naming variants for '{named_figure}' exported to: {output_path}")

def generate_epithets_for_figure(named_figure: str, json_path: str, output_path: str):
    """
    Generate a frequency-based wordlist of epithets for a specific figure.

    This function reads the categorization JSON file of a work, filters all
    entries in which the field "Benannte Figur" matches the given figure,
    and extracts all non-empty values from the grouped fields
    "Epitheta 1–5".

    The extracted epithets are aggregated into a frequency list,
    sorted in descending order by occurrence, and written to a CSV file
    with the columns "Token" and "Count".

    Parameters:
        named_figure (str):
            Name of the figure to analyze. The value is expected to be
            validated prior to calling this function.

        json_path (str):
            Path to the categorization JSON file.

        output_path (str):
            Destination path for the generated CSV file.

    Returns:
        None
    """
    entries = safe_read_json(json_path, default=[])

    filtered = [e for e in entries if e.get("Benannte Figur") == named_figure]
    values = []
    for entry in filtered:
        for i in range(1, 6):
            val = entry.get(f"Epitheta {i}")
            if isinstance(val, str) and val.strip():
                values.append(val.strip())

    counts = Counter(values)
    most_common = counts.most_common()

    write_csv_table(
        output_path,
        header=["Token", "Count"],
        rows=[(val, count) for val, count in most_common],
    )

    print(f"Epithets for '{named_figure}' exported to: {output_path}")

def generate_combined_naming_variants_epithets(named_figure: str, json_path: str, output_path: str):
    """
    Generate a combined frequency-based wordlist of naming variants and epithets
    for a specific figure.

    This function reads the categorization JSON file of a work, filters all
    entries in which the field "Benannte Figur" matches the given figure,
    and extracts all non-empty values from the grouped fields
    "Bezeichnung 1–4" and "Epitheta 1–5".

    All extracted tokens (naming variants and epithets) are aggregated into
    a single frequency list without distinguishing their source category.
    Identical string values occurring in either group are counted jointly.

    The resulting token frequencies are sorted in descending order and
    written to a CSV file with the columns "Token" and "Count".

    Parameters:
        named_figure (str):
            Name of the figure to analyze. The value is expected to be
            validated prior to calling this function.

        json_path (str):
            Path to the categorization JSON file.

        output_path (str):
            Destination path for the generated CSV file.

    Returns:
        None
    """
    entries = safe_read_json(json_path, default=[])

    filtered = [e for e in entries if e.get("Benannte Figur") == named_figure]
    values = []
    for entry in filtered:
        for i in range(1, 5):
            val = entry.get(f"Bezeichnung {i}")
            if isinstance(val, str) and val.strip():
                values.append(val.strip())
        for i in range(1, 6):
            val = entry.get(f"Epitheta {i}")
            if isinstance(val, str) and val.strip():
                values.append(val.strip())

    counts = Counter(values)
    most_common = counts.most_common()

    write_csv_table(
        output_path,
        header=["Token", "Count"],
        rows=[(val, count) for val, count in most_common],
    )

    print(f"Naming variants and epithets for '{named_figure}' exported to: {output_path}")

# =============================================================================
# NAMING FIGURE ANALYSIS (menu + profile generators)
# =============================================================================

def run_naming_figure_analysis(_config_data, paths, data, book_name):
    """
    Interactive CLI dispatcher for the "Naming figure analysis" workflow.

    This function coordinates the selection and execution of analysis paths
    related to naming figures within a work. It loads naming data from the
    available sources (JSON with optional Excel fallback), validates user input,
    and delegates the selected task to the corresponding analysis function.

    Available analysis paths:
        1. Overview of naming figures (frequency of namers for a given figure)
        2. Naming profile by namer (figure–namer perspective)
        3. Named figure profile by lemma (lemma-specific analysis)

    The function performs no analytical computation itself. It is responsible
    for user interaction, validated selection, and controlled delegation to
    downstream analysis functions.

    Parameters:
        _config_data (dict):
            Configuration settings for the project session. Currently unused
            within this function but retained for interface consistency.

        paths (dict):
            Dictionary of resolved project paths. Must include access to
            the categorization JSON file.

        data (dict):
            Preloaded session data (e.g., TEI structures and Excel tables)
            required for source selection and fallback logic.

        book_name (str):
            Identifier of the active work. Used for contextual analysis
            and output labeling in downstream functions.

    Returns:
        None

    Behavior:
        - Loads naming sources using centralized selection logic.
        - Prompts the user for a valid target figure.
        - Displays an interactive sub-menu for selecting the analysis type.
        - Delegates execution to the respective analysis function.
        - Returns to the main analysis menu upon completion or invalid input.
    """
    # --- Load naming sources (JSON primary, Excel fallback if required) ---
    df_json, df_excel = load_naming_sources_with_excel_fallback(paths, data)

    # --- Request validated target figure (interactive resolution loop) ---
    target_figure = ask_valid_figure_name(paths["categorization_json"])
    if not target_figure:
        print("No figure name provided.")
        return

    # --- Sub-menu: select analysis mode for the chosen figure ---
    print(
        "\nWhich list should be output?\n"
        "[1] Overview of namers (naming figures)\n"
        "[2] Naming profile by namer\n"
        "[3] Named figure profile by lemma"
    )
    choice = ask_user_choice("> Select option:", ["1", "2", "3"])

    # ======================================================================
    # [1] Overview of naming figures (figure → namer frequency)
    # ======================================================================
    if choice == "1":
        selection = select_naming_data(book_name, df_json, df_excel)

        # Emit source-selection diagnostics (if any)
        for msg in selection.get("messages", []):
            print(msg)

        try:
            # Validate required structural features for this analysis mode
            check_naming_requirements(
                selection,
                require_target=True,
                require_namer=True,
                require_content=True,
                context="Overview",
            )
        except ValueError as e:
            print(str(e))
            return

        # Delegate analytical computation
        analyze_overview_of_naming_figures(
            book_name,
            selection["df"],
            selection["cols"],
            target_figure,
        )
        return

    # ======================================================================
    # [2] Naming profile by namer (interactive namer selection)
    # ======================================================================
    if choice == "2":

        try:
            selection = select_naming_data(book_name, df_json, df_excel)

            for msg in selection.get("messages", []):
                print(msg)

            df = selection["df"]
            cols = selection["cols"]

            tcol = cols["target"]
            ncol = cols["namer"]

            # Restrict dataset to selected target figure
            df_sub = df.loc[df[tcol] == target_figure]
            if df_sub.empty:
                print(f"No entries found for named figure: {target_figure}")
                return

            # Compute frequency distribution of namers
            freq = df_sub[ncol].value_counts().to_dict()
            print("\nList of namers (sorted by frequency):")

            sorted_namers = sorted(freq.items(), key=lambda t: (-t[1], t[0]))

            # Display enumerated namer list for interactive selection
            for i, (name, count) in enumerate(sorted_namers, start=1):
                print(f"{i} {name} ({count})")

        except Exception as e:
            print(f"(An error occurred while preparing the selection list: {e})")
            return

        # --- Interactive namer selection ---
        if not freq:
            print("No namers available.")
            return

        print("\nSelect a namer by number (press Enter to return to analysis menu).")
        valid_choices = [str(i) for i in range(1, len(sorted_namers) + 1)]

        raw = input("> ").strip()
        if not raw:
            return

        while raw not in valid_choices:
            print("Invalid selection. Please enter one of the listed numbers.")
            raw = input("> ").strip()
            if not raw:
                return

        selected_namer = sorted_namers[int(raw) - 1][0]

        # Delegate analytical computation
        analyze_naming_profile_by_figure(
            book_name, df_json, df_excel, target_figure, selected_namer
        )
        return

    # ======================================================================
    # [3] Named figure profile by lemma
    # ======================================================================
    if choice == "3":
        # Interactive lemma input (empty input returns to analysis menu)
        query_lemma = input("\nPlease enter the lemma: ").strip()
        if not query_lemma:
            print("No lemma provided. Returning to analysis menu.")
            return

        # Delegate analytical computation
        analyze_figure_profile_by_lemma(
            book_name, df_json, df_excel, target_figure, query_lemma
        )
        return

    # --- Fallback safeguard (should not be reachable with validated input) ---
    print("Invalid choice. Returning to main menu.")

def analyze_overview_of_naming_figures(book_name, df, cols, target_figure):
    """
    Generate an overview table of naming activity for a selected figure.

    This function analyzes how often a given target figure is mentioned by
    individual namers ("Nennende Figur") and determines how many of those
    mentions qualify as name-based mentions according to the project’s
    matching logic.

    Rules applied:
        - Each row in the dataset counts as exactly one total mention,
          regardless of how many designation or epithet fields are filled.
        - Name-based mentions are detected via `match_name_to_lemma(...)`.
        - The data is expected to contain normalized, canonical forms of
          figure names. After user validation of the target figure, no
          additional similarity heuristics are applied beyond the defined
          matching function.
        - If no name-based mentions are detected, the function optionally
          suggests the closest lemma form present in the data and offers
          a single confirmation dialog.
            • If confirmed ("y"), name-based mentions are recounted using
              the confirmed form.
            • If rejected ("n") or no suitable suggestion exists, a reduced
              table (without name-based statistics) is produced.

    Output:
        A CSV file written to:
            data/<book_name>/analysis/<target_figure>_naming_overview.csv

        Depending on the mode:
            - Full table:
                Namer | Total mentions | Name mentions | Share of name mentions (%)
            - Reduced table:
                Namer | Total mentions

    Parameters:
        book_name (str):
            Identifier of the active work.

        df (pandas.DataFrame):
            Unified naming dataset (JSON or Excel fallback), containing
            the relevant naming columns.

        cols (dict):
            Column mapping dictionary providing keys for:
                - "target": column name for "Benannte Figur"
                - "namer": column name for "Nennende Figur"
                - "naming_variant_cols": list of column names for "Bezeichnung 1–4"
                - "epithet_cols": list of column names for "Epitheta 1–5"

        target_figure (str):
            Canonical name of the selected figure, already validated via
            interactive resolution.

    Returns:
        None
    """
    # --- Resolve column mapping from unified naming dataset ---
    tcol = cols["target"]
    ncol = cols["namer"]
    nvcols = cols["naming_variant_cols"]
    ecols = cols["epithet_cols"]

    # --- Restrict dataset to the selected canonical target figure ---
    dft = df.loc[df[tcol] == target_figure]

    # --- Aggregation containers ---
    # counts_total: all mentions per namer (1 row = 1 mention)
    # counts_name: name-based mentions per namer
    counts_total = {}
    counts_name = {}

    # Optional alias hook (currently empty; retained for interface consistency)
    aliases = []

    # Inventory of all designation/epithet lemmas
    # Used only if close-match suggestion becomes necessary
    lemmas_all = set()

    # ======================================================================
    # Pass 1: Aggregate totals and detect name-based mentions
    # ======================================================================
    for _, row in dft.iterrows():
        namer = row.get(ncol)
        if not isinstance(namer, str) or namer.strip() == "":
            continue

        # Each row counts exactly once toward total mentions
        counts_total[namer] = counts_total.get(namer, 0) + 1

        # Collect naming variant + epithet values present in this row
        lemmas = []
        for c in nvcols:
            val = row.get(c)
            if isinstance(val, str) and val.strip() != "":
                lemmas.append(val)
                lemmas_all.add(val)

        for c in ecols:
            val = row.get(c)
            if isinstance(val, str) and val.strip() != "":
                lemmas.append(val)
                lemmas_all.add(val)

        # Determine whether this row qualifies as name-based
        if any(match_name_to_lemma(target_figure, lm, aliases=aliases) for lm in lemmas):
            counts_name[namer] = counts_name.get(namer, 0) + 1

    # ======================================================================
    # Optional fallback: no name-based hits detected
    # ======================================================================
    name_hits_sum = sum(counts_name.values())
    reduced_mode = False
    suggested = None

    if name_hits_sum == 0:
        # Attempt to propose the closest lemma form present in the data
        if lemmas_all:
            lowered = {lm.lower(): lm for lm in lemmas_all}
            best = difflib.get_close_matches(
                target_figure.lower(),
                list(lowered.keys()),
                n=1,
                cutoff=0.6,
            )
            if best:
                suggested = lowered[best[0]]

        if suggested:
            # Single interactive confirmation dialog
            print(
                f'{target_figure} could not be found in the '
                '"Bezeichnung" or "Epitheta" columns.'
            )
            yn = ask_user_choice(
                f'Could "{suggested}" be a variant of the name? (y/n)',
                ["y", "n"]
            )

            if yn == "y":
                # Recount name-based mentions using strict equality
                counts_name = {}

                for _, row in dft.iterrows():
                    namer = row.get(ncol)
                    if not isinstance(namer, str) or namer.strip() == "":
                        continue

                    hit = False

                    for c in nvcols:
                        v = row.get(c)
                        if isinstance(v, str) and v.lower().strip() == suggested.lower().strip():
                            hit = True
                            break

                    if not hit:
                        for c in ecols:
                            v = row.get(c)
                            if isinstance(v, str) and v.lower().strip() == suggested.lower().strip():
                                hit = True
                                break

                    if hit:
                        counts_name[namer] = counts_name.get(namer, 0) + 1
            else:
                # User rejected suggestion → reduced output mode
                reduced_mode = True
        else:
            # No reasonable suggestion available → reduced output mode
            reduced_mode = True

    # ======================================================================
    # Output generation
    # ======================================================================
    os.makedirs(os.path.join("data", book_name, "analysis"), exist_ok=True)
    out_path = os.path.join(
        "data",
        book_name,
        "analysis",
        f"{target_figure}_naming_overview.csv",
    )

    if reduced_mode:
        # Output without name-based statistics
        write_csv_table(
            out_path,
            header=["Namer", "Total mentions"],
            rows=[
                (namer, total)
                for namer, total in sorted(
                    counts_total.items(),
                    key=lambda t: (-t[1], t[0]),
                )
            ],
        )
        print(f"Naming figure overview written to: {out_path}")
        return

    # Full output including percentage share
    rows = []
    for namer, total in counts_total.items():
        nhits = counts_name.get(namer, 0)
        pct = int(round((nhits / total) * 100)) if total > 0 else 0
        rows.append((namer, total, nhits, pct))

    rows.sort(key=lambda t: (-t[1], t[0]))

    write_csv_table(
        out_path,
        header=[
            "Namer",
            "Total mentions",
            "Name mentions",
            "Share of name mentions (%)",
        ],
        rows=[
            (namer, total, nhits, pct)
            for namer, total, nhits, pct in rows
        ],
    )

    print(f"Namer overview for '{target_figure}' exported to: {out_path}")

def analyze_naming_profile_by_figure(book_name, df_json, df_excel, target_figure, selected_namer):
    """
    Generate a naming variant profile for a specific naming figure
    with respect to a selected target figure.

    This function extracts all naming variants used by a given namer
    ("Nennende Figur") when referring to a selected target figure
    ("Benannte Figur") and exports them as a CSV list.

    Data handling logic:
        - The naming dataset is loaded via `prepare_naming_data(...)`,
          which unifies JSON and Excel sources (with fallback logic).
        - Rows are filtered where:
              target == target_figure
              namer  == selected_namer
        - If an unnumbered column "Bezeichnung" exists
          (raw, non-lemmatized form), only this column is used.
        - Otherwise, numbered naming variant columns ("Bezeichnung 1–n")
          and all epithet columns ("Epitheta 1–n") are included.
        - If a verse column is available, its value is included
          alongside each extracted token.

    Output:
        A CSV file written to:
            data/<book_name>/analysis/
            <target_figure>_naming_profile_by_<selected_namer>.csv

        The file contains the columns:
            Verse | Token

        One row is written per naming variant/epithet occurrence.

    Parameters:
        book_name (str):
            Identifier of the active work.

        df_json (pandas.DataFrame | None):
            JSON-based naming dataset (can be None or empty).

        df_excel (pandas.DataFrame | None):
            Excel-based fallback dataset (can be None or empty).

        target_figure (str):
            Canonical name of the selected target figure.

        selected_namer (str):
            Canonical name of the selected naming figure.

    Returns:
        None
    """
    # --- Load unified naming dataset (JSON primary, Excel fallback if needed) ---
    source, df, cols = prepare_naming_data(book_name, df_json, df_excel)

    # --- Resolve relevant column mappings ---
    tcol = cols["target"]
    ncol = cols["namer"]
    nvcols = cols["naming_variant_cols"]
    ecols = cols["epithet_cols"]
    vcol  = cols.get("verse_col")
    has_raw = bool(cols.get("has_unnumbered_naming_variant"))

    # --- Restrict dataset to selected target figure and naming figure ---
    dff = df.loc[(df[tcol] == target_figure) & (df[ncol] == selected_namer)]

    rows = []

    # ======================================================================
    # Extract naming variants / epithets per matching row
    # ======================================================================
    for _, row in dff.iterrows():

        # --- Optional verse extraction (defensive access) ---
        verse = ""
        if isinstance(vcol, str):
            val = row.get(vcol)
            if isinstance(val, str):
                verse = val
            elif val is not None:
                verse = str(val)

        # --- Raw naming variant mode ---
        # If an unnumbered "Bezeichnung" column exists, use only this field.
        if has_raw:
            raw_col = None
            for c in nvcols:
                cname = c.strip().lower()
                if cname == "bezeichnung":
                    raw_col = c
                    break

            if raw_col:
                val = row.get(raw_col)
                if isinstance(val, str) and val.strip() != "":
                    rows.append((verse, val))
            continue  # Skip numbered/epithet extraction in raw mode

        # --- Lemmatized mode ---
        # Collect numbered naming variants (excluding unnumbered column)
        for c in nvcols:
            cname = c.strip().lower()
            if cname == "bezeichnung":
                continue
            val = row.get(c)
            if isinstance(val, str) and val.strip() != "":
                rows.append((verse, val))

        # Collect epithets
        for c in ecols:
            val = row.get(c)
            if isinstance(val, str) and val.strip() != "":
                rows.append((verse, val))


    # ======================================================================
    # Output generation
    # ======================================================================
    os.makedirs(os.path.join("data", book_name, "analysis"), exist_ok=True)

    out_path = os.path.join(
        "data",
        book_name,
        "analysis",
        f"{target_figure}_naming_profile_by_{selected_namer}.csv",
    )

    # Write CSV with consistent header schema
    write_csv_table(
        out_path,
        header=["Verse", "Token"],
        rows=[(serialize_verse_value(verse), bez) for verse, bez in rows],
    )

    print(
        f"Naming profile by namer '{selected_namer}' "
        f"for '{target_figure}' exported to: {out_path}"
    )

def analyze_figure_profile_by_lemma(book_name, df_json, df_excel, target_figure, query_lemma):
    """
    Analyze which naming figures use a given lemma for a specific target figure.

    The function restricts the dataset to rows where the target column matches
    the provided target_figure exactly. Within this subset, it checks whether
    query_lemma (case-insensitive, stripped) occurs as an exact match in any
    numbered naming_variant_cols or epithet_cols column. The unnumbered raw
    column "Bezeichnung" is intentionally excluded.

    For each row containing at least one match, the corresponding namer
    (cols["namer"]) is counted once per row.

    Output:
        data/{book_name}/analysis/
        {target_figure}_figure_profile_by_{query_lemma}.csv

    Export format:
        Header: ["Namer", "Count"]
        Sorted by descending count, then alphabetically by namer.

    If query_lemma is empty or no matches are found, no file is written and
    a diagnostic message is printed.

    Notes (BETA semantics):
        - Matching is strict (strip + lower); no additional normalization
          or similarity heuristics are applied.
        - The function relies on prepare_naming_data(...) for column mapping
          and source selection.
    """
    # unify data
    source, df, cols = prepare_naming_data(book_name, df_json, df_excel)
    tcol = cols["target"]
    ncol = cols["namer"]
    nvcols = [c for c in cols["naming_variant_cols"] if str(c).strip().lower() != "bezeichnung"]
    ecols = cols["epithet_cols"]

    dft = df.loc[df[tcol] == target_figure]

    q = (query_lemma or "").strip().lower()
    if q == "":
        print("No results found for: (empty lemma)")
        return

    counts = {}

    for _, row in dft.iterrows():
        namer = row.get(ncol)
        if not isinstance(namer, str) or namer.strip() == "":
            continue

        hit = False

        for c in nvcols:
            val = row.get(c)
            if isinstance(val, str) and val.strip().lower() == q:
                hit = True
                break

        if not hit:
            for c in ecols:
                val = row.get(c)
                if isinstance(val, str) and val.strip().lower() == q:
                    hit = True
                    break

        if hit:
            counts[namer] = counts.get(namer, 0) + 1

    if not counts:
        print(f"No results found for: {query_lemma}")
        return

    os.makedirs(os.path.join("data", book_name, "analysis"), exist_ok=True)
    out_path = os.path.join(
        "data", book_name, "analysis", f"{target_figure}_figure_profile_by_{query_lemma}.csv"
    )

    # write CSV
    write_csv_table(
        out_path,
        header=["Namer", "Count"],
        rows=[
            (namer, cnt)
            for namer, cnt in sorted(counts.items(), key=lambda t: (-t[1], t[0]))
        ],
    )

    print(f"Namer profile for lemma '{query_lemma}' in '{target_figure}' exported to: {out_path}")

# =============================================================================
# KEYWORD ANALYSIS (menu + G² computation)
# =============================================================================

def run_keyword_menu(_config_data, paths, _data, book_name):
    """
    Interactive CLI wrapper for running a keyword analysis on a given book.

    The user can choose between:
        - analyzing the whole book (with a reference corpus),
        - analyzing a specific target figure within the book.

    For whole-book analysis, available reference books are derived from the
    data directory (excluding the current book). If none are detected, the
    user may enter reference book names manually.

    The user then selects the comparison unit:
        - "bezeichnung"  (Naming variants),
        - "epitheta" (Epithtets),
        - "combined".

    A significance threshold (Log-Likelihood G²) can be provided; if omitted
    or invalid, the default value 3.84 is used.

    The function constructs an output filename of the form:
        keywords_<Unit>_<Target>_<Book>.csv
    and delegates computation to generate_keywords(...).

    Control flow:
        - Runs inside a while-loop to allow repeated analyses.
        - Returns to the analysis menu when the user selects 'n'.

    Parameters:
        _config_data (dict): Configuration container (kept for interface consistency).
        paths (dict): Path dictionary; must contain key "categorization_json".
        _data (dict): Loaded TEI/Excel data (kept for interface consistency).
        book_name (str): Identifier of the current book (used for output and context).

    Returns:
        None
    """
    while True:
        target_json = paths.get("categorization_json")
        if not target_json:
            print("Missing required path: 'categorization_json'. Cannot run keyword analysis.")
            return None

        # Ensure analysis output directory exists for the current book
        output_dir = os.path.join("data", book_name, "analysis")
        os.makedirs(output_dir, exist_ok=True)

        # Select analysis scope: whole book or specific figure
        print("\nDo you want to analyze the whole book or a specific figure?")
        print("[1] Whole book")
        print("[2] Specific figure")

        target_choice = ask_user_choice("> ", ["1", "2"])

        reference_books = None  # default; only used for whole-book analysis

        if target_choice == "2":
            # Resolve and validate target figure against categorization JSON
            target = ask_valid_figure_name(target_json)
            if target is None:
                return None
            target_type = "figure"

        else:
            # Whole-book analysis; current book acts as target corpus
            target = book_name
            target_type = "whole_book"

            # Derive available reference books from data directory (excluding current book)
            available = list_available_reference_books(exclude=book_name)

            if not available:
                # Fallback: manual reference corpus specification
                print("No reference books could be derived from the data directory.")
                print("Please enter the names of the books to include in the reference corpus (comma-separated):")
                references = input("> ").strip()
                reference_books = [r.strip() for r in references.split(",") if r.strip()]
            else:
                # Interactive numeric selection of reference books
                print("\nAvailable reference books (derived from data folders):")
                for i, b in enumerate(available, start=1):
                    print(f"[{i}] {b}")

                print("\nSelect reference books by number (e.g., 1-3,5).")
                print("Press Enter to use ALL listed reference books.")
                raw = input("> ").strip()

                if raw == "":
                    reference_books = available
                else:
                    indices = parse_token_selection(raw, len(available))
                    while not indices:
                        print("Invalid input. Please enter numbers/ranges like 1-3,5 or press Enter for all.")
                        raw = input("> ").strip()
                        if raw == "":
                            indices = list(range(1, len(available) + 1))
                            break
                        indices = parse_token_selection(raw, len(available))

                    reference_books = [available[i - 1] for i in indices]

        # Select comparison unit for keyword extraction
        print("\nWhat should be the unit of comparison?")
        print("[1] Naming variants (Bezeichnungen)")
        print("[2] Epithets (Epitheta)")
        print("[3] Combined")

        unit_choice = ask_user_choice("> ", ["1", "2", "3"])
        unit = {
            "1": "bezeichnung",
            "2": "epitheta",
            "3": "combined"
        }[unit_choice]

        # Read significance threshold (Log-Likelihood G²); fallback to default if empty/invalid
        print("\nType in significance threshold (Log-Likelihood G²), for default = 3.84 press 'Enter':")
        threshold_input = input("> ").strip()
        try:
            threshold = float(threshold_input) if threshold_input else 3.84
        except ValueError:
            print("Invalid input – using default threshold 3.84")
            threshold = 3.84

        # Build output filename (normalized target label)
        target_label = target.replace(" ", "_")
        unit_file = {"bezeichnung": "Bezeichnung", "epitheta": "Epitheta", "combined": "combined"}[unit]
        output_file = f"keywords_{unit_file}_{target_label}_{book_name}.csv"
        output_path = os.path.join(output_dir, output_file)

        # Dispatch to core keyword analysis function
        target_figure = target if target_type == "figure" else None
        ref_books = None if target_type == "figure" else reference_books

        generate_keywords(
            target_figure=target_figure,
            reference_books=ref_books,
            unit=unit,
            threshold=threshold,
            target_json=target_json,
            output_path=output_path
        )

        print(f"Keyword analysis exported to: {output_path}")

        # Offer repeated execution within same session
        print("\nDo you want to run another keyword analysis? [y/n]")
        again = ask_user_choice("> ", ["y", "n"])
        if again == "y":
            continue
        else:
            print("Returning to analysis menu.")
            return None

def generate_keywords(
    target_figure: str | None,
    reference_books: list[str] | None,
    unit: str,
    threshold: float,
    target_json: str,
    output_path: str
):
    """
    Compute keyword scores (G² log-likelihood) for a target figure or a whole book.

    The function compares token frequencies of the target corpus against a
    reference corpus and calculates keyness values using the log-likelihood
    (G²) statistic. Only tokens whose keyness value meets or exceeds the
    specified threshold are retained.

    Target corpus:
        - If target_figure is provided, only entries whose
          "Benannte Figur" matches the figure are considered.
        - If target_figure is None, the entire book is used as target corpus.

    Reference corpus:
        - If reference_books is provided, categorized JSON files of the
          listed books are loaded and combined.
        - Otherwise, all entries of the current book except those belonging
          to the target_figure are used as reference corpus.

    Token extraction:
        Tokens are derived via extract_tokens(...) according to the selected unit:
            - "bezeichnung"
            - "epitheta"
            - "combined"

    Output:
        A CSV file written to output_path with the columns:
            Token | Target count | Reference count | Keyness | Polarity

        Polarity is defined as:
            - "positive" if the token is more frequent in the target corpus
            - "negative" if more frequent in the reference corpus
            - "neutral" if frequencies are equal

    Edge cases (BETA semantics):
        - If both corpora are empty, no file is written.
        - If no token reaches the specified threshold, no file is written.
        - No additional normalization or similarity heuristics are applied;
          matching and counting rely on preprocessed categorized data.

    Parameters:
        target_figure (str | None): Target figure for analysis (None = whole book).
        reference_books (list[str] | None): Books forming the reference corpus.
        unit (str): Token unit ("bezeichnung", "epitheta", "combined").
        threshold (float): Minimum G² value required for inclusion.
        target_json (str): Path to the categorized JSON file of the current book.
        output_path (str): Destination path for the CSV export.

    Returns:
        None
    """
    all_entries = safe_read_json(target_json, default=[])
    target_entries = all_entries

    # Restrict target corpus to a specific figure if requested
    if target_figure:
        target_entries = [e for e in target_entries if e.get("Benannte Figur") == target_figure]

    # Extract tokens from target corpus according to selected unit
    target_tokens = extract_tokens(target_entries, unit)

    # Initialize reference corpus container
    reference_entries = []

    if reference_books:
        # Load and merge categorized entries from selected reference books
        for book in reference_books:
            path = os.path.join("data", book, f"categorization_{book}.json")
            reference_entries.extend(safe_read_json(path, default=[]))
    else:
        # Fallback: use entries from current book excluding the target figure
        reference_entries = [
            e for e in all_entries
            if not target_figure or e.get("Benannte Figur") != target_figure
        ]

    # Extract tokens from reference corpus
    reference_tokens = extract_tokens(reference_entries, unit)

    # Count token frequencies in both corpora
    target_counts = Counter(target_tokens)
    reference_counts = Counter(reference_tokens)

    results = []
    total_target = sum(target_counts.values())
    total_ref = sum(reference_counts.values())

    # Guard against empty corpora (prevents division by zero)
    if total_target + total_ref == 0:
        print("No results found (empty target and reference corpora).")
        return

    # Compute G² keyness values for each token in the target corpus
    for token, count_t in target_counts.items():
        count_r = reference_counts.get(token, 0)

        # Skip tokens that are absent in both corpora
        if count_t + count_r == 0:
            continue

        # Estimate pooled probability and expected frequencies
        p = (count_t + count_r) / (total_target + total_ref)
        expected_t = p * total_target
        expected_r = p * total_ref

        # Compute log components safely (avoid log(0))
        log_t = count_t * math.log2(count_t / expected_t) if count_t > 0 and expected_t > 0 else 0
        log_r = count_r * math.log2(count_r / expected_r) if count_r > 0 and expected_r > 0 else 0

        keyness = 2 * (log_t + log_r)

        # Apply significance threshold
        if keyness >= threshold:
            if count_t > count_r:
                typ = "positive"
            elif count_r > count_t:
                typ = "negative"
            else:
                typ = "neutral"

            results.append((token, count_t, count_r, round(keyness, 2), typ))

    # Sort results by descending keyness, then alphabetically by token
    results.sort(key=lambda x: (-x[3], x[0]))

    # Do not write empty result files
    if not results:
        print(f"No results found for threshold >= {threshold}.")
        return

    # Export keyword table as CSV
    write_csv_table(
        output_path,
        header=["Token", "Target count", "Reference count", "Keyness", "Polarity"],
        rows=results,
    )

# =============================================================================
# COLLOCATION ANALYSIS (menu + KWIC generation)
# =============================================================================

def run_collocation_menu(config_data, paths, data, book_name):
    """
    Interactive CLI wrapper for collocation analysis (KWIC-based).

    The user selects whether to analyze the whole book or restrict the analysis
    to a specific target figure. A lemma is then entered as search term.
    Results can either be displayed in the console or exported as a CSV file
    in KWIC format.

    Control flow:
        - The target figure (if selected) is validated against the
          categorization JSON.
        - Lemma input is enforced to be non-empty.
        - Output mode (console or CSV) is selected interactively.
        - If CSV export fails due to an open file (PermissionError),
          the user is prompted to close the file and retry.

    Output:
        - Console mode: results are printed directly.
        - CSV mode: file is written to
              data/{book_name}/analysis/
              collocations_<figure_or_whole_book>_<lemma>_<book_name>.csv

    Parameters:
        config_data (dict): Global configuration passed to the collocation engine.
        paths (dict): Path dictionary; must contain key "categorization_json".
        data (dict): Loaded TEI and Excel data used for token/context extraction.
        book_name (str): Identifier of the current book (used for context and filenames).

    Returns:
        None
    """
    # Resolve categorization JSON path required for figure validation
    categorization_path = paths.get("categorization_json")
    if not categorization_path:
        print("Missing required path: 'categorization_json'. Cannot run collocation analysis.")
        return

    # Select analysis scope: whole book or restricted to a specific figure
    print("\nDo you want to analyze the whole book or only a specific figure?")
    print("[1] Whole book")
    print("[2] Specific figure")

    target_mode = ask_user_choice("> ", ["1", "2"])
    only_figure = None

    if target_mode == "2":
        # Validate figure name against categorized entries
        only_figure = ask_valid_figure_name(categorization_path)
        if only_figure is None:
            return

    # Enforce non-empty lemma input
    while True:
        lemma_value = input("Please enter the lemma to search for (e.g., \"küene\"):\n> ").strip()
        if lemma_value:
            break
        print("Input cannot be empty. Please enter a lemma to search for.")

    # Select output mode and handle retry in case of open CSV file
    while True:
        print("\nWhere should the results be displayed?")
        print("[1] Console")
        print("[2] Save as CSV file")

        output_choice = ask_user_choice("> ", ["1", "2"])
        output_target = "console" if output_choice == "1" else "csv"

        if output_target == "csv":
            # Prepare output path for KWIC CSV export
            lemma_label = lemma_value.replace(" ", "_")
            fig_label = only_figure.replace(" ", "_") if only_figure else "whole_book"
            output_dir = os.path.join("data", book_name, "analysis")
            os.makedirs(output_dir, exist_ok=True)
            filename = f"collocations_{fig_label}_{lemma_label}_{book_name}.csv"
            output_path = os.path.join(output_dir, filename)
        else:
            output_path = None

        try:
            # Delegate computation to collocation engine
            generate_collocations(
                data=data,
                lemma_value=lemma_value,
                book_name=book_name,
                config_data=config_data,
                only_figure=only_figure,
                output_target=output_target,
                output_path=output_path
            )
            break  # Exit retry loop on success

        except PermissionError:
            # Handle open-file scenario for CSV export
            print("\nThe Excel file appears to be open.")
            print("Please close it and try again.")
            print("Returning to output choice...\n")

def generate_collocations(
    data: dict,
    lemma_value: str,
    book_name: str,
    config_data: dict,
    only_figure: str | None,
    output_target: str,
    output_path: str | None
):
    """
    Extract KWIC-style collocation contexts for a given lemma within a book.

    The function searches categorized entries of the current book for
    occurrences of the specified lemma in numbered naming_variant and
    epithet fields. If only_figure is provided, entries are restricted
    to that target figure.

    For each matching entry, the corresponding verse and collocation
    context are retrieved from the Excel sheet containing a
    "Kollokationen" column. If the Excel sheet cannot be loaded,
    a TEI-based fallback reconstruction is attempted.

    Lemma matching:
        - Matching is case-insensitive and whitespace-normalized.
        - Variants are expanded via lemma_normalization.json.
        - If no direct canonical entry exists, reverse lookup within
          the normalization mapping is performed.
        - KWIC highlighting is applied to all recognized variants.

    Output behavior:
        - If no matching collocations are found, a diagnostic message
          is printed and no output is produced.
        - In "console" mode, results are printed in formatted KWIC layout.
        - In "csv" mode, results are written to output_path with columns:
              Verse | Named figure | Left | Hit | Right

    Error handling (BETA semantics):
        - If the Excel collocation sheet is unavailable, a TEI fallback
          is used (if TEI data is present in data["xml"]).
        - If neither Excel nor TEI data are available, the function aborts.
        - No additional fuzzy matching or similarity heuristics are applied.

    Parameters:
        data (dict): Loaded data container (Excel/TEI).
        lemma_value (str): Lemma to search for.
        book_name (str): Current book identifier.
        config_data (dict): Configuration passed to Excel loader.
        only_figure (str | None): Optional figure restriction.
        output_target (str): "console" or "csv".
        output_path (str | None): Destination path for CSV export.

    Returns:
        None
    """
    # Load categorized entries of the current book
    json_path = os.path.join("data", book_name, f"categorization_{book_name}.json")
    entries = safe_read_json(json_path, default=[])

    # Load lemma normalization mapping (canonical → variants)
    lemma_map = cast(
        dict[str, list[str]],
        safe_read_json("data/lemma_normalization.json", default={})
    )

    # Restrict entries to a specific target figure if requested
    if only_figure:
        entries = [e for e in entries if e.get("Benannte Figur") == only_figure]

    # Attempt to load collocation sheet (Excel-based primary source)
    df = load_collocation_sheet(config_data, book_name)
    if df is None:
        print("Could not load the Excel sheet with 'Kollokationen'.")
        print("Falling back to TEI to reconstruct collocations.")
        xml_root = data.get("xml")
        if xml_root is None:
            print("No TEI source found in data['xml']. Cannot reconstruct collocations.")
            return
        df = build_fallback_collocation_df_from_tei(xml_root)

    results = []

    # Iterate through categorized entries and identify lemma matches
    for entry in entries:
        # Collect numbered naming_variant and epithet fields
        all_lemma_fields = [
            entry.get(f"Bezeichnung {i}") for i in range(1, 5)
        ] + [
            entry.get(f"Epitheta {i}") for i in range(1, 6)
        ]

        vers = entry.get("Vers")
        figur = entry.get("Benannte Figur")

        # Determine canonical textual anchor for matching Excel rows
        original_text = get_first_valid_text(
            entry.get("Erzähler"),
            entry.get("Bezeichnung"),
            entry.get("Eigennennung")
        )

        # Locate matching row in collocation DataFrame
        match = df[
            (df["Vers"] == vers) &
            (df["Benannte Figur"] == figur) &
            (df.apply(lambda r: get_first_valid_text(
                r.get("Erzähler"),
                r.get("Bezeichnung"),
                r.get("Eigennennung")
            ) == original_text, axis=1))
        ]

        if match.empty:
            continue

        collocation = match.iloc[0].get("Kollokationen")
        if not isinstance(collocation, str) or not collocation.strip():
            continue

        # Build normalized variant set for robust matching and highlighting
        norm_input: str = lemma_value.strip().lower()

        raw_variants: list[str] = lemma_map.get(lemma_value, [])
        variants_set: set[str] = {
            v.strip().lower()
            for v in raw_variants
            if isinstance(v, str) and v.strip()
        }
        variants_set.add(norm_input)

        # Reverse lookup if lemma_value is itself a variant
        if not raw_variants:
            for canon, vs in lemma_map.items():
                if any(isinstance(v, str) and v.strip().lower() == norm_input for v in (vs or [])):
                    variants_set.add(str(canon).strip().lower())
                    variants_set.update([
                        v.strip().lower()
                        for v in (vs or [])
                        if isinstance(v, str) and v.strip()
                    ])
                    break

        # Keep entry only if at least one naming field matches any variant
        if not any(
            isinstance(t, str) and t.strip().lower() in variants_set
            for t in all_lemma_fields
        ):
            continue

        # Generate KWIC segments (left context, hit, right context)
        left, hit, right = format_kwic(collocation, list(variants_set))
        results.append((vers, figur, left, hit, right))

    # Abort if no collocations were found
    if not results:
        print(f"No results found for: {lemma_value}")
        return

    # Output results according to selected target
    if output_target == "console":
        for _, _, left, hit, right in results:
            print(f"{left.strip():>40}  \033[1m\033[93m{hit}\033[0m  {right.strip():<40}")

    elif output_target == "csv" and output_path:
        write_csv_table(
            output_path,
            header=["Verse", "Named figure", "Left", "Hit", "Right"],
            rows=results,
        )
        print(f"Collocation results exported to: {output_path}")

# =============================================================================
# VISUALIZATION
# =============================================================================

# -----------------------------------------------------------------------------
# CLI menu (Visualization entry)
# -----------------------------------------------------------------------------

def run_visualization_menu(paths, book_name, data):
    """
    Interactive CLI menu for running visualization modules.

    Presents a selection of available visualizations for the current book
    and delegates execution to the corresponding visualization function.

    Available options:
        [1] Verse-based naming variant and epithet distribution
        [2] Intra-named-figure co-occurrence heatmap
        [3] Sunburst visualization
        [4] Return to analysis menu

    Control flow:
        - Runs inside a loop to allow multiple visualizations in one session.
        - Returns to the analysis menu when option [4] is selected.

    Parameters:
        paths (dict): Path configuration used by visualization modules.
        book_name (str): Identifier of the current book.
        data (dict): Loaded data container (passed to modules requiring TEI/JSON access).

    Returns:
        None
    """
    while True:
        print("\nWhich visualization do you want to run?")
        print("[1] Verse-based naming variants/ epithets distribution")
        print("[2] Intra-Named-Figure co-occurrence heatmap")
        print("[3] Sunburst visualization")
        print("[4] Returning to analysis menu")

        choice = ask_user_choice("> ", ["1", "2", "3", "4"])

        if choice == "1":
            visualize_verse_naming_distribution(paths, book_name)
        elif choice == "2":
            visualize_intra_figure_cooccurrence_heatmap(paths, book_name)
        elif choice == "3":
            run_sunburst_visualization(paths, book_name, data)

        elif choice == "4":
            print("Returning to analysis menu.")
            break

# =============================================================================
# VISUALIZATION – VERSE-BASED DISTRIBUTION (Scatter)
# =============================================================================

def visualize_verse_naming_distribution(paths, book_name):
    """
    Interactive CLI interface for verse-based token distribution visualization (Plotly).

    The user selects:
        - a target figure (validated against the categorization JSON),
        - a token unit to visualize ("Naming variants", "Epithets", or "Combined"),
        - specific tokens to include (numeric selection from frequency lists).

    Data source and preparation:
        - Categorized entries are loaded from paths["categorization_json"].
        - prepare_naming_data(...) is used to obtain a standardized DataFrame and
          a column mapping (naming_variant_cols, epithet_cols, target, verse_col).
        - The function normalizes the target and verse columns to the stable names
          "Benannte Figur" and "Vers" for downstream processing.

    Plot:
        - Tokens are displayed as y-axis categories and verse numbers as x-axis values.
        - One scatter trace per token is created; in "Combined" mode, tokens are
          color-coded by category (Naming variants vs Epithets) with a compact legend.
        - A dropdown menu allows interactive Show-N filtering (14, 28, 42, … up to max),
          updating trace visibility, y-axis tick labels, and marker/tick sizing.

    Output:
        - An interactive HTML file is exported via export_visualization_output(...).
        - The filename encodes unit and figure name (viz_<unit>_<figure>.html).

    Guards and BETA semantics:
        - Aborts with a diagnostic message if categorization data is missing,
          if the selected figure has no entries, or if required token columns
          are unavailable for the chosen unit.
        - The function relies on canonical/normalized categorized data and does not
          apply additional fuzzy matching or similarity heuristics.

    Parameters:
        paths (dict): Path dictionary; must include key "categorization_json".
        book_name (str): Identifier of the current book (used for output context).

    Returns:
        None
    """
    categorization_path = paths.get("categorization_json")
    if not categorization_path:
        print("Missing required path: 'categorization_json'. Cannot run this visualization.")
        return

    # ======================================================================
    # [1] Load and normalize categorized data
    # ======================================================================

    entries = safe_read_json(categorization_path, default=[])
    if not entries:
        print("No categorization data available.")
        return

    # Normalize via shared source-selection helper (JSON-only for visualization)
    df_json = pd.DataFrame(entries)
    _, df, cols = prepare_naming_data(book_name, df_json, None)

    # Harmonize core column names to stable downstream identifiers
    tcol = cols.get("target")
    vcol = cols.get("verse_col")
    rename_map = {}
    if tcol and tcol != "Benannte Figur":
        rename_map[tcol] = "Benannte Figur"
    if vcol and vcol != "Vers":
        rename_map[vcol] = "Vers"
    if rename_map:
        df = df.rename(columns=rename_map)

    # ======================================================================
    # [2] User input: figure and token unit
    # ======================================================================
    figure_name = ask_valid_figure_name(categorization_path)
    if figure_name is None:
        return

    print("\nWhat should be visualized?")
    print("[1] Naming variants")
    print("[2] Epithets")
    print("[3] Combined")
    variant_type = ask_user_choice("> ", ["1", "2", "3"])

    variant_label = {
        "1": "Naming variants",
        "2": "Epithets",
        "3": "Naming variants & epithets"
    }[variant_type]

    # ======================================================================
    # [3] Data preparation: restrict to figure and collect token columns
    # ======================================================================

    df_figure = df[df["Benannte Figur"] == figure_name].copy()
    if df_figure.empty:
        print(f"No entries found for figure: {figure_name}")
        return

    # Retrieve token columns from shared column mapping
    naming_cols = [
        c for c in cols.get("naming_variant_cols", [])
        if str(c).strip().lower() != "bezeichnung"
    ]
    epithet_cols = cols.get("epithet_cols", [])

    # Determine active token columns depending on selected unit
    if variant_type == "1":
        selected_cols = naming_cols
    elif variant_type == "2":
        selected_cols = epithet_cols
    else:
        selected_cols = naming_cols + epithet_cols

    # Guard: abort early if no relevant token columns are available
    if variant_type == "1" and not naming_cols:
        print("No naming variant columns available for this book.")
        return
    if variant_type == "2" and not epithet_cols:
        print("No epithet columns available for this book.")
        return
    if variant_type == "3" and not selected_cols:
        print("No naming variant or epithet columns available for this book.")
        return

    # Validate presence of required structural columns
    required_base_cols = ["Benannte Figur", "Vers"]
    required_token_cols = selected_cols
    missing_cols = [c for c in (required_base_cols + required_token_cols) if c not in df.columns]
    if missing_cols:
        print("Missing required columns for this visualization:")
        for c in missing_cols:
            print(f"   - {c}")
        return

    # Optional contextual metadata (included if present)
    meta_cols = ["Nennende Figur", "Erzähler", "Eigennennung"]

    # ======================================================================
    # [4] Build long-format token table
    # ======================================================================
    all_entries = []
    for col in selected_cols:
        keep_cols = ["Vers", col] + [c for c in meta_cols if c in df_figure.columns]
        temp = df_figure[keep_cols].dropna(subset=["Vers", col]).rename(columns={col: "Token"})
        all_entries.append(temp)

    df_combined = pd.concat(all_entries, ignore_index=True)

    # Normalize token strings and coerce verse to numeric
    df_combined["Token"] = df_combined["Token"].astype(str).str.strip()
    df_combined["Vers"] = pd.to_numeric(df_combined["Vers"], errors="coerce")

    # normalize meta cols (safe even if some cols are missing)
    for c in meta_cols:
        if c in df_combined.columns:
            df_combined[c] = df_combined[c].astype(str).str.strip()

    # ======================================================================
    # [5] Frequency calculation and interactive token selection
    # ======================================================================
    naming_values = [
        v.strip()
        for col in naming_cols
        for v in df_figure[col].dropna().astype(str)
        if v.strip() != ""
    ]

    epithet_values = [
        v.strip()
        for col in epithet_cols
        for v in df_figure[col].dropna().astype(str)
        if v.strip() != ""
    ]

    naming_list = Counter(naming_values).most_common()
    epithet_list = Counter(epithet_values).most_common()

    selected_naming = []
    selected_epithets = []

    if variant_type in ("1", "3"):
        print(f"\nAvailable naming variants for {figure_name}:")
        for i, (token, freq) in enumerate(naming_list, 1):
            print(f"{i}. {token} – {freq}")
        while True:
            input_str = input(
                "\nWhich naming variants should be included? (e.g., 1–3, 5)\n"
                "Note: Selecting more than 14 entries in total may reduce visual clarity.\n> "
            ).strip()
            indices = parse_token_selection(input_str, len(naming_list))
            if indices:
                selected_naming = [naming_list[i - 1][0] for i in indices]
                break
            print("Invalid input – please try again.")

    if variant_type in ("2", "3"):
        print(f"\nAvailable epithets for {figure_name}:")
        for i, (token, freq) in enumerate(epithet_list, 1):
            print(f"{i}. {token} – {freq}")
        while True:
            input_str = input(
                "\nWhich epithets should be included? (e.g., 1–3, 5)\n"
                "Note: Selecting more than 14 entries in total may reduce visual clarity.\n> "
            ).strip()
            indices = parse_token_selection(input_str, len(epithet_list))
            if indices:
                selected_epithets = [epithet_list[i - 1][0] for i in indices]
                break
            print("Invalid input – please try again.")

    # Combine selected tokens across units
    tokens_to_plot = selected_naming + selected_epithets
    if not tokens_to_plot:
        print("No tokens selected – aborting.")
        return

    # ======================================================================
    # [6] Filter and prepare plotting data
    # ======================================================================

    df_plot = df_combined[df_combined["Token"].isin(tokens_to_plot)].copy()

    # Remove rows without valid verse numbers (non-renderable points)
    df_plot = df_plot.dropna(subset=["Vers"])

    # Determine token order by descending frequency in filtered set
    plot_token_counts = Counter(df_plot["Token"])
    sorted_tokens = [token for token, _ in plot_token_counts.most_common()]

    # Create HTML-formatted token labels (italicized)
    df_plot["Token_html"] = df_plot["Token"].apply(lambda x: f"<i>{x}</i>")
    df_plot["Token_html"] = pd.Categorical(
        df_plot["Token_html"],
        categories=[f"<i>{t}</i>" for t in sorted_tokens],
        ordered=True
    )

    # Add categorical unit labels in combined mode
    if variant_type == "3":
        df_plot["Category"] = df_plot["Token"].apply(
            lambda x: "Naming variants" if x in selected_naming else "Epithets"
        )

    # ======================================================================
    # [7] Create Plotly figure and traces
    # ======================================================================

    # Track trace-to-token mapping for Show-N visibility control (used in combined mode)
    trace_tokens: list[str | None] = []

    if variant_type == "3":
        # Combined mode: category coloring (Naming variants vs Epithets)
        fig = go.Figure()

        # Map categories to global palette colors
        color_map = {
            "Naming variants": GLOBAL_VISUAL_STYLE["colors"]["categories"]["Naming variants"],
            "Epithets": GLOBAL_VISUAL_STYLE["colors"]["categories"]["Epithets"],
        }

        # Enforce stable token order (frequency-descending)
        token_order = sorted_tokens

        # Legend dummy trace: Naming variants (single compact legend entry)
        fig.add_trace(
            go.Scatter(
                x=[None],
                y=[None],
                mode="markers",
                marker={"size": 18, "opacity": 0.7, "color": color_map["Naming variants"]},
                name="Naming variants",
                legendgroup="Naming variants",
                showlegend=True,
                hoverinfo="skip",
            )
        )
        trace_tokens.append(None)

        # Legend dummy trace: Epithets (single compact legend entry)
        fig.add_trace(
            go.Scatter(
                x=[None],
                y=[None],
                mode="markers",
                marker={"size": 18, "opacity": 0.7, "color": color_map["Epithets"]},
                name="Epithets",
                legendgroup="Epithets",
                showlegend=True,
                hoverinfo="skip",
            )
        )
        trace_tokens.append(None)

        # One trace per token (enables token-level Show-N visibility)
        for token in token_order:
            df_token = df_plot[df_plot["Token"] == token].copy()
            if df_token.empty:
                continue

            # Render token labels in italics (HTML)
            token_html = f"<i>{token}</i>"

            # Derive category for the token (Naming variants or Epithets)
            category = df_token["Category"].iloc[0]

            # Select marker color by category (fallback to auxiliary color)
            marker_color = color_map.get(category, GLOBAL_VISUAL_STYLE["colors"]["levels"]["AUXILIARY"])

            fig.add_trace(
                go.Scatter(
                    x=df_token["Vers"],
                    y=[token_html] * len(df_token),
                    mode="markers",
                    marker={"opacity": 0.7, "color": marker_color},
                    name=token_html,            # token identifier (used by some downstream logic)
                    legendgroup=category,
                    showlegend=False,           # legend handled by dummy traces above
                    meta=token_html,            # used in hovertemplate
                    hovertemplate="Vers: %{x}<br>Token: %{meta}<extra></extra>",
                )
            )
            trace_tokens.append(token)  # raw token id for visibility toggles

        fig.update_layout(title=f"{variant_label} for '{figure_name}'")
    else:
        # Single-unit mode: uniform coloring (either naming variants or epithets)
        fig = go.Figure()

        # trace_tokens is not used in this branch (visibility is currently name-based)
        trace_tokens = []  # keep consistent type; not used in this branch

        # Pick unit color from global palette
        base_color = (
            GLOBAL_VISUAL_STYLE["colors"]["categories"]["Naming variants"]
            if variant_type == "1"
            else GLOBAL_VISUAL_STYLE["colors"]["categories"]["Epithets"]
        )

        # Enforce stable token order (frequency-descending)
        token_order = sorted_tokens

        # One trace per token (legend suppressed)
        for token in token_order:
            df_token = df_plot[df_plot["Token"] == token].copy()
            if df_token.empty:
                continue

            token_html = f"<i>{token}</i>"

            fig.add_trace(
                go.Scatter(
                    x=df_token["Vers"],
                    y=[token_html] * len(df_token),
                    mode="markers",
                    marker={"size": 18, "opacity": 0.7, "color": base_color},
                    name=token_html,            # retained for future unified visibility logic
                    showlegend=False,           # IMPORTANT: no per-token legend
                    meta=token_html,
                    hovertemplate="Vers: %{x}<br>Token: %{meta}<extra></extra>",
                )
            )

    # ======================================================================
    # [8] Configure axes, defaults, and global visual styling
    # ======================================================================

    # Default Top-N view: calibrated for A4 readability (up to 14 tokens)
    max_n = len(sorted_tokens)
    default_show_n = min(14, max_n)

    # Derive initial y-axis category list (HTML-italic labels)
    top_tokens = sorted_tokens[:default_show_n]
    top_categories = [f"<i>{t}</i>" for t in top_tokens]

    # Configure y-axis as categorical with explicit ordering
    fig.update_yaxes(
        type="category",
        categoryorder="array",
        categoryarray=top_categories,
        tickmode="array",
        tickvals=top_categories,
        ticktext=top_categories,
        title_text=variant_label,
    )

    # Configure x-axis label
    fig.update_xaxes(title_text="Verses")

    # Apply global visual defaults (fonts, margins, template, etc.)
    apply_global_visual_style(fig)

    # Apply global visibility defaults (legend only in combined mode)
    apply_global_visual_visibility(fig, show_legend=(variant_type == "3"))

    # Keep plot-specific height (axis-heavy chart)
    fig.update_layout(height=800)

    # ======================================================================
    # [9] Interactive "Show N" dropdown (Top-14, Top-28, ...)
    # ======================================================================
    def compute_tick_size(token_count: int) -> int:
        """
        Compute y-axis tick font size based on the number of displayed tokens.

        Calibrated for A4 readability:
        - Up to 14 tokens → fixed size (36).
        - Above 14 tokens → proportional downscaling.
        - Lower bound ensures minimal legibility.
        """
        if token_count <= 14:
            return 36
        size = int(round(36 * 14 / token_count))
        return max(12, size)

    def compute_marker_size(token_count: int) -> int:
        """
        Compute marker size proportional to the tick font size.

        Calibrated for A4 readability:
        - Up to 14 tokens → fixed size (18).
        - Above 14 tokens → proportional downscaling.
        - Lower bound prevents markers from becoming visually indistinct.
        """
        if token_count <= 14:
            return 18
        size = int(round(18 * 14 / token_count))
        return max(6, size)

    # Build Show-N step list (14, 28, 42, ... up to max)
    show_steps = list(range(14, max_n + 1, 14))

    if not show_steps:
        show_steps = [max_n]
    elif show_steps[-1] != max_n:
        show_steps.append(max_n)

    # Initialize trace visibility for the default Top-N view
    if variant_type == "3":
        # Combined mode: visibility driven by trace_tokens (includes dummy legend traces)
        initial_visible = []
        top_set = set(top_tokens)

        for tok in trace_tokens:
            if tok is None:
                initial_visible.append(True)         # always show legend dummies
            else:
                initial_visible.append(tok in top_set)

        for i, tr in enumerate(fig.data):
            tr.visible = initial_visible[i]
    else:
        # Single-unit mode: visibility currently driven by trace.name (HTML token label)
        visible_by_name = {f"<i>{t}</i>": (t in top_tokens) for t in sorted_tokens}

        initial_visible = []
        for tr in fig.data:
            trace_name = str(getattr(tr, "name", None))
            initial_visible.append(visible_by_name.get(trace_name, True))

        for i, tr in enumerate(fig.data):
            tr.visible = initial_visible[i]

    # Build dropdown buttons for each Show-N step
    buttons = []
    for n in show_steps:
        current_tokens = sorted_tokens[:n]
        current_categories = [f"<i>{t}</i>" for t in current_tokens]
        current_tick_size = compute_tick_size(n)
        current_marker_size = compute_marker_size(n)

        # Compute trace visibility list for this step
        if variant_type == "3":
            current_set = set(current_tokens)
            visible_list = [(tok is None) or (tok in current_set) for tok in trace_tokens]
        else:
            visible_map = {f"<i>{t}</i>": (t in current_tokens) for t in sorted_tokens}
            visible_list = []
            for tr in fig.data:
                trace_name = str(getattr(tr, "name", None))
                visible_list.append(visible_map.get(trace_name, True))

        buttons.append({
            "label": f"Show {n}",
            "method": "update",
            "args": [
                {
                    "visible": visible_list,
                    "marker.size": current_marker_size,
                },
                {
                    "yaxis": {
                        "type": "category",
                        "categoryorder": "array",
                        "categoryarray": current_categories,
                        "tickmode": "array",
                        "tickvals": current_categories,
                        "ticktext": current_categories,
                        "tickfont": {"size": current_tick_size},
                    }
                },
            ],
        })

    # Attach dropdown to the figure layout
    fig.update_layout(
        updatemenus=[
            {
                "type": "dropdown",
                "direction": "down",
                "x": 0.0,
                "y": 1.18,
                "xanchor": "left",
                "yanchor": "top",
                "showactive": True,
                "buttons": buttons,
            }
        ]
    )

    # Ensure initial tick size matches the default Top-N view
    fig.update_yaxes(tickfont={"size": compute_tick_size(default_show_n)})

    # Ensure initial marker size matches the default Top-N view
    default_marker_size = compute_marker_size(default_show_n)

    fig.update_traces(
        marker={"size": default_marker_size},
        selector=dict(mode="markers")
    )

    # ======================================================================
    # [10] Export (HTML)
    # ======================================================================

    filename_label = "combined" if variant_type == "3" else "epithets" if variant_type == "2" else "naming"
    filename = f"viz_{filename_label}_{figure_name}.html"

    export_visualization_output(
        fig,
        paths=paths,
        book_name=book_name,
        filename=filename,
        export_func=apply_global_visual_modebar_export,
    )

# =============================================================================
# VISUALIZATION – INTRA-NAMED-FIGURE CO-OCCURRENCE HEATMAP
# =============================================================================


def visualize_intra_figure_cooccurrence_heatmap(paths: dict, book_name: str) -> None:
    """
    Interactive workflow for intra-figure co-occurrence heatmap generation.

    This function provides the CLI-driven entry point for computing and
    visualizing intra-figure co-occurrence of labeling tokens
    (naming variants and/or epithets) within a selected book.

    It loads categorization data, filters entries by the selected figure,
    computes unordered token pair frequencies on a per-entry basis
    (binary co-occurrence per entry), applies thresholding and Top-N
    selection, and renders a symmetric heatmap visualization.

    Percent values in the heatmap represent pair-based shares:
    each unordered token pair contributes once to the total (100%),
    although it is displayed twice due to symmetric matrix rendering.

    Parameters:
        paths (dict):
            Dictionary containing resolved project paths.
            Must include the key "categorization_json".

        book_name (str):
            Short identifier of the currently active book; used for
            export naming and contextual labeling.

    Returns:
        None

    Behavior:
        Prompts the user interactively for figure selection and labeling scope.
        Prints diagnostic messages if required data is missing or if no
        co-occurring pairs meet the defined threshold.
        Renders and exports the resulting visualization.
    """
    categorization_path = paths.get("categorization_json")
    if not categorization_path:
        print("\nMissing path: paths['categorization_json'] is not set.")
        return

    # ======================================================================
    # [1] User input: figure and labeling scope
    # ======================================================================

    figure_name = ask_valid_figure_name(categorization_path)

    print("\nWhat should be visualized?")
    print("[1] Combined")
    print("[2] Naming variants")
    print("[3] Epithets")
    variant_type = ask_user_choice("> ", ["1", "2", "3"])

    include_naming_variants = (variant_type in ("1", "2"))
    include_epithets = (variant_type in ("1", "3"))

    # Fixed defaults (silent)
    min_pair_count = 2
    top_n = 30

    # ======================================================================
    # [2] Load categorized data and restrict to target figure
    # ======================================================================
    entries = safe_read_json(categorization_path, default=[])
    rows = [e for e in entries if isinstance(e, dict) and e.get("Benannte Figur") == figure_name]

    # Collect per-entry token lists (naming variants and/or epithets)
    token_rows = [
        collect_tokens_for_cooccurrence(r, include_naming_variants, include_epithets)
        for r in rows
    ]

    # Guard: only entries with at least two tokens can produce co-occurrence pairs
    token_rows = [t for t in token_rows if len(t) >= 2]

    # ======================================================================
    # [3] Pair counting: unordered, per-entry deduplicated (binary co-occurrence)
    # ======================================================================

    pair_counter: Counter = Counter()

    for toks in token_rows:
        # Binary per-entry co-occurrence:
        # remove duplicates within an entry and enforce a stable order.
        toks_unique = sorted(set(toks))

        # Unordered pairs (stable because toks_unique is sorted)
        for a, b in combinations(toks_unique, 2):
            pair = (a, b)
            pair_counter[pair] += 1

    # Apply minimum frequency threshold
    pair_counter = Counter({p: c for p, c in pair_counter.items() if c >= min_pair_count})
    if not pair_counter:
        print("\nNo co-occurring pairs met the minimum threshold.")
        return

    # ======================================================================
    # [4] Top-N selection and matrix construction (symmetric rendering)
    # ======================================================================

    # Select most frequent pairs (bounded display)
    top_pairs = pair_counter.most_common(top_n)

    # Derive the displayed token set from the selected pairs
    tokens = sorted(set([t for p, _ in top_pairs for t in p]))
    index = {t: i for i, t in enumerate(tokens)}

    # Build full symmetric matrix: each undirected pair is written to both halves
    size = len(tokens)
    matrix = np.zeros((size, size), dtype=float)

    for (a, b), c in top_pairs:
        i, j = index[a], index[b]
        if i == j:
            continue  # diagonal stays 0
        matrix[i, j] = c
        matrix[j, i] = c

    # ======================================================================
    # [5] Normalization: pair-based percent shares
    # ======================================================================

    # Pair-based normalization:
    # each unordered pair contributes once to the total (100%),
    # although it appears twice in the symmetric matrix.
    total_pairs = float(sum(c for _, c in top_pairs))
    if total_pairs > 0:
        matrix_pct = (matrix / total_pairs) * 100.0
    else:
        matrix_pct = matrix.copy()

    # ======================================================================
    # [6] Colorscale and contrast cap (global palette alignment)
    # ======================================================================

    levels = GLOBAL_VISUAL_STYLE["colors"]["levels"]

    # Two-stage sequential gradient: STRUCTURE → AUXILIARY → CORE
    heatmap_colorscale = (
        [rgb_tuple_to_plotly_color(c) for c in n_colors(
            hex_color_to_rgb_tuple(levels["STRUCTURE"]),
            hex_color_to_rgb_tuple(levels["AUXILIARY"]),
            24,
            colortype="tuple",
        )]
        + [rgb_tuple_to_plotly_color(c) for c in n_colors(
            hex_color_to_rgb_tuple(levels["AUXILIARY"]),
            hex_color_to_rgb_tuple(levels["CORE"]),
            24,
            colortype="tuple",
        )][1:]
    )

    # Robust upper cap to preserve contrast under skewed distributions
    z_cap = float(np.nanpercentile(matrix_pct, 99))
    z_cap = min(100.0, max(10.0, z_cap))  # safety bounds

    # ======================================================================
    # [7] Plotly heatmap rendering (hover = percent + absolute count)
    # ======================================================================

    fig = px.imshow(
        matrix_pct,
        x=tokens,
        y=tokens,
        labels=dict(x="Token", y="Token", color="Co-occurrence share (%)"),
        aspect="auto",
        range_color=(0, z_cap),
        color_continuous_scale=heatmap_colorscale,
    )

    # Hover shows:
    # - share (% of all displayed unordered pairs)
    # - absolute pair count (from the non-normalized matrix)
    fig.update_traces(
        customdata=matrix,
        hovertemplate=(
            "Token %{y} × %{x}<br>"
            "Share: %{z:.1f}%<br>"
            "Count: %{customdata}<extra></extra>"
        )
    )

    fig.update_layout(
        title={
            "text": f"Intra-figure co-occurrence of {figure_name}",
            "pad": {"t": 10},
        }
    )

    # ======================================================================
    # [8] Apply global visual defaults and plot-specific layout tuning
    # ======================================================================

    apply_global_visual_style(fig, tick_font_size=20, show_grid=False)
    apply_global_visual_visibility(fig, show_legend=False)

    # Heatmap-specific: ensure enough top margin for title after global style is applied
    fig.update_layout(margin={**GLOBAL_VISUAL_STYLE["layout"]["margins"], "t": 100})

    # ======================================================================
    # [9] Export (HTML)
    # ======================================================================

    filename_label = "intra_figure_cooccurrence"
    filename = f"viz_{filename_label}_{figure_name}.html"
    filename_stub = os.path.splitext(filename)[0]

    export_visualization_output(
        fig,
        paths=paths,
        book_name=book_name,
        filename=filename,
        filename_stub=filename_stub,
        export_func=apply_global_visual_modebar_export,
    )

# =============================================================================
# VISUALIZATION – SUNBURST
# =============================================================================

def run_sunburst_visualization(paths, book_name, data):
    """
    Central interactive dispatcher for Sunburst visualizations.

    This function provides the CLI entry point for selecting and executing
    one of the available Sunburst visualization modes:
    - figure-centered view
    - book-centered overview

    It does not perform any data transformation itself but delegates execution
    to the respective visualization functions based on validated user input.
    The menu runs in a blocking loop until the user exits.

    Parameters:
        paths (dict):
            Dictionary containing resolved project paths required by the
            downstream Sunburst visualization functions.

        book_name (str):
            Short identifier of the currently active book; passed to
            visualization functions for contextual labeling and export naming.

        data (dict):
            Loaded project data structures required by the Sunburst views.

    Returns:
        None

    Behavior:
        Prints interactive prompts to stdout and waits for user input.
        Delegates control flow to the selected Sunburst visualization.
        Exits the loop when the user chooses to return.
    """
    while True:
        print("\nSunburst Visualization")
        print("[1] Figure-centered view")
        print("[2] Book-centered overview")
        print("[3] Back to visualization menu")

        choice = ask_user_choice("> ", ["1", "2", "3"])

        if choice == "1":
            visualize_sunburst_figure_view(paths, book_name, data)

        elif choice == "2":
            visualize_sunburst_book_overview(paths, book_name, data)

        elif choice == "3":
            print("Returning to visualization menu.")
            return

def visualize_sunburst_figure_view(paths, book_name, data):
    """
    Interactive figure-centered Sunburst visualization workflow.

    This function provides the CLI-driven entry point for generating
    Sunburst visualizations centered on a selected figure within a book.
    It supports two structurally distinct modes:

        [1] Named Figure-centered mode:
            center_figure → type_group → lemma
            (distribution of naming variants and epithets by type)

        [2] Namer-centered mode:
            center_figure → namer → lemma
            (distribution by naming figure and/or narrator)

    Data source and preparation:
        - Naming data is loaded via load_naming_sources_with_excel_fallback(...).
        - prepare_naming_data(...) is used to obtain a standardized DataFrame
          and column mapping.
        - The function assumes canonical, pre-categorized naming data
          (no additional fuzzy matching or similarity heuristics are applied).

    Aggregation logic:
        - Frequencies are computed per lemma relative to the selected figure.
        - Percent shares are normalized to the total frequency of the
          center figure (pct_of_center).
        - Sorting is applied for internal consistency; Plotly retains
          its hierarchical structure independently.

    Visualization:
        - Plotly Sunburst chart (px.sunburst) with explicit path definition.
        - Color mapping derived from GLOBAL_VISUAL_STYLE["colors"].
        - Minimal post-processing for:
            * root and ring coloring,
            * optional lexeme-level refinement for epithets,
            * WCAG-oriented text contrast adjustment.
        - Hover data includes group, category, lemma, frequency,
          and share relative to the center figure.

    Output handling:
        - Delegates export logic to export_visualization_output(...).
        - Supports save, show in browser (temporary file), or both.
        - Filename encodes visualization variant and figure identifier.

    Guards and BETA semantics:
        - Aborts with a diagnostic message if required naming data
          is unavailable or if aggregation yields no result.
        - Performs no additional schema validation beyond what
          prepare_naming_data(...) provides (BETA state).

    Parameters:
        paths (dict):
            Dictionary containing resolved project paths.
            Must include the key "categorization_json".

        book_name (str):
            Identifier of the currently active book;
            used for contextual labeling and export naming.

        data (dict):
            Loaded project data structures required by
            naming data loaders and aggregation helpers.

    Returns:
        None

    Behavior:
        Prompts the user interactively for figure selection and mode.
        Builds and renders a Sunburst visualization.
        Exports and/or displays the result according to user choice.
    """
    # ======================================================================
    # [1] Load and normalize naming data
    # ======================================================================
    df_json, df_excel = load_naming_sources_with_excel_fallback(paths, data)
    _, df, cols = prepare_naming_data(book_name, df_json, df_excel)

    # Guard: abort early if no usable naming data is available
    if df is None or df.empty:
        print("No naming data available after prepare_naming_data.")
        return

    # ======================================================================
    # [2] Figure selection (validated against categorization JSON)
    # ======================================================================

    categorization_path = paths.get("categorization_json") if isinstance(paths, dict) else None
    if not categorization_path:
        print("\nMissing required path: 'categorization_json'. Cannot run Sunburst figure view.")
        return
    figure_name = ask_valid_figure_name(categorization_path)

    # ======================================================================
    # [3] Mode selection (hierarchy definition)
    # ======================================================================

    print()
    print("[1] Named Figure-centered mode: Types → Lemma")
    print("[2] Namer-centered mode: Naming figures and/or narrator → Lemma")
    mode_choice = ask_user_choice("> ", ["1", "2"])

    # ======================================================================
    # [4] Aggregation (mode-dependent data preparation)
    # ======================================================================

    if mode_choice == "1":
        sunburst_df = build_sunburst_data_types_lemma(df, cols, figure_name)
    else:
        sunburst_df = build_sunburst_data_namer_lemma(df, cols, figure_name)

    # Guard: abort if aggregation yields no result
    if sunburst_df is None or sunburst_df.empty:
        print("No data available for this sunburst configuration.")
        return

    # ======================================================================
    # [5] Compute shares relative to center figure
    # ======================================================================

    total_freq = sunburst_df["frequency"].sum()
    if total_freq > 0:
        sunburst_df["pct_of_center"] = sunburst_df["frequency"] / total_freq
    else:
        sunburst_df["pct_of_center"] = 0.0

    # ======================================================================
    # [6] Internal sorting (stability for hover + grouping logic)
    # ======================================================================

    sunburst_df["type_group"] = pd.Categorical(
        sunburst_df["type_group"],
        categories=["Naming variants", "Epithets"],
        ordered=True,
    )

    sunburst_df = sunburst_df.sort_values(
        ["type_group", "lemma"],
        ascending=[True, True],
    ).reset_index(drop=True)

    # ======================================================================
    # [7] Centralized color configuration
    # ======================================================================

    categories = GLOBAL_VISUAL_STYLE["colors"]["categories"]
    levels = GLOBAL_VISUAL_STYLE["colors"]["levels"]

    # Single source of truth for semantic colors
    color_map = {
        "Proper name": categories.get("Proper name"),
        "Antonomasia": categories.get("Antonomasia"),
        "Naming variants": categories.get("Naming variants"),
        "Epithets": categories.get("Epithets"),
        "Epithets (lexeme level)": categories.get("Epithets (lexeme level)"),
        "STRUCTURE": levels.get("STRUCTURE"),
    }

    # ======================================================================
    # [8] Mode 1 – Figure-centered (Types → Lemma)
    # ======================================================================

    if mode_choice == "1":
        name_color = color_map.get("Proper name")
        naming_variants_ring = color_map.get("Naming variants")
        epithets_ring = color_map.get("Epithets")
        epithets_lexeme = color_map.get("Epithets (lexeme level)")

        fig = px.sunburst(
            sunburst_df,
            path=["center_figure", "type_group", "lemma"],
            values="frequency",
            color="color_group",
            color_discrete_map=color_map,
        )

        if fig.data:
            trace = fig.data[0]

            # Harmonized segment borders
            border_rgba = hex_color_to_rgba(levels.get("AUXILIARY"), 0.35)
            trace.update(
                marker=dict(
                    line=dict(
                        color=border_rgba,
                        width=0.8,
                    )
                )
            )

            labels = list(trace["labels"])
            parents = list(trace["parents"])
            colors = list(trace["marker"]["colors"])

            epithets_parent = f"{figure_name}/Epithets"

            for i, (lab, par) in enumerate(zip(labels, parents)):

                # Root node
                if lab == figure_name and (par is None or par == ""):
                    colors[i] = name_color
                    continue

                # Ring nodes
                ring_colors = {
                    "Naming variants": naming_variants_ring,
                    "Epithets": epithets_ring,
                }
                if lab in ring_colors:
                    colors[i] = ring_colors[lab]
                    continue

                # Epithets lexeme-level refinement
                if str(par) == epithets_parent:
                    colors[i] = epithets_lexeme

            trace["marker"]["colors"] = colors
            apply_accessible_text_colors(
                trace,
                colors,
                levels.get("NEUTRAL_TEXT"),
                levels.get("LIGHT_TEXT"),
            )

            # Build hover customdata
            group_totals = (
                sunburst_df.groupby("type_group", observed=False)["frequency"]
                .sum()
                .to_dict()
            )

            lemma_map: dict[str, list[object]] = {
                str(row["lemma"]): [
                    str(row["type_group"]),  # Group
                    str(row["color_group"]),  # Category
                    str(row["lemma"]),  # Lemma
                    float(row["frequency"]),
                    float(row["pct_of_center"]),
                ]
                for _, row in sunburst_df.iterrows()
            }

            # Assemble customdata aligned with trace["labels"].
            customdata: list[list[object]] = []
            for lab in labels:
                # Root node (center figure)
                if lab == figure_name:
                    customdata.append(["", "", str(lab), float(total_freq), 1.0])

                # Ring nodes (Naming variants / Epithets)
                elif lab in group_totals:
                    cnt = float(group_totals.get(lab, 0.0))
                    share = (cnt / total_freq) if total_freq > 0 else 0.0
                    customdata.append([str(lab), "", str(lab), cnt, share])

                # Leaf nodes (lemma level)
                else:
                    info = lemma_map.get(str(lab), ["", "", str(lab), 0.0, 0.0])
                    customdata.append(info)

            trace["customdata"] = customdata
            trace["hovertemplate"] = (
                "Center figure: %{root}<br>"
                "Group: %{customdata[0]}<br>"
                "Category: %{customdata[1]}<br>"
                "Lemma: %{customdata[2]}<br>"
                "Count: %{customdata[3]}<br>"
                "Share (center): %{customdata[4]:.1%}<extra></extra>"
            )

    # ======================================================================
    # [9] Mode 2 – Namer-centered (Namer → Lemma)
    # ======================================================================

    else:
        # Normalize namer display string (extract last segment after "/" or "#")
        sunburst_df["namer_display"] = sunburst_df["namer"].apply(
            lambda v: (
                v[max(v.rfind("/"), v.rfind("#")) + 1:].strip()
                if isinstance(v, str) and max(v.rfind("/"), v.rfind("#")) != -1
                else v
            )
        )

        # Aggregate frequencies per namer (ring level)
        per_namer_raw = (
            sunburst_df.groupby("namer_display", dropna=False)["frequency"]
            .sum()
            .to_dict()
        )
        total_all = float(sum(per_namer_raw.values())) or 0.0

        fig = px.sunburst(
            sunburst_df,
            path=["center_figure", "namer_display", "lemma"],
            values="frequency",
            color="color_group",
            color_discrete_map=color_map,
        )

        if fig.data:
            trace = fig.data[0]

            # Harmonized segment borders
            border_rgba = hex_color_to_rgba(color_map.get("STRUCTURE"), 0.35)
            trace.update(
                marker=dict(line=dict(color=border_rgba, width=0.8))
            )

            labels = list(trace["labels"])
            parents = list(trace["parents"])
            values = list(trace["values"])
            colors = list(trace["marker"]["colors"])

            # Leaf-type lookup: (namer_display, lemma) → color_group
            df_leaf_type = (
                sunburst_df.loc[:, ["namer_display", "lemma", "color_group"]]
                .dropna(subset=["namer_display", "lemma"])
                .drop_duplicates(subset=["namer_display", "lemma"])
            )
            leaf_type_by_pair: dict[tuple[str, str], str] = {
                (str(r.namer_display), str(r.lemma)): str(r.color_group)
                for r in df_leaf_type.itertuples(index=False)
            }

            # Row count per namer (= actual number of mentions, not token sum)
            namer_mentions = (
                sunburst_df.drop_duplicates(subset=["namer_display"])
                .set_index("namer_display")["namer_row_count"]
                .to_dict()
            )
            total_mentions = float(sum(namer_mentions.values())) or 0.0

            # Build hover payload
            customdata: list[list[object]] = []
            for lab, par, val in zip(labels, parents, values):

                # Root node
                if lab == figure_name and (par is None or par == ""):
                    customdata.append(["", "", str(lab), float(total_all), 1.0])
                    continue

                # Namer ring
                if par == figure_name:
                    mentions = float(namer_mentions.get(lab, 0))
                    share_namer = (mentions / total_mentions) if total_mentions > 0 else 0.0
                    customdata.append([str(lab), "", "", mentions, share_namer])
                    continue

                # Leaf node
                par_str = "" if par is None else str(par)
                lab_str = "" if lab is None else str(lab)

                type_ui = leaf_type_by_pair.get((par_str, lab_str), "")
                freq = float(val or 0.0)
                share = (freq / total_all) if total_all > 0 else 0.0
                customdata.append([par_str, type_ui, lab_str, freq, share])

            trace["customdata"] = customdata
            trace["hovertemplate"] = (
                "Center figure: %{root}<br>"
                "Namer: %{customdata[0]}<br>"
                "Type: %{customdata[1]}<br>"
                "Lemma: %{customdata[2]}<br>"
                "Count: %{customdata[3]}<br>"
                "Share (center): %{customdata[4]:.1%}<extra></extra>"
            )

            # Minimal color patching: center + namer ring
            structure_rgba = hex_color_to_rgba(color_map.get("STRUCTURE"), 0.55)
            name_color = color_map.get("Proper name")

            for i, (lab, par) in enumerate(zip(labels, parents)):
                # Root node
                if lab == figure_name and (par is None or par == ""):
                    colors[i] = name_color
                    continue

                # First ring (namer level) → parent is center figure
                if str(par) == str(figure_name):
                    colors[i] = structure_rgba

            trace["marker"]["colors"] = colors

            # Apply accessible text colors after final color assignment
            apply_accessible_text_colors(
                trace,
                colors,
                levels.get("NEUTRAL_TEXT"),
                levels.get("LIGHT_TEXT"),
            )

    # ======================================================================
    # [10] Global styling and layout
    # ======================================================================

    if "fig" not in locals() or fig is None:
        print("No figure could be created for this configuration.")
        return

    apply_global_visual_style(fig, has_axes=False)
    apply_global_visual_visibility(fig, show_axis_labels=False)

    fig.update_layout(
        title={
        "text": f"Sunburst – {figure_name} ({book_name})",
        "pad": {"t": 10},        },
        margin = {"t": 80, "l": 20, "r": 20, "b": 20},
    )

    # ======================================================================
    # [11] Export (delegated to centralized helper)
    # ======================================================================

    # sanitize figure_name for filename
    safe_figure = "".join(
        c if c.isalnum() or c in ("_", "-") else "_" for c in str(figure_name)
    )
    variant_label = "sunburst_figure_types" if mode_choice == "1" else "sunburst_figure_namers"

    filename_stub = f"viz_{variant_label}_{safe_figure}"
    filename = f"{filename_stub}.html"

    export_visualization_output(
        fig,
        paths=paths,
        book_name=book_name,
        filename=filename,
        export_func=apply_global_visual_modebar_export,
        filename_stub=filename_stub,
    )

def visualize_sunburst_book_overview(paths, book_name, data):
    """
    Interactive work-centered Sunburst visualization workflow.

    This function provides the CLI-driven entry point for generating
    a Sunburst visualization centered on a single work (book overview).
    The hierarchy is structurally defined as:

        root (book_name)
            → figure (Top-K by total naming frequency)
                → type
                    ("Proper name", "Antonomasia", "Epithets")

    Data source and preparation:
        - Naming data is loaded via load_naming_sources_with_excel_fallback(...).
        - prepare_naming_data(...) is used to obtain a standardized DataFrame
          and column mapping.
        - The function assumes canonical, pre-categorized naming data.
          No additional fuzzy matching, similarity heuristics, or
          re-categorization is applied at this stage.

    Aggregation logic:
        - Figures are ranked by total naming frequency within the selected work.
        - The user selects the number of top figures (Top-K; default = 12).
        - Aggregation is performed via build_sunburst_data_book_overview(...),
          producing a hierarchical structure (root → figure → type).
        - Values represent absolute naming frequencies.

    Visualization:
        - Plotly Sunburst chart (px.sunburst) with explicit path definition
          ["root", "figure", "type"].
        - Color mapping derived from GLOBAL_VISUAL_STYLE["colors"].
        - Post-processing ensures:
            * consistent root and figure-ring styling,
            * harmonized segment borders,
            * contrast-aware text color adjustment for all rings
              (per-segment foreground selection based on background color).

    Output handling:
        - Delegates export logic to export_visualization_output(...).
        - Supports save, show in browser (temporary file), or both.
        - Filename encodes visualization variant and book identifier.

    Guards and BETA semantics:
        - Aborts with a diagnostic message if naming data is unavailable
          after normalization or if aggregation yields no result.
        - Relies on prepare_naming_data(...) for structural guarantees.
          No additional schema validation is performed here (BETA state).
        - Assumes required keys exist in GLOBAL_VISUAL_STYLE;
          missing critical style keys result in a controlled abort.

    Parameters:
        paths (dict):
            Dictionary containing resolved project paths required by
            loaders and export helpers.

        book_name (str):
            Identifier of the currently active book;
            used as root label and for export naming.

        data (dict):
            Loaded project data structures required by
            naming data loaders and aggregation helpers.

    Returns:
        None

    Behavior:
        Prompts the user interactively for Top-K selection.
        Builds and renders a work-centered Sunburst visualization.
        Exports and/or displays the result according to user choice.
    """
    # ======================================================================
    # [1] Load and normalize naming data
    # ======================================================================

    # Load naming sources (JSON preferred, Excel fallback) using the centralized loader.
    df_json, df_excel = load_naming_sources_with_excel_fallback(paths, data)

    # Normalize naming data; obtain standardized DataFrame + column mapping.
    _, df, cols = prepare_naming_data(book_name, df_json, df_excel)

    # Guard: abort if normalization yields no usable naming data.
    if df is None or df.empty:
        print("No naming data available after prepare_naming_data.")
        return

    # ======================================================================
    # [2] Top-K selection (CLI interaction)
    # ======================================================================

    # Inform user about automatic ranking basis (total naming frequency).
    print()
    print("ℹ Top figures will be selected automatically based on total naming frequency.")

    default_top_k = 12

    # Prompt user for number of top figures (positive integer or Enter for default).
    prompt = (
        f"Enter number of top figures to include "
        f"[Press Enter to use default: {default_top_k}]:\n> "
    )

    # Validate numeric input; enforce strictly positive integer.
    while True:
        user_input = input(prompt).strip()

        if not user_input:
            top_k = default_top_k
            break

        if user_input.isdigit():
            value = int(user_input)
            if value > 0:
                top_k = value
                break

        print("Please enter a positive integer or press Enter for the default.")

    # ======================================================================
    # [3] Build aggregated data structure (root → figure → type)
    # ======================================================================

    # Aggregate naming frequencies per figure and type for work-centered overview.
    sunburst_df = build_sunburst_data_book_overview(df, cols, top_k, book_name)

    # Guard: abort if aggregation yields no result.
    if sunburst_df is None or sunburst_df.empty:
        print("No data available for work-centered sunburst overview.")
        return

    # ======================================================================
    # [4] Visualization setup (Sunburst construction)
    # ======================================================================

    # Access global color configuration (categories + structural levels).
    categories = GLOBAL_VISUAL_STYLE["colors"]["categories"]
    levels = GLOBAL_VISUAL_STYLE["colors"]["levels"]

    # Extract semantic category colors for type ring.
    proper_name_color = categories.get("Proper name")
    antonomasia_color = categories.get("Antonomasia")
    epithets_color = categories.get("Epithets")

    # Map naming types to their respective UI colors.
    type_color_map = {
        "Proper name": proper_name_color,
        "Antonomasia": antonomasia_color,
        "Epithets": epithets_color,
    }

    # Create hierarchical Sunburst (root → figure → type).
    fig = px.sunburst(
        sunburst_df,
        path=["root", "figure", "type"],
        values="value",
        color="type",
        color_discrete_map=type_color_map,
    )

    # ======================================================================
    # [5] Post-processing: structural ring coloring + text contrast
    # ======================================================================

    # Apply structural styling only if a trace was successfully created.
    if fig.data:
        trace = fig.data[0]

        # Extract label hierarchy and current background colors.
        labels = list(trace["labels"])
        parents = list(trace["parents"])
        colors = list(trace["marker"]["colors"])

        # Define root identifier (book name).
        root_label = str(book_name)

        # Retrieve STRUCTURE level color for first ring (figures).
        structure_hex = levels.get("STRUCTURE")
        if not structure_hex:
            print("Missing GLOBAL_VISUAL_STYLE['colors']['levels']['STRUCTURE'] – cannot style figure ring.")
            return

        # Semi-transparent STRUCTURE color for figure ring.
        fig_ring_color = hex_color_to_rgba(structure_hex, 0.55)

        # Retrieve CORE level color for root (book node).
        core_hex = levels.get("CORE")
        if not core_hex:
            print("Missing GLOBAL_VISUAL_STYLE['colors']['levels']['CORE'] – cannot style root.")
            return

        root_color = core_hex

        # Iterate over segments and enforce structural ring coloring.
        for i, (lab, par) in enumerate(zip(labels, parents)):

            # Root node (center of visualization).
            if str(lab) == root_label and par == "":
                colors[i] = root_color

            # First ring (figures directly under root).
            elif str(par) == root_label:
                colors[i] = fig_ring_color

        # Apply finalized background colors to trace.
        trace["marker"]["colors"] = colors
        apply_accessible_text_colors(
            trace,
            colors,
            levels.get("NEUTRAL_TEXT"),
            levels.get("LIGHT_TEXT"),
        )

        # Harmonized border styling between segments.
        border_rgba = hex_color_to_rgba(levels.get("NEUTRAL_TEXT"), 0.22)

        trace.update(
            marker=dict(
                line=dict(
                    color=border_rgba,
                    width=0.8,
                )
            )
        )

    # ======================================================================
    # [6] Global layout styling
    # ======================================================================

    # Apply centralized visual style (fonts, layout defaults).
    apply_global_visual_style(fig, has_axes=False)

    # Hide axis labels (Sunburst is axis-less but helper keeps API consistent).
    apply_global_visual_visibility(fig, show_axis_labels=False)

    # Set title and margins.
    fig.update_layout(
        title={
        "text": f"Sunburst – {book_name} (book overview)",
        "pad": {"t": 10},        },
        margin = {"t": 80, "l": 20, "r": 20, "b": 20},
    )

    # ======================================================================
    # [7] Export handling (delegated)
    # ======================================================================

    # Define output directory inside project data structure.
    output_dir = os.path.join("data", book_name, "visualization")

    variant_label = "sunburst_work"
    filename_stub = f"viz_{variant_label}_{book_name}"
    filename = f"{filename_stub}.html"

    # Delegate export/show logic to centralized helper.
    export_visualization_output(
        fig,
        paths=paths,
        book_name=book_name,
        filename=filename,
        export_func=apply_global_visual_modebar_export,
        filename_stub=filename_stub,
        output_dir=output_dir
    )

def build_sunburst_data_types_lemma(df, cols, figure_name):
    """
    Build aggregated data for a figure-centered Sunburst (Type → Lemma).

    The selected figure serves as the center node. The resulting structure
    corresponds to:

        center_figure → type_group → lemma

    Data scope and assumptions:
        - Expects a normalized DataFrame and column mapping as returned by
          prepare_naming_data(...).
        - Uses naming variant columns (excluding the unnormalized base column
          "Bezeichnung") and epithet columns as defined in cols.
        - Assumes canonical, pre-categorized naming data; no additional
          fuzzy matching or similarity heuristics are applied here.

    Aggregation logic:
        - Rows are filtered to those where target == figure_name.
        - Within each row:
            * Naming variants are counted once per lemma (row-internal dedup).
            * Epithets are counted once per lemma (row-internal dedup).
            * A lemma may be counted once as naming variant and once as epithet
              within the same row (category-level separation).
        - Naming variants are internally classified as:
            * "Proper name" if matched via resolve_name_lemmas_for_figure(...)
            * otherwise "Antonomasia"
        - Epithets are always assigned to type "Epithets".

    Output structure:
        Returns a pandas DataFrame with one row per (type, lemma) combination
        containing:
            - center_figure (str)
            - type_group (ring-level grouping: "Naming variants" / "Epithets")
            - color_group (leaf-level category for color mapping)
            - lemma (str)
            - frequency (int)

    Guards and behavior:
        - Raises ValueError if the required 'target' column mapping is missing.
        - Returns an empty DataFrame if no matching rows exist for the figure.

    Parameters:
        df (pd.DataFrame):
            Normalized naming data.
        cols (dict):
            Column mapping dictionary produced by prepare_naming_data(...).
        figure_name (str):
            Identifier of the selected center figure.

    Returns:
        pd.DataFrame
            Aggregated data suitable for px.sunburst with
            path=["center_figure", "type_group", "lemma"].
    """
    # ----------------------------------------------------------------------
    # Resolve required column mappings
    # ----------------------------------------------------------------------

    # Target column identifies the named figure in each row.
    target_col = cols.get("target")

    # Use naming variant columns as provided by prepare_naming_data(...),
    # but explicitly exclude the unnormalized base column "Bezeichnung".
    naming_variant_cols_all = cols.get("naming_variant_cols", [])
    naming_variant_cols = [
        c for c in naming_variant_cols_all
        if str(c).strip().lower() != "bezeichnung"
    ]

    # Epithet columns (already normalized via prepare_naming_data).
    epithet_cols = cols.get("epithet_cols", [])

    # Guard: required column mapping must be present.
    if target_col is None:
        raise ValueError("Column mapping 'target' is missing in cols.")

    # ----------------------------------------------------------------------
    # Determine proper-name lemmas for the selected figure
    # ----------------------------------------------------------------------

    # Resolve which lemmas count as "Proper name" for this figure.
    # The helper encapsulates name-matching logic.
    lemmas_matched_as_proper_name = resolve_name_lemmas_for_figure(df, cols, figure_name)

    # ----------------------------------------------------------------------
    # Filter DataFrame to rows where the selected figure is the target
    # ----------------------------------------------------------------------

    # Restrict analysis to rows referring to the selected center figure.
    dff = df[df[target_col].astype(str).str.strip() == str(figure_name).strip()].copy()

    # Reset index for stable iteration (index itself is not semantically used).
    dff = dff.reset_index(drop=True)

    # ----------------------------------------------------------------------
    # Aggregate counts per (type_label, lemma)
    # ----------------------------------------------------------------------

    counts = Counter()

    for _, row in dff.iterrows():

        # Deduplicate per row separately for naming variants and epithets.
        used_lemmas_naming = set()
        used_lemmas_epithets = set()

        # ------------------------------
        # Naming variants (Bezeichnung 1–4)
        # ------------------------------

        for col in naming_variant_cols:
            val = row.get(col)
            if not isinstance(val, str):
                continue

            lemma = val.strip()
            if not lemma or lemma in used_lemmas_naming:
                continue

            used_lemmas_naming.add(lemma)

            # Classify naming variant internally as Proper name or Antonomasia.
            if lemma in lemmas_matched_as_proper_name:
                type_label = "Proper name"
            else:
                type_label = "Antonomasia"

            counts[(type_label, lemma)] += 1

        # ------------------------------
        # Epithets
        # ------------------------------

        for col in epithet_cols:
            val = row.get(col)
            if not isinstance(val, str):
                continue

            lemma = val.strip()
            if not lemma or lemma in used_lemmas_epithets:
                continue

            used_lemmas_epithets.add(lemma)

            counts[("Epithets", lemma)] += 1

    # ----------------------------------------------------------------------
    # Build output DataFrame for Sunburst visualization
    # ----------------------------------------------------------------------

    rows = []

    # Map leaf-level types to ring-level grouping (Ring 1).
    type_group_map = {
        "Proper name": "Naming variants",
        "Antonomasia": "Naming variants",
        "Epithets": "Epithets",
    }

    for (type_label, lemma), freq in counts.items():
        type_group = type_group_map.get(type_label, "Epithets")
        color_group = type_label
        safe_lemma = f"{lemma} " if lemma == figure_name else lemma
        rows.append(
            {
                "center_figure": figure_name,
                "type_group": type_group,
                "color_group": color_group,
                "lemma": safe_lemma,
                "frequency": freq,
            }
        )

    # Return DataFrame suitable for px.sunburst with
    # path=["center_figure", "type_group", "lemma"].
    return pd.DataFrame(rows)

def build_sunburst_data_namer_lemma(df, cols, figure_name):
    """
    Figure-centered aggregation for Namer → Lemma Sunburst visualization.

    This function prepares aggregated frequency data for the
    Namer-centered Sunburst mode. The hierarchical structure is:

        center_figure → namer → lemma

    Internal type labels ("Proper name", "Antonomasia", "Epithets")
    are retained for semantic classification and color mapping,
    but are not used as an explicit hierarchy level in the Sunburst path.

    Data assumptions:
        - The DataFrame `df` has already been normalized via
          prepare_naming_data(...).
        - Column mappings (target, namer, naming_variant_cols,
          epithet_cols) are provided via `cols`.
        - Only normalized designation columns (e.g. "Bezeichnung 1–4")
          are considered; the unnormalized base column "Bezeichnung"
          is explicitly excluded.

    Aggregation logic:
        - Rows are filtered to those where the target figure matches
          `figure_name`.
        - Naming variants and epithets are counted separately per row,
          with per-row deduplication to avoid artificial inflation.
        - For naming variants, lemmas are classified as:
              * "Proper name" (via resolve_name_lemmas_for_figure and
                optional name-matching fallback), or
              * "Antonomasia".
        - Epithets are always labeled as "Epithets".
        - Frequencies are aggregated per (namer, type_label, lemma).

    Namer resolution:
        - Primary source: column mapped as `namer`.
        - Fallback: narrator column (if available), resolved either via
          explicit mapping in `cols` or heuristic column-name matching.
        - Rows without a resolvable namer are skipped.

    Output structure:
        Returns a DataFrame with the following columns:
            - center_figure
            - namer
            - type_group      (ring-level grouping: Naming variants / Epithets)
            - color_group     (leaf-level color key)
            - type            (internal type label)
            - lemma
            - frequency

        The resulting DataFrame is intended for use with:
            px.sunburst(..., path=["center_figure", "namer", "lemma"])

    Parameters:
        df (pd.DataFrame):
            Normalized naming dataset.

        cols (dict):
            Column mapping dictionary produced by prepare_naming_data(...).

        figure_name (str):
            The selected center figure for which the aggregation is performed.

    Returns:
        pd.DataFrame:
            Aggregated frequency table suitable for Sunburst visualization.
    """
    # ----------------------------------------------------------------------
    # Resolve required column mappings
    # ----------------------------------------------------------------------

    # Column identifying the named figure in each row.
    target_col = cols.get("target")

    # Column identifying the naming entity (namer).
    namer_col = cols.get("namer")

    # Use normalized naming variant columns as provided by prepare_naming_data(...),
    # but explicitly exclude the unnormalized base column "Bezeichnung".
    naming_variant_cols_all = cols.get("naming_variant_cols", [])
    naming_variant_cols = [
        c for c in naming_variant_cols_all
        if str(c).strip().lower() != "bezeichnung"
    ]

    # Normalized epithet columns.
    epithet_cols = cols.get("epithet_cols", [])

    # Guard: required column mappings must exist.
    if target_col is None:
        raise ValueError("Column mapping 'target' is missing in cols.")
    if namer_col is None:
        raise ValueError("Column mapping 'namer' is missing in cols.")

    # ----------------------------------------------------------------------
    # Resolve optional narrator fallback column
    # ----------------------------------------------------------------------

    # Try to determine a narrator column via explicit mapping in `cols`.
    narrator_col = None
    for key in ("narrator", "narrator_col"):
        if key in cols:
            narrator_col = cols[key]
            break

    # If not explicitly mapped, try heuristic column-name matching.
    if narrator_col is None:
        for c in df.columns:
            if str(c).strip().lower() in ("erzähler", "erzaehler", "narrator"):
                narrator_col = c
                break

    # ----------------------------------------------------------------------
    # Determine proper-name lemmas for the selected figure
    # ----------------------------------------------------------------------

    # Resolve canonical proper-name lemmas for classification of naming variants.
    name_lemmas = resolve_name_lemmas_for_figure(df, cols, figure_name)

    # ----------------------------------------------------------------------
    # Filter DataFrame to rows where the selected figure is the target
    # ----------------------------------------------------------------------

    # Restrict analysis to rows referring to the selected center figure.
    dff = df[df[target_col].astype(str).str.strip() == str(figure_name).strip()].copy()

    # Reset index for stable iteration.
    dff = dff.reset_index(drop=True)

    # ----------------------------------------------------------------------
    # Aggregate counts per (namer, type_label, lemma)
    # ----------------------------------------------------------------------

    counts = Counter()
    namer_row_counts: Counter = Counter()

    for _, row in dff.iterrows():

        # Resolve namer from primary column.
        raw_namer = row.get(namer_col)
        namer = raw_namer.strip() if isinstance(raw_namer, str) else ""

        # Fallback: use narrator column if namer is empty.
        if not namer and narrator_col is not None:
            raw_narr = row.get(narrator_col)
            namer = raw_narr.strip() if isinstance(raw_narr, str) else ""

        # Skip rows without a resolvable namer.
        if not namer:
            continue

        # Track row count per namer (one row = one mention)
        namer_row_counts[namer] += 1

        # Deduplicate per row separately for naming variants and epithets.
        used_naming_lemmas = set()
        used_epithet_lemmas = set()

        # ------------------------------
        # Naming variants (Bezeichnung 1–4)
        # ------------------------------

        for col in naming_variant_cols:
            val = row.get(col)
            if not isinstance(val, str):
                continue
            lemma = val.strip()
            if not lemma or lemma in used_naming_lemmas:
                continue
            used_naming_lemmas.add(lemma)

            # Default classification for naming variants.
            type_label = "Antonomasia"

            # Canonical proper-name resolution.
            if lemma in name_lemmas:
                type_label = "Proper name"
            else:
                # Optional heuristic fallback.
                try:
                    if match_name_to_lemma(figure_name, lemma, aliases=None):
                        type_label = "Proper name"
                except (TypeError, ValueError, AttributeError):
                    pass

            counts[(namer, type_label, lemma)] += 1

        # ------------------------------
        # Epithets
        # ------------------------------

        type_label = "Epithets"
        for col in epithet_cols:
            val = row.get(col)
            if not isinstance(val, str):
                continue

            lemma = val.strip()
            if not lemma or lemma in used_epithet_lemmas:
                continue

            used_epithet_lemmas.add(lemma)

            counts[(namer, type_label, lemma)] += 1

    # ----------------------------------------------------------------------
    # Build output DataFrame for Sunburst visualization
    # ----------------------------------------------------------------------

    rows = []

    # Map internal type_label to ring-level grouping and color key.
    type_group_map = {
        "Proper name": "Naming variants",
        "Antonomasia": "Naming variants",
        "Epithets": "Epithets",
    }
    color_group_map = {
        "Proper name": "Proper name",
        "Antonomasia": "Antonomasia",
        "Epithets": "Epithets",
    }

    for (namer, type_label, lemma), freq in counts.items():
        type_group = type_group_map.get(type_label, "Epithets")
        color_group = color_group_map.get(type_label, "Epithets")

        rows.append(
            {
                "center_figure": figure_name,
                "namer": namer,
                "type_group": type_group,
                "color_group": color_group,
                "type": type_label,
                "lemma": lemma,
                "frequency": freq,
                "namer_row_count": namer_row_counts.get(namer, 0),
            }
        )

    # Return DataFrame suitable for px.sunburst with
    # path=["center_figure", "namer", "lemma"].
    return pd.DataFrame(rows)

def build_sunburst_data_book_overview(df, cols, top_k, book_name):
    """
    Book-centered aggregation for Book → Figure → Type Sunburst visualization.

    This function prepares entry-based frequency data for the
    book-centered (book overview) Sunburst mode. The hierarchical structure is:

        book_name → figure → type

    Aggregation model (entry-based semantics):
        - Each row in `df` represents one mention (entry) of a figure.
        - `total_entries_per_figure` counts how many entries mention a figure.
        - For each entry, the presence of naming types is detected:
              * "Proper name"
              * "Antonomasia"
              * "Epithets"
          Each type is counted at most once per entry (presence-based, not token-based).
        - Type frequencies therefore represent:
              “In how many entries of this figure does this type occur?”

    Naming variant handling:
        - Only normalized naming variants columns (e.g. "Bezeichnung 1–4")
          are considered.
        - The unnormalized base column "Bezeichnung" is explicitly excluded.
        - Proper names are determined via match_name_to_lemma(...)
          (heuristic matching is intentionally retained in BETA).

    Top-K selection:
        - Figures are ranked by total entry count
          (i.e. number of mentions in the dataset).
        - If `top_k` is provided and > 0, only the top K figures are included.
        - If `top_k` is None or ≤ 0, all figures are included.

    Percentage semantics:
        - `pct_of_figure` is calculated relative to
          `total_entries_per_figure[figure]`.
        - It represents the share of entries of a figure
          in which the respective type is present.
        - This is not token-frequency share, but entry-share.

    Output structure:
        Returns a DataFrame with the following columns:
            - root                (book_name)
            - figure              (named figure)
            - type                (Proper name / Antonomasia / Epithets)
            - value               (entry-based presence count)
            - total_for_figure    (total entry count for figure)
            - pct_of_figure       (entry-share of this type)

        The resulting DataFrame is intended for use with:
            px.sunburst(..., path=["root", "figure", "type"])

    Parameters:
        df (pd.DataFrame):
            Normalized naming dataset (prepared via prepare_naming_data).

        cols (dict):
            Column mapping dictionary.

        top_k (int | None):
            Number of top figures to include (based on entry count).

        book_name (str):
            Identifier of the current book (used as root node label).

    Returns:
        pd.DataFrame:
            Aggregated frequency table suitable for book-centered
            Sunburst visualization.
    """
    # ----------------------------------------------------------------------
    # Resolve required column mappings
    # ----------------------------------------------------------------------

    # Column identifying the named figure in each row.
    target_col = cols.get("target")

    # Use normalized naming variant columns provided by prepare_naming_data(...),
    # but explicitly exclude the unnormalized base column "Bezeichnung".
    naming_variant_cols_all = cols.get("naming_variant_cols", [])
    naming_variant_cols = [
        c for c in naming_variant_cols_all
        if str(c).strip().lower() != "bezeichnung"
    ]

    # Normalized epithet columns.
    epithet_cols = cols.get("epithet_cols", [])

    # Guard: required column mapping must exist.
    if target_col is None:
        raise ValueError("Column mapping 'target' is missing in cols.")

    # ----------------------------------------------------------------------
    # Entry-based counting per figure
    # ----------------------------------------------------------------------

    # total_entries_per_figure:
    # Counts how many entries (rows) mention a given figure.
    total_entries_per_figure = Counter()

    # type_presence_counts:
    # Counts in how many entries a given type is present
    # (at most once per type per entry).
    type_presence_counts = Counter()

    # ----------------------------------------------------------------------
    # Iterate over dataset (each row = one mention / entry)
    # ----------------------------------------------------------------------

    for _, row in df.iterrows():
        raw_target = row.get(target_col)

        # Normalize figure label (robust against non-string values).
        figure = raw_target.strip() if isinstance(raw_target, str) else ""
        if not figure:
            continue

        # Each row represents one mention of the figure.
        total_entries_per_figure[figure] += 1

        # Track which types are present in this entry
        # (presence-based, not token-frequency based).
        types_present_in_entry = set()

        # ------------------------------
        # Naming variants
        # ------------------------------

        for col in naming_variant_cols:
            val = row.get(col)
            if not isinstance(val, str):
                continue

            lemma = val.strip()
            if not lemma:
                continue

            # Default classification for naming variants.
            type_label = "Antonomasia"

            # Heuristic name matching
            try:
                if match_name_to_lemma(figure, lemma, aliases=None):
                    type_label = "Proper name"
            except (TypeError, ValueError, AttributeError):
                pass

            types_present_in_entry.add(type_label)

        # ------------------------------
        # Epithets
        # ------------------------------

        for col in epithet_cols:
            val = row.get(col)
            if not isinstance(val, str):
                continue

            lemma = val.strip()
            if not lemma:
                continue

            types_present_in_entry.add("Epithets")

        # Count presence per type (max 1 per type per entry).
        for type_label in types_present_in_entry:
            type_presence_counts[(figure, type_label)] += 1

    # ----------------------------------------------------------------------
    # Early return if no figures were detected
    # ----------------------------------------------------------------------

    if not total_entries_per_figure:
        return pd.DataFrame(
            columns=[
                "root",
                "figure",
                "type",
                "value",
                "total_for_figure",
                "pct_of_figure",
            ]
        )

    # ----------------------------------------------------------------------
    # Determine Top-K figures (based on entry count)
    # ----------------------------------------------------------------------

    sorted_figures = sorted(
        total_entries_per_figure.items(),
        key=lambda item: (-item[1], str(item[0]).lower()),
    )

    if top_k is not None and top_k > 0:
        top_figures = {name for name, _ in sorted_figures[:top_k]}
    else:
        top_figures = {name for name, _ in sorted_figures}

    # ----------------------------------------------------------------------
    # Build output rows (book → figure → type)
    # ----------------------------------------------------------------------

    rows = []

    for (figure, type_label), count in type_presence_counts.items():
        if figure not in top_figures:
            continue

        # Entry-based total per figure.
        total = total_entries_per_figure.get(figure, 0)

        # Share of entries of this figure in which the type is present.
        pct = count / total if total > 0 else 0.0

        rows.append(
            {
                "root": book_name,
                "figure": figure,
                "type": type_label,
                "value": count,
                "total_for_figure": total,
                "pct_of_figure": pct,
            }
        )

    # Return DataFrame suitable for px.sunburst with
    # path=["root", "figure", "type"].
    return pd.DataFrame(rows)

# =============================================================================
# CLI INPUT AND VALIDATION HELPERS
# =============================================================================

def ask_valid_figure_name(json_path: str) -> str | None:
    """
    Prompt the user for a figure name until it can be resolved successfully.

    This function loads categorization data from the provided JSON path
    and repeatedly asks the user to enter a figure name. The input is
    validated and resolved using `resolve_figure_name(...)`.

    Behavior
    --------
    - Loads entries from the categorization JSON file.
    - Prompts the user for non-empty input.
    - Delegates name resolution (exact match or fuzzy suggestion)
      to `resolve_figure_name(...)`.
    - Repeats until a valid figure name is returned.

    Side effects
    ------------
    - Reads from stdin (interactive input loop).
    - Prints guidance and validation messages to stdout.
    - Depends on `safe_read_json(...)` for loading data.

    Parameters
    ----------
    json_path : str
        Path to the categorization JSON file containing entries
        with the key "Benannte Figur".

    Returns
    -------
    str | None
        A resolved canonical figure name (as present in the data).

        In practice, this function does not terminate with None
        under normal control flow because the loop continues until
        a valid resolution occurs.
    """
    # Load categorization entries.
    # If the file cannot be read, default to an empty list.
    entries = safe_read_json(json_path, default=[])

    # Interactive retry loop: continues until a valid figure is resolved.
    while True:
        # Prompt user for figure name and remove surrounding whitespace.
        raw = input("Please enter the figure name:\n> ").strip()

        # Reject empty input explicitly.
        if not raw:
            print("Input cannot be empty.")
            continue

        # Attempt resolution via exact match or fuzzy suggestion.
        resolved = resolve_figure_name(raw, entries)

        # Successful resolution → return canonical figure name.
        if resolved is not None:
            return resolved

        # Resolution failed or was rejected → inform user and retry.
        print("No matching figure found. Please try again.")

    # Defensive return for static type checking (unreachable in practice).
    return None

# =============================================================================
# GLOBAL VISUALIZATION CONFIGURATION AND EXPORT
# =============================================================================

GLOBAL_VISUAL_STYLE: dict[str, Any] = {
    "typography": {
        "font_family": "Times New Roman",
        "base_size": 18,
        "tick_size": 36,
        "legend_size": 18,
    },
    "background": {
        "transparent": "rgba(0,0,0,0)",
    },
    "layout": {
        "margins": {"l": 60, "r": 30, "t": 20, "b": 60},
        "show_title_default": True,
        "show_legend_default": True,
        "show_axis_labels_default": True,
        "title_x": 0.5,
        "title_xanchor": "center",
    },
    "export": {
        "dpi": 300,
        "formats": ("png", "svg", "html"),
    },
    "colors": {
        "categories": {
            # Naming variants (group level)
            "Naming variants": "#8C6A4A",

            # Naming variants (subcategories)
            "Proper name": "#EFE4D4",
            "Antonomasia": "#F9C691",

            # Counterpole
            "Epithets": "#2F4A6D",

            # Epitheta (leaf/lexeme level) — optional but useful for consistent leaf styling
            "Epithets (lexeme level)": "#6A97B8",
        },

        "levels": {
            "CORE": "#0D1E26",
            "STRUCTURE": "#A6B4A0",
            "NEUTRAL_TEXT": "#2D2926",
            "AUXILIARY": "#A3A39A",
            "LIGHT_TEXT": "#F4F6F6",
        },
    },
}


def apply_global_visual_style(fig, *, tick_font_size=None, show_grid=None, has_axes: bool = True):
    """
    Apply centralized visual defaults to a Plotly figure.

    This helper enforces the project-wide presentation layer defined in
    GLOBAL_VISUAL_STYLE. It standardizes typography, background behavior,
    margins, and legend styling while remaining strictly presentation-only:
    no data transformations, semantic encodings, or plot-specific logic
    are introduced here.

    Behavior:
        - Applies global font family, base font size, and neutral text color.
        - Sets transparent paper and plot backgrounds (poster-ready output).
        - Applies standardized margins from the layout configuration.
        - Normalizes legend typography.
        - Optionally updates axis styling (ticks, titles, grid) if the
          figure contains axes.

    Axis handling:
        - If `has_axes` is True, x- and y-axes are updated using global
          typography and auxiliary grid colors.
        - If `has_axes` is False (e.g., Sunburst, Pie, network plots),
          axis updates are skipped to avoid unintended side effects.
        - `show_grid` defaults to the configuration value unless explicitly set.

    Parameters:
        fig (plotly.graph_objects.Figure):
            The figure to be styled in place.

        tick_font_size (int | None):
            Optional override for axis tick font size.
            If None, the global default is used.

        show_grid (bool | None):
            Whether grid lines should be shown (axes only).
            If None, the configured default is applied.

        has_axes (bool):
            Indicates whether the figure contains standard Cartesian axes.
            Defaults to True.

    Returns:
        plotly.graph_objects.Figure:
            The same figure instance, styled in place.
    """
    style: dict[str, Any] = GLOBAL_VISUAL_STYLE

    typography: dict[str, Any] = style["typography"]
    layout_cfg: dict[str, Any] = style["layout"]
    levels: dict[str, Any] = style["colors"]["levels"]

    tick_size = tick_font_size if tick_font_size is not None else typography["tick_size"]

    fig.update_layout(
        font={
            "family": typography["font_family"],
            "size": typography["base_size"],
            "color": levels["NEUTRAL_TEXT"],
        },
        paper_bgcolor=style["background"]["transparent"],
        plot_bgcolor=style["background"]["transparent"],
        margin=layout_cfg["margins"],
        legend={
            "font": {"family": typography["font_family"], "size": typography["legend_size"]}
        },
    )

    if show_grid is None:
        show_grid = layout_cfg.get("show_grid_default", True)

    if has_axes:
        fig.update_xaxes(
            tickfont={"family": typography["font_family"], "size": tick_size, "color": levels["NEUTRAL_TEXT"]},
            title_font={"family": typography["font_family"], "size": typography["base_size"],
                        "color": levels["NEUTRAL_TEXT"]},
            showgrid=show_grid,
            gridcolor=levels["AUXILIARY"],
            zeroline=False,
        )
        fig.update_yaxes(
            tickfont={"family": typography["font_family"], "size": tick_size, "color": levels["NEUTRAL_TEXT"]},
            title_font={"family": typography["font_family"], "size": typography["base_size"],
                        "color": levels["NEUTRAL_TEXT"]},
            showgrid=show_grid,
            gridcolor=levels["AUXILIARY"],
            zeroline=False,
        )

    return fig

def apply_global_visual_visibility(fig, *, show_title=None, show_legend=None, show_axis_labels=None):
    """
    Apply standardized visibility controls to a Plotly figure.

    This helper centralizes the toggling of common layout elements
    (title, legend, and axis labels) according to the global
    configuration in GLOBAL_VISUAL_STYLE["layout"].

    It operates purely at the presentation layer and must not
    modify data, traces, or semantic encodings.

    Behavior:
        - Each visibility flag (show_title, show_legend, show_axis_labels)
          can be explicitly set or left as None.
        - If a flag is None, the corresponding default from the global
          layout configuration is applied.
        - Title positioning (x, xanchor) is standardized when enabled.
        - Axis label titles are removed when show_axis_labels is False.
        - Legend visibility is controlled via layout.showlegend.

    Parameters:
        fig (plotly.graph_objects.Figure):
            The figure to be updated in place.

        show_title (bool | None):
            Whether the figure title should be displayed.
            If None, the global default is used.

        show_legend (bool | None):
            Whether the legend should be displayed.
            If None, the global default is used.

        show_axis_labels (bool | None):
            Whether axis titles should be displayed.
            If None, the global default is used.

    Returns:
        plotly.graph_objects.Figure:
            The same figure instance, updated in place.
    """
    layout_cfg: dict[str, Any] = GLOBAL_VISUAL_STYLE["layout"]

    if show_title is None:
        show_title = layout_cfg["show_title_default"]
    if show_legend is None:
        show_legend = layout_cfg["show_legend_default"]
    if show_axis_labels is None:
        show_axis_labels = layout_cfg["show_axis_labels_default"]

    if not show_title:
        fig.update_layout(title=None)
    else:
        fig.update_layout(title={"x": layout_cfg["title_x"], "xanchor": layout_cfg["title_xanchor"]})

    fig.update_layout(showlegend=show_legend)

    if not show_axis_labels:
        fig.update_xaxes(title_text=None)
        fig.update_yaxes(title_text=None)

    return fig

def build_global_visual_export_filename(prefix: str = "viz") -> str:
    """
    Generate a timestamp-based filename stub for visualization exports.

    This helper creates a deterministic, time-stamped filename prefix
    for Plotly Modebar exports. No file extension is included; the
    calling export routine is responsible for appending the appropriate
    extension (e.g.: .html, .png, .svg).

    Format:
        {prefix}_YYYY_MM_DD_HHMM

    The timestamp reflects the current local system time at the moment
    of invocation.

    Parameters:
        prefix (str):
            Leading identifier for the export (default: "viz").
            Used to distinguish visualization types or contexts.

    Returns:
        str:
            Filename stub without extension.
    """
    return f"{prefix}_{datetime.now():%Y_%m_%d_%H%M}"

def apply_global_visual_modebar_export(
        fig,
        output_path: str | Path,
        *,
        filename_stub: str | None = None,
) -> None:
    """
    Export a Plotly figure to an interactive HTML file with a standardized
    in-page export overlay.

    This helper writes a self-contained (Plotly JS via CDN) HTML document that:
        - preserves full interactivity (zoom, pan, hover, modebar),
        - replaces the default Plotly "toImage" camera action with two explicit,
          globally standardized download buttons:
              * SVG export
              * PNG export (A4 landscape, 300 dpi target dimensions)

    Export behavior inside the HTML:
        - The overlay buttons are injected via a post-render JavaScript hook.
        - During image export, Plotly updatemenus (e.g., dropdown filters) are
          temporarily removed to prevent them from appearing in exported images,
          then restored immediately afterward.
        - The overlay is positioned dynamically relative to the modebar so it
          does not overlap UI controls; positioning updates on resize and after
          Plotly redraw events.

    Scope and semantics:
        - This function affects only the exported HTML (client-side behavior).
        - No changes are applied to the underlying figure object beyond
          serialization for HTML output.
        - The exported HTML remains interactive after any download operation.

    Parameters:
        fig (plotly.graph_objects.Figure):
            Plotly figure to export.

        output_path (str | pathlib.Path):
            Destination path for the HTML file. Parent directories are created
            if missing.

        filename_stub (str | None):
            Base filename used by the in-page download buttons (no extension).
            If None, a timestamp-based stub is generated via
            build_global_visual_export_filename(...).

    Returns:
        None
            Writes the HTML file to disk.
    """
    out_path = Path(output_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    if filename_stub is None:
        filename_stub = build_global_visual_export_filename("viz")

    # A4 landscape @ 300 dpi
    png_width = 3508
    png_height = 2480

    config: dict[str, Any] = {
        "displaylogo": False,
        "responsive": True,
        "modeBarButtonsToRemove": ["toImage"],  # remove default camera button
    }

    post_script = f"""
(function () {{
  function whenPlotIsReady(cb) {{
    var tries = 0;
    var maxTries = 200; // ~20s @ 100ms
    var timer = setInterval(function () {{
      tries += 1;
      var gd = document.getElementById('{{plot_id}}');
      if (gd && gd._fullLayout) {{
        clearInterval(timer);
        cb(gd);
        return;
      }}
      if (tries >= maxTries) {{
        clearInterval(timer);
        console.warn("Plot not ready – custom export overlay was not added.");
      }}
    }}, 100);
  }}

  function getUpdatemenus(gd) {{
    var lay = (gd && gd.layout) ? gd.layout : null;
    return (lay && lay.updatemenus) ? lay.updatemenus : [];
  }}

  function setUpdatemenus(gd, menus) {{
    return Plotly.relayout(gd, {{ 'updatemenus': menus }});
  }}

  function hideMenusThenDownload(gd, dlOpts) {{
    var currentMenus = getUpdatemenus(gd);

    var oldMargin = (gd.layout && gd.layout.margin) ? gd.layout.margin : {{}};
    var oldTitle  = (gd.layout && gd.layout.title)  ? gd.layout.title  : null;

    var newMargin = Object.assign({{}}, oldMargin, {{
      t: Math.max(oldMargin.t || 0, 120)
    }});

    var newTitle = oldTitle
      ? Object.assign({{}}, oldTitle, {{ pad: {{ t: 10 }} }})
      : oldTitle;

    return Plotly.relayout(gd, {{
        updatemenus: [],
        margin: newMargin,
        title: newTitle
      }})
      .then(function () {{
        return Plotly.downloadImage(gd, dlOpts);
      }})
      .then(function () {{
        return Plotly.relayout(gd, {{
          updatemenus: currentMenus,
          margin: oldMargin,
          title: oldTitle
        }});
      }})
      .catch(function (err) {{
        try {{
          Plotly.relayout(gd, {{
            updatemenus: currentMenus,
            margin: oldMargin,
            title: oldTitle
          }});
        }} catch (e) {{}}
        console.error(err);
      }});
  }}

  function ensureContainerPositioning(container) {{
    // Ensure absolute positioning works (Plotly already uses relative in most cases,
    // but make it explicit and safe).
    var cs = window.getComputedStyle(container);
    if (!cs.position || cs.position === "static") {{
      container.style.position = "relative";
    }}
  }}

  function positionOverlay(container, wrap, gd) {{
    // Position overlay to the LEFT of the modebar (no overlap), dynamic per viewport.
    var modebar = container.querySelector('.modebar');

    // If no modebar found yet, keep a conservative fallback.
    if (!modebar) {{
      wrap.style.top = "8px";
      wrap.style.right = "160px";
      return;
    }}

    var mb = modebar.getBoundingClientRect();
    var c = container.getBoundingClientRect();

    // right(px) = containerRight - modebarLeft + gap
    var gap = 10;
    var rightPx = Math.max(8, (c.right - mb.left) + gap);

    wrap.style.top = "8px";
    wrap.style.right = rightPx + "px";
  }}

  function addOverlayButtons(gd) {{
    // Prevent duplicates
    if (gd.__globalExportOverlayAdded) return;
    gd.__globalExportOverlayAdded = true;

    var container = gd; // graph div
    ensureContainerPositioning(container);

    var wrap = document.createElement("div");
    wrap.style.position = "absolute";
    wrap.style.top = "8px";
    wrap.style.right = "160px"; // fallback, will be recomputed
    wrap.style.zIndex = "9999";
    wrap.style.display = "flex";
    wrap.style.gap = "8px";
    wrap.style.alignItems = "center";

    function makeBtn(label) {{
      var b = document.createElement("button");
      b.type = "button";
      b.textContent = label;
      b.style.fontFamily = "Times New Roman, serif";
      b.style.fontSize = "12px";
      b.style.padding = "4px 8px";
      b.style.border = "1px solid rgba(0,0,0,0.25)";
      b.style.borderRadius = "4px";
      b.style.background = "rgba(255,255,255,0.95)";
      b.style.cursor = "pointer";
      b.onmouseenter = function () {{ b.style.background = "rgba(245,245,245,0.98)"; }};
      b.onmouseleave = function () {{ b.style.background = "rgba(255,255,255,0.95)"; }};
      return b;
    }}

    var btnSvg = makeBtn("Download plot as SVG");
    btnSvg.addEventListener("click", function () {{
      hideMenusThenDownload(gd, {{
        format: "svg",
        filename: "{filename_stub}"
      }});
    }});

    var btnPng = makeBtn("Download plot as PNG (A4 landscape, 300 dpi)");
    btnPng.addEventListener("click", function () {{
      var targetW = {png_width};
      var targetH = {png_height};

      // current plot size (as rendered in the browser)
      var currentW = (gd && gd._fullLayout && gd._fullLayout.width) ? gd._fullLayout.width : 1169;
      var currentH = (gd && gd._fullLayout && gd._fullLayout.height) ? gd._fullLayout.height : 827;

      // scale to fit inside A4@300dpi while preserving proportions
      var s = Math.min(targetW / currentW, targetH / currentH);
      if (!isFinite(s) || s <= 0) s = 3;
      if (s > 8) s = 8;

      var oldPaperBg = (gd.layout && gd.layout.paper_bgcolor) ? gd.layout.paper_bgcolor : null;
      var oldPlotBg  = (gd.layout && gd.layout.plot_bgcolor)  ? gd.layout.plot_bgcolor  : null;

      Plotly.relayout(gd, {{
        paper_bgcolor: "white",
        plot_bgcolor:  "white"
      }}).then(function () {{
        return hideMenusThenDownload(gd, {{
          format: "png",
          filename: "{filename_stub}",
          scale: 3
        }});
      }}).then(function () {{
        return Plotly.relayout(gd, {{
          paper_bgcolor: oldPaperBg !== null ? oldPaperBg : "rgba(0,0,0,0)",
          plot_bgcolor:  oldPlotBg  !== null ? oldPlotBg  : "rgba(0,0,0,0)"
        }});
      }}).catch(function (err) {{
        Plotly.relayout(gd, {{
          paper_bgcolor: oldPaperBg !== null ? oldPaperBg : "rgba(0,0,0,0)",
          plot_bgcolor:  oldPlotBg  !== null ? oldPlotBg  : "rgba(0,0,0,0)"
        }});
        console.error(err);
      }});
    }});


    wrap.appendChild(btnSvg);
    wrap.appendChild(btnPng);
    container.appendChild(wrap);

    // Initial positioning (now that wrap is in DOM)
    positionOverlay(container, wrap, gd);

    // Reposition on resize
    window.addEventListener("resize", function () {{
      positionOverlay(container, wrap, gd);
    }});

    // Reposition after Plotly redraws (modebar can change layout)
    gd.on("plotly_afterplot", function () {{
      positionOverlay(container, wrap, gd);
    }});

    // Also poll a bit at start to catch late modebar creation
    var tries = 0;
    var t = setInterval(function () {{
      tries += 1;
      positionOverlay(container, wrap, gd);
      if (tries >= 30) clearInterval(t);
    }}, 150);
  }}

  whenPlotIsReady(function (gd) {{
    addOverlayButtons(gd);
  }});
}})();
"""

    html = pio.to_html(
        fig,
        full_html=True,
        include_plotlyjs="cdn",
        config=config,
        post_script=post_script,
    )
    out_path.write_text(html, encoding="utf-8")