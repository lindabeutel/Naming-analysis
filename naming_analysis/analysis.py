"""
analysis.py

This module provides the core analysis functionality for the naming analysis pipeline.
It includes:

- Interactive CLI menus for triggering various analysis types
- Generation of wordlists (by column or figure)
- Keyword analysis with reference corpora
- Collocation analysis and KWIC display
- Simple figure-specific visualization using plotly

All analytical operations are performed on categorized JSON data,
optionally filtered by figures or type.
"""
from __future__ import annotations

import os
import math
import csv
import difflib
import numpy as np
import uuid
import webbrowser
from collections import Counter
from datetime import datetime
from itertools import combinations
from pathlib import Path
from plotly.colors import n_colors
from typing import Any, List

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio

from naming_analysis.shared import (
    ask_user_choice,
    get_first_valid_text,
    hex_color_to_rgba,
    hex_color_to_rgb_tuple,
    pick_accessible_text_color,
    rgb_tuple_to_plotly_color
)
from naming_analysis.io_utils import safe_read_json
from naming_analysis.loaders import load_collocation_sheet, build_fallback_collocation_df_from_tei, load_naming_sources_with_excel_fallback
from naming_analysis.shared import prepare_naming_data, serialize_verse_value

def run_analysis_menu(config_data, paths, data, book_name):
    """
    Entry point for interactive analysis tasks.

    Offers the user a menu to select one of the following options:
    - Wordlist generation
    - Keyword analysis
    - Collocation extraction
    - Visualization
    - Exit

    Parameters:
        config_data (dict): Loaded configuration data.
        paths (dict): Dictionary of relevant file paths.
        data (dict): Loaded TEI and Excel data.
        book_name (str): Short identifier of the current book.
    """
    while True:
        print("📊 Which type of analysis do you want to perform?")
        print("[1] Wordlist")
        print("[2] Naming figure analysis")
        print("[3] Keywords")
        print("[4] Collocations")
        print("[5] Visualization")
        print("[6] Exit analysis")

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
            print("📦 Analysis completed.")
            break

def run_wordlist_menu(paths, book_name):
    """
    Interactive menu for generating wordlists from categorization data.

    The user can choose to extract:
    - All values from a selected column group
    - All naming variants for a given figure
    - All epithets for a given figure
    - A combined list of naming variants and epithets

    Parameters:
        paths (dict): Dictionary of file paths.
        book_name (str): Name of the current book for output file naming.
    """
    json_path = paths["categorization_json"]
    output_dir = os.path.join("data", book_name, "analysis")
    os.makedirs(output_dir, exist_ok=True)

    while True:
        print("\n📁 What kind of wordlist do you want to generate?")
        print("[1] All values from a column (e.g., 'Benannte Figur')")
        print("[2] All naming variants (Bezeichnungen) for a specific figure")
        print("[3] All epithets (Epitheta) for a specific figure")
        print("[4] Combined naming variants and epithets")
        print("[5] Back to main analysis menu")

        choice = ask_user_choice("> ", ["1", "2", "3", "4", "5"])

        if choice == "1":
            print("\n📑 Available column groups:")
            print("- Benannte Figur")
            print("- Bezeichnung")
            print("- Epitheta")

            column_input = input("Please enter one of the above column names:\n> ").strip()
            valid_columns = ["Benannte Figur", "Bezeichnung", "Epitheta"]

            while column_input not in valid_columns:
                print("⚠️ Invalid input. Please enter one of: Benannte Figur, Bezeichnung, Epitheta")
                column_input = input("> ").strip()

            filename = f"wordlist_{column_input}_{book_name}.csv".replace(" ", "_")
            output_path = os.path.join(output_dir, filename)
            generate_wordlist_by_column(column_input, json_path, output_path)

        elif choice == "2":
            figure = ask_valid_figure_name(paths["categorization_json"])
            if figure is None:
                return
            filename = f"wordlist_Bezeichnung_{figure}.csv".replace(" ", "_")
            output_path = os.path.join(output_dir, filename)
            generate_naming_variants_for_figure(figure, json_path, output_path)

        elif choice == "3":
            figure = ask_valid_figure_name(paths["categorization_json"])
            if figure is None:
                return
            filename = f"wordlist_Epitheta_{figure}.csv".replace(" ", "_")
            output_path = os.path.join(output_dir, filename)
            generate_epithets_for_figure(figure, json_path, output_path)

        elif choice == "4":
            figure = ask_valid_figure_name(paths["categorization_json"])
            if figure is None:
                return
            filename = f"wordlist_Combined_{figure}.csv".replace(" ", "_")
            output_path = os.path.join(output_dir, filename)
            generate_combined_naming_variants_epithets(figure, json_path, output_path)

        elif choice == "5":
            print("↩️ Returning to analysis menu.")
            return

def generate_wordlist_by_column(column_name: str, json_path: str, output_path: str):
    """
    Generates a frequency list from a single column group (e.g. "Bezeichnung", "Epitheta").

    The result is written to a CSV file.

    Parameters:
        column_name (str): Logical column group to analyze.
        json_path (str): Path to the categorization JSON file.
        output_path (str): Output path for the resulting CSV.
    """
    entries = safe_read_json(json_path, default=[])

    if column_name.lower() == "bezeichnung":
        columns = [f"Bezeichnung {i}" for i in range(1, 5)]
    elif column_name.lower() == "epitheta":
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

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["Wert", "Anzahl"])
        for value, count in most_common:
            writer.writerow([value, count])

    print(f"✅ Wordlist written to: {output_path}")

def resolve_figure_name(name: str, entries: list[dict]) -> str | None:
    """
    Attempts to resolve an input name to a known figure from categorization data.

    If no exact match is found, it suggests the closest match via fuzzy comparison.

    Parameters:
        name (str): The name entered by the user.
        entries (list[dict]): List of figure entries to compare against.

    Returns:
        str | None: The resolved name or None if rejected by the user.
    """
    all_names = {
        str(name).strip()
        for name in [e.get("Benannte Figur") for e in entries]
        if isinstance(name, str) and name.strip()
    }
    if name in all_names:
        return name

    suggestions = difflib.get_close_matches(name, all_names, n=1, cutoff=0.6)
    if suggestions:
        print(f'⚠️ Figure "{name}" not found.')
        print(f'❓ Did you mean "{suggestions[0]}"? [y/n]')
        answer = ask_user_choice("> ", ["y", "n"])
        if answer == "y":
            return suggestions[0]
        else:
            print("⚠️ No valid figure selected.")
            print("Please enter a valid name exactly as it appears in your categorization data.")
            return None
    else:
        print(f'⚠️ Figure "{name}" not found and no similar name could be suggested.')
        return None

def ask_valid_figure_name(json_path: str) -> str | None:
    """
    Repeatedly prompts the user to enter a figure name until it can be resolved.

    Parameters:
        json_path (str): Path to the categorization JSON file.

    Returns:
        str | None: A valid figure name, or None if resolution failed.
    """
    entries = safe_read_json(json_path, default=[])

    while True:
        raw = input("✍ Please enter the figure name:\n> ").strip()
        if not raw:
            print("⚠️ Input cannot be empty.")
            continue

        resolved = resolve_figure_name(raw, entries)
        if resolved is not None:
            return resolved

        print("⚠️ No matching figure found. Please try again.")

    return None  # ← für statische Typprüfung, wird nie erreicht

def generate_naming_variants_for_figure(figure_name: str, json_path: str, output_path: str):
    """
    Generates a frequency list of naming variants for a given figure.

    Parameters:
        figure_name (str): Name of the figure (already validated).
        json_path (str): Path to the categorization JSON file.
        output_path (str): Path to the output CSV.
    """
    entries = safe_read_json(json_path, default=[])
    # no need to resolve again – already handled
    resolved_name = figure_name

    filtered = [e for e in entries if e.get("Benannte Figur") == resolved_name]
    values = []
    for entry in filtered:
        for i in range(1, 5):
            val = entry.get(f"Bezeichnung {i}")
            if isinstance(val, str) and val.strip():
                values.append(val.strip())

    counts = Counter(values)
    most_common = counts.most_common()

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["Wert", "Anzahl"])
        for val, count in most_common:
            writer.writerow([val, count])

    print(f"✅ Wordlist for '{resolved_name}' written to: {output_path}")

def generate_epithets_for_figure(figure_name: str, json_path: str, output_path: str):
    """
    Generates a frequency list of epithets for a given figure.

    Parameters:
        figure_name (str): Name of the figure (already validated).
        json_path (str): Path to the categorization JSON file.
        output_path (str): Path to the output CSV.
    """
    entries = safe_read_json(json_path, default=[])
    # no need to resolve again – already handled
    resolved_name = figure_name

    filtered = [e for e in entries if e.get("Benannte Figur") == resolved_name]
    values = []
    for entry in filtered:
        for i in range(1, 6):
            val = entry.get(f"Epitheta {i}")
            if isinstance(val, str) and val.strip():
                values.append(val.strip())

    counts = Counter(values)
    most_common = counts.most_common()

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["Wert", "Anzahl"])
        for val, count in most_common:
            writer.writerow([val, count])

    print(f"✅ Wordlist for epithets of '{resolved_name}' written to: {output_path}")

def generate_combined_naming_variants_epithets(figure_name: str, json_path: str, output_path: str):
    """
    Generates a combined frequency list of all naming variants and epithets
    for a selected figure and saves it as CSV.

    Parameters:
        figure_name (str): Name of the target figure.
        json_path (str): Path to the categorization JSON file.
        output_path (str): Path to the output CSV.
    """
    entries = safe_read_json(json_path, default=[])
    # no need to resolve again – already handled
    resolved_name = figure_name

    filtered = [e for e in entries if e.get("Benannte Figur") == resolved_name]
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

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["Wert", "Anzahl"])
        for val, count in most_common:
            writer.writerow([val, count])

    print(f"✅ Combined wordlist for '{resolved_name}' written to: {output_path}")

def run_naming_figure_analysis(_config_data, paths, data, book_name):
    """
    Interactive CLI orchestration for 'Naming figure analysis' (validated inputs).
    Loads JSON/Excel via safe_read_json and prepare_naming_data,
    then runs one of:
      [1] Overview of naming figures
      [2] Naming profile by figure
      [3] Figure profile by lemma
    """

    # --- 0) Load naming data via central loader (JSON → Excel fallback) ---
    df_json, df_excel = load_naming_sources_with_excel_fallback(paths, data)

    # --- 1) ask for target figure (validated) ---
    target_figure = ask_valid_figure_name(paths["categorization_json"])
    if not target_figure:
        print("No figure name provided.")
        return

    # --- 2) sub-menu (validated numeric choice: 1, 2, 3) ---
    print(
        "\nWhich list should be output?\n"
        "[1] Overview of naming figures\n"
        "[2] Naming profile by figure\n"
        "[3] Figure profile by lemma"
    )
    choice = ask_user_choice("> Select option:", ["1", "2", "3"])

    # --- [1] Overview of naming figures ---
    if choice == "1":
        # ensure correct data source
        src_df, src_kind = _require_and_pick_source(
            df_json,
            df_excel,
            required_all=["Benannte Figur", "Nennende Figur"],
            required_any_prefixes=["Bezeichnung", "Epitheta"],
            book_name=book_name,
        )

        if src_df is None:
            print("No valid data source found for overview analysis.")
            return

        analyze_overview_of_naming_figures(
            book_name,
            df_json,  # give JSON as-is
            df_excel,  # give Excel as-is
            target_figure
        )
        return

    # --- [2] Naming profile by figure ---
    if choice == "2":
        try:
            source, df, cols = prepare_naming_data(book_name, df_json, df_excel)
            # explicit one-line reason if Excel was chosen
            if source == "excel":
                missing = []
                if df_json is None or df_json.empty:
                    missing.append("no JSON data")
                else:
                    if "Benannte Figur" not in df_json.columns:
                        missing.append("'Benannte Figur'")
                    if "Nennende Figur" not in df_json.columns:
                        missing.append("'Nennende Figur'")
                if missing:
                    print(f"⚠️ JSON missing {', '.join(missing)} — loading Excel file.")

            tcol = cols["target"]
            ncol = cols["namer"]
            df_sub = df.loc[df[tcol] == target_figure]
            if df_sub.empty:
                print(f"No entries found for target figure: {target_figure}")
                return

            freq = df_sub[ncol].value_counts().to_dict()
            print("\nList of naming figures (sorted by frequency):")
            for i, (name, count) in enumerate(
                sorted(freq.items(), key=lambda t: (-t[1], t[0])), start=1
            ):
                print(f"{i} {name} ({count})")
        except Exception as e:
            print(f"(Preview unavailable: {e})")

        # read raw input; allow empty -> abort
        raw = input("\n✍ Please enter the name of the figure: ").strip()
        if not raw:
            print("No naming figure provided. Aborting.")
            return
        selected_namer = raw
        analyze_naming_profile_by_figure(
            book_name, df_json, df_excel, target_figure, selected_namer
        )
        return

    # --- [3] Figure profile by lemma ---
    if choice == "3":
        # read the lemma from stdin (do NOT pass the prompt string anywhere)
        query_lemma = input("\n✍ Please enter the lemma: ").strip()
        if not query_lemma:
            print("No lemma provided. Returning to analysis menu.")
            return

        # pass BOTH sources so prepare_naming_data can fall back if needed
        analyze_figure_profile_by_lemma(
            book_name, df_json, df_excel, target_figure, query_lemma
        )
        return

    print("Invalid choice. Returning to main menu.")

# --- helper: ensure proper source selection (json vs. excel) ---
def _has_any_prefixed_columns(df, prefixes):
    """Return True if df contains any column starting with any given prefix."""
    cols = [str(c).strip().lower() for c in df.columns]
    return any(any(c.startswith(p.lower()) for c in cols) for p in prefixes)


def _require_and_pick_source(df_json, df_excel, required_all=None, required_any_prefixes=None, book_name=""):
    """
    Pick JSON if it satisfies requirements; otherwise fall back to Excel.
    Prints a clear reason when falling back.
    """
    required_all = required_all or []
    required_any_prefixes = required_any_prefixes or []

    def ok(df):
        if df is None or df.empty:
            return False
        if not all(req in df.columns for req in required_all):
            return False
        if required_any_prefixes and not _has_any_prefixed_columns(df, required_any_prefixes):
            return False
        return True

    # prefer JSON if valid
    if ok(df_json):
        return df_json, "json"

    # explain why fallback is needed
    missing = []
    if df_json is None or df_json.empty:
        missing.append("empty JSON")
    else:
        for col in required_all:
            if col not in df_json.columns:
                missing.append(f'missing "{col}"')
        if required_any_prefixes and not _has_any_prefixed_columns(df_json, required_any_prefixes):
            fam = " or ".join(required_any_prefixes)
            missing.append(f"missing family ({fam})")

    print(f'⚠️ JSON does not satisfy requirements for "{book_name}": {", ".join(missing)} – loading Excel fallback.')

    # try Excel
    if ok(df_excel):
        print("ℹ️ Excel fallback successfully loaded.")
        return df_excel, "excel"

    print(f'⚠ Neither JSON nor Excel provide required columns for "{book_name}".')
    return None, "none"

def match_name_to_lemma(target_canon, lemma, aliases=None):
    """
    Check if a lemma should count as a name-based mention of the target figure.

    Parameters:
        target_canon (str): canonical form of the target figure name
        lemma (str): lemma or designation to check
        aliases (list or None): optional list of known alias spellings

    Returns:
        bool: True if the lemma counts as a name mention of the target, else False

    Logic:
      - compares lowercase normalized forms
      - allows direct match or alias match
      - ignores empty/non-string values
    """
    if not isinstance(lemma, str) or lemma.strip() == "":
        return False

    norm_target = target_canon.lower().strip()
    norm_lemma  = lemma.lower().strip()

    # exact match
    if norm_lemma == norm_target:
        return True

    # alias match
    if aliases:
        for a in aliases:
            if isinstance(a, str) and a.lower().strip() == norm_lemma:
                return True

    return False

def analyze_overview_of_naming_figures(book_name, df_json, df_excel, target_figure):
    """
    Build the overview table:
      Nennende Figur | Gesamtnennungen | Namensnennungen | Anteil der Namensnennungen (%)

    Rules:
      - exactly 1 occurrence per row (regardless of how many designation/epithet fields are filled)
      - name-based mentions detected via match_name_to_lemma(target, lemma, aliases)
      - if no name-based matches found, ask once whether a suggested close form should be used;
        if 'y': recount using that exact form; if 'n': write reduced table
          (Nennende Figur | Gesamtnennungen)
    Output:
      data/{book_name}/analysis/{target_figure}_naming_overview.csv
    """
    # unify data & detect columns
    source, df, cols = prepare_naming_data(book_name, df_json, df_excel)
    tcol = cols["target"]
    ncol = cols["namer"]
    dcols = cols["designation_cols"]
    ecols = cols["epithet_cols"]

    # filter to the requested target figure
    dft = df.loc[df[tcol] == target_figure]

    # counts per naming figure
    counts_total = {}
    counts_name = {}

    # aliases hook (extend if you have alias retrieval attached to resolve_figure_name)
    aliases = []

    # first pass: name detection through match_name_to_lemma
    for _, row in dft.iterrows():
        namer = row.get(ncol)
        if not isinstance(namer, str) or namer.strip() == "":
            continue

        # exactly one occurrence per row
        counts_total[namer] = counts_total.get(namer, 0) + 1

        # collect lemmas for name-based check (designation + epithets)
        lemmas = []
        for c in dcols:
            val = row.get(c)
            if isinstance(val, str) and val.strip() != "":
                lemmas.append(val)
        for c in ecols:
            val = row.get(c)
            if isinstance(val, str) and val.strip() != "":
                lemmas.append(val)

        if any(match_name_to_lemma(target_figure, lm, aliases=aliases) for lm in lemmas):
            counts_name[namer] = counts_name.get(namer, 0) + 1

    # If no name-based mentions were found, offer a single confirmation for a suggested close match.
    name_hits_sum = sum(counts_name.values())
    reduced_mode = False
    suggested = None

    if name_hits_sum == 0:
        # Suggest the closest lemma form present in the data (case-insensitive)
        lemmas_all = set()
        for _, row in dft.iterrows():
            for c in dcols:
                v = row.get(c)
                if isinstance(v, str) and v.strip() != "":
                    lemmas_all.add(v)
            for c in ecols:
                v = row.get(c)
                if isinstance(v, str) and v.strip() != "":
                    lemmas_all.add(v)

        if lemmas_all:
            # difflib proposal (lowercased comparison space)
            lowered = {lm.lower(): lm for lm in lemmas_all}
            best = difflib.get_close_matches(target_figure.lower(), list(lowered.keys()), n=1, cutoff=0.6)
            if best:
                suggested = lowered[best[0]]

        if suggested:
            # UI dialog (as specified)
            print(f'{target_figure} could not be found in the "Bezeichnung" column.')
            yn = input(f'Could "{suggested}" be a variant of the name? (y/n) ').strip().lower()
            if yn == "y":
                # recount name-based using the exact suggested form
                counts_name = {}
                for _, row in dft.iterrows():
                    namer = row.get(ncol)
                    if not isinstance(namer, str) or namer.strip() == "":
                        continue
                    # one per row still applies
                    # name-based hit if any designation/epithet equals the suggested form (case-insensitive)
                    hit = False
                    for c in dcols:
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
                # produce reduced table (no name-based column)
                reduced_mode = True
        else:
            # no reasonable suggestion → reduced mode
            reduced_mode = True

    # Prepare rows
    os.makedirs(os.path.join("data", book_name, "analysis"), exist_ok=True)
    out_path = os.path.join("data", book_name, "analysis", f"naming_overview_{target_figure}.csv")

    if reduced_mode:
        # Only: Nennende Figur | Gesamtnennungen
        with open(out_path, "w", encoding="utf-8", newline="") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["Nennende Figur", "Gesamtnennungen"])
            for namer, total in sorted(counts_total.items(), key=lambda t: (-t[1], t[0])):
                w.writerow([namer, total])
        print(f"✅ Naming figure overview written to: {out_path}")
        return

    # Full table with integer percent
    rows = []
    for namer, total in counts_total.items():
        nhits = counts_name.get(namer, 0)
        pct = int(round((nhits / total) * 100)) if total > 0 else 0
        rows.append((namer, total, nhits, pct))
    rows.sort(key=lambda t: (-t[1], t[0]))

    with open(out_path, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Nennende Figur", "Gesamtnennungen", "Namensnennungen", "Anteil der Namensnennungen (%)"])
        for namer, total, nhits, pct in rows:
            w.writerow([namer, total, nhits, pct])

    print(f"✅ Naming figure overview written to: {out_path}")

def analyze_naming_profile_by_figure(book_name, df_json, df_excel, target_figure, selected_namer):
    """
    Create a list of designations used by a specific naming figure
    for the selected target figure.

    Output file:
      data/{book_name}/analysis/{target_figure}_naming_profile_by_{selected_namer}.csv

    Logic:
      - filter rows where target == target_figure and namer == selected_namer
      - if an unnumbered column 'Bezeichnung' exists → use only this one (raw, not lemmatized)
      - else use numbered Bezeichnung* + Epitheta* (already lemmatized)
      - if a verse column exists, include it
    """

    # unify data
    source, df, cols = prepare_naming_data(book_name, df_json, df_excel)
    tcol = cols["target"]
    ncol = cols["namer"]
    dcols = cols["designation_cols"]
    ecols = cols["epithet_cols"]
    vcol  = cols.get("verse_col")
    has_raw = bool(cols.get("has_unnumbered_designation"))

    # filter relevant rows
    dff = df.loc[(df[tcol] == target_figure) & (df[ncol] == selected_namer)]

    rows = []

    for _, row in dff.iterrows():
        # verse (optional)
        verse = ""
        if isinstance(vcol, str) and vcol in row:
            val = row[vcol]
            if isinstance(val, str):
                verse = val
            elif val is not None:
                verse = str(val)

        # if unnumbered 'Bezeichnung' exists → only this column
        if has_raw:
            raw_col = next((c for c in dcols if c.strip().lower() == "bezeichnung"), None)
            if raw_col:
                val = row.get(raw_col)
                if isinstance(val, str) and val.strip() != "":
                    rows.append((verse, val))
            continue

        # fallback: numbered designations and epithets
        for c in dcols:
            if c.strip().lower() == "bezeichnung":
                continue
            val = row.get(c)
            if isinstance(val, str) and val.strip() != "":
                rows.append((verse, val))
        for c in ecols:
            val = row.get(c)
            if isinstance(val, str) and val.strip() != "":
                rows.append((verse, val))

    # ensure output directory
    os.makedirs(os.path.join("data", book_name, "analysis"), exist_ok=True)
    out_path = os.path.join(
        "data", book_name, "analysis", f"naming_profile_by_{selected_namer}_{target_figure}.csv"
    )

    # write the result (header always included)
    with open(out_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["Vers", "Bezeichnung"])
        for verse, bez in rows:
            writer.writerow([serialize_verse_value(verse), bez])

    print(f"✅ Naming profile by figure written to: {out_path}")

def analyze_figure_profile_by_lemma(book_name, df_json, df_excel, target_figure, query_lemma):
    """
    Aggregate which naming figures use a given lemma for the target figure.

    Output file:
      data/{book_name}/analysis/{target_figure}_figure_profile_by_{query_lemma}.csv

    Logic:
      - restrict rows to target_figure
      - count matches of query_lemma across numbered Bezeichnung* and Epitheta* columns
        (skip unnumbered raw 'Bezeichnung' on purpose)
      - export: Nennende Figur | Count
      - if no results: print 'No results found for: <lemma>' and return without writing
    """
    # unify data
    source, df, cols = prepare_naming_data(book_name, df_json, df_excel)
    tcol = cols["target"]
    ncol = cols["namer"]
    dcols = [c for c in cols["designation_cols"] if str(c).strip().lower() != "bezeichnung"]
    ecols = cols["epithet_cols"]

    # restrict to target
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

        # numbered designations (skip unnumbered 'Bezeichnung')
        for c in dcols:
            if c.strip().lower() == "bezeichnung":
                continue
            val = row.get(c)
            if isinstance(val, str) and val.strip().lower() == q:
                hit = True
                break

        # epithets if needed
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

    # ensure output directory
    os.makedirs(os.path.join("data", book_name, "analysis"), exist_ok=True)
    out_path = os.path.join(
        "data", book_name, "analysis", f"figure_profile_by_{query_lemma}_{target_figure}.csv"
    )

    # write CSV
    with open(out_path, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Nennende Figur", "Count"])
        for namer, cnt in sorted(counts.items(), key=lambda t: (-t[1], t[0])):
            w.writerow([namer, cnt])

    print(f"✅ Figure profile by lemma written to: {out_path}")

def run_keyword_menu(config_data, paths, data, book_name):
    """
    Interactive interface for performing a keyword analysis.

    The user selects a target (whole work or figure), comparison unit (Bezeichnung, Epitheton, both),
    and reference corpus (if needed). The result is saved as a CSV file.

    Parameters:
        config_data (dict): Configuration and reference setup.
        paths (dict): File path dictionary.
        data (dict): Loaded TEI and Excel data.
        book_name (str): Short identifier for output and context.
    """
    target_json = paths["categorization_json"]
    output_dir = os.path.join("data", book_name, "analysis")
    os.makedirs(output_dir, exist_ok=True)

    print("\n📌 Do you want to analyze the whole work or a specific figure?")
    print("[1] Whole work")
    print("[2] Specific figure")

    target_choice = ask_user_choice("> ", ["1", "2"])

    reference_books = None

    if target_choice == "2":
        target = ask_valid_figure_name(target_json)
        if target is None:
            return None
        target_type = "figure"

    else:
        target = book_name
        target_type = "whole_work"
        print("📘 Please enter the names of the works to include in the reference corpus (comma-separated):")
        references = input("> ").strip()
        reference_books = [r.strip() for r in references.split(",") if r.strip()]

    print("\n🎯 What should be the unit of comparison?")
    print("[1] Naming variants (Bezeichnungen)")
    print("[2] Epithets (Epitheta)")
    print("[3] Combined")

    unit_choice = ask_user_choice("> ", ["1", "2", "3"])
    unit = {
        "1": "bezeichnung",
        "2": "epitheta",
        "3": "combined"
    }[unit_choice]

    print("\n🧪 Type in significance threshold (Log-Likelihood G²), for default = 3.84 press 'Enter':")
    threshold_input = input("> ").strip()
    try:
        threshold = float(threshold_input) if threshold_input else 3.84
    except ValueError:
        print("⚠️ Invalid input – using default threshold 3.84")
        threshold = 3.84

    # Prepare output filename
    target_label = target.replace(" ", "_")
    output_file = f"keywords_{unit}_{target_label}_{book_name}.csv"
    output_path = os.path.join(output_dir, output_file)

    # Call the actual keyword function
    if target_type == "figure":
        generate_keywords(
            target_figure=target,
            reference_books=None,
            unit=unit,
            threshold=threshold,
            target_json=target_json,
            output_path=output_path
        )
    else:
        generate_keywords(
            target_figure=None,
            reference_books=reference_books,
            unit=unit,
            threshold=threshold,
            target_json=target_json,
            output_path=output_path
        )

    print(f"✅ Keyword analysis written to: {output_path}")

    print("\n🔁 Do you want to run another keyword analysis? [y/n]")
    again = ask_user_choice("> ", ["y", "n"])
    if again == "y":
        return run_keyword_menu(config_data, paths, data, book_name)
    else:
        print("↩️ Returning to analysis menu.")
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
    Calculates keyword scores (G² Log-Likelihood) for a figure or whole work.

    Compares token frequencies against a reference corpus and filters by a significance threshold.

    Parameters:
        target_figure (str | None): Figure to analyze (None = whole work).
        reference_books (list[str] | None): List of reference corpus books (optional).
        unit (str): Token unit ("bezeichnung", "epitheta", "combined").
        threshold (float): Minimum G² value for significance.
        target_json (str): Path to JSON with categorized entries.
        output_path (str): Output path for CSV.
    """
    target_entries = safe_read_json(target_json, default=[])

    # Filter target corpus
    if target_figure:
        target_entries = [e for e in target_entries if e.get("Benannte Figur") == target_figure]

    target_tokens = extract_tokens(target_entries, unit)

    # Load reference corpus
    reference_entries = []

    if reference_books:
        for book in reference_books:
            path = os.path.join("data", f"categorization_{book}.json")
            reference_entries += safe_read_json(path, default=[])
    else:
        # fallback: all entries except target_figure
        reference_entries = [
            e for e in safe_read_json(target_json, default=[])
            if not target_figure or e.get("Benannte Figur") != target_figure
        ]

    reference_tokens = extract_tokens(reference_entries, unit)

    # Count occurrences
    target_counts = Counter(target_tokens)
    reference_counts = Counter(reference_tokens)

    results = []
    total_target = sum(target_counts.values())
    total_ref = sum(reference_counts.values())

    for token, count_t in target_counts.items():
        count_r = reference_counts.get(token, 0)

        if count_t + count_r == 0:
            continue

        # Log-likelihood Berechnung (G²)
        p = (count_t + count_r) / (total_target + total_ref)
        expected_t = p * total_target
        expected_r = p * total_ref

        log_t = count_t * math.log2(count_t / expected_t) if count_t > 0 and expected_t > 0 else 0
        log_r = count_r * math.log2(count_r / expected_r) if count_r > 0 and expected_r > 0 else 0

        keyness = 2 * (log_t + log_r)

        if keyness >= threshold:
            if count_t > count_r:
                typ = "positive"
            elif count_r > count_t:
                typ = "negative"
            else:
                typ = "neutral"

            results.append((token, count_t, count_r, round(keyness, 2), typ))

    # Sort descending
    results.sort(key=lambda x: (-x[3], x[0]))

    # Write CSV
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["Wort", "Zielanzahl", "Referenzanzahl", "Keyness", "Typ"])
        for row in results:
            writer.writerow(row)

def extract_tokens(entries: list[dict], unit: str) -> list[str]:
    """
    Extracts all naming variant and/or epithet tokens from categorized entries.

    Parameters:
        entries (list[dict]): Categorization data entries.
        unit (str): Token type: "bezeichnung", "epitheta", or "combined".

    Returns:
        list[str]: Flat list of normalized tokens.
    """
    tokens = []

    for entry in entries:
        if unit in ("bezeichnung", "combined"):
            for i in range(1, 5):
                val = entry.get(f"Bezeichnung {i}")
                if isinstance(val, str) and val.strip():
                    tokens.append(val.strip())

        if unit in ("epitheta", "combined"):
            for i in range(1, 6):
                val = entry.get(f"Epitheta {i}")
                if isinstance(val, str) and val.strip():
                    tokens.append(val.strip())

    return tokens

def run_collocation_menu(config_data, paths, data, book_name):
    """
    Interactive CLI interface for collocation analysis.

    The user selects a target figure and search term. Results are shown in console
    or saved as CSV in KWIC format (context display).

    Parameters:
        config_data (dict): Global config.
        paths (dict): File path dictionary.
        data (dict): TEI and Excel data.
        book_name (str): Book identifier for file naming.
    """
    _ = config_data
    categorization_path = paths["categorization_json"]

    print("\n📌 Do you want to analyze the whole work or only a specific figure?")
    print("[1] Whole work")
    print("[2] Specific figure")

    target_mode = ask_user_choice("> ", ["1", "2"])
    only_figure = None

    if target_mode == "2":
        only_figure = ask_valid_figure_name(categorization_path)
        if only_figure is None:
            return

    type_value = input("🔍 Please enter the type to search for (e.g., \"küene\"):\n> ").strip()

    while True:
        print("\n📤 Where should the results be displayed?")
        print("[1] Console")
        print("[2] Save as CSV file")

        output_choice = ask_user_choice("> ", ["1", "2"])
        output_target = "console" if output_choice == "1" else "csv"

        if output_target == "csv":
            type_label = type_value.replace(" ", "_")
            fig_label = only_figure.replace(" ", "_") if only_figure else "whole_work"
            output_dir = os.path.join("data", book_name, "analysis")
            os.makedirs(output_dir, exist_ok=True)
            filename = f"collocations_{fig_label}_{type_label}_{book_name}.csv"
            output_path = os.path.join(output_dir, filename)
        else:
            output_path = None

        try:
            generate_collocations(
                data=data,
                type_value=type_value,
                book_name=book_name,
                config_data=config_data,
                only_figure=only_figure,
                output_target=output_target,
                output_path=output_path
            )
            break  # ✅ innerhalb von while

        except PermissionError:
            print("\n⚠️ The Excel file appears to be open.")
            print("❗ Please close it and try again.")
            print("↩️ Returning to output choice...\n")

def generate_collocations(
    data: dict,
    type_value: str,
    book_name: str,
    config_data: dict,
    only_figure: str | None,
    output_target: str,
    output_path: str | None
):
    """
    Finds collocation contexts (KWIC) for a given type and figure.

    Searches categorized entries for naming variants or epithets matching the search term,
    and extracts their collocations from Excel or TEI fallback.

    Parameters:
        data (dict): Full data set (Excel, TEI).
        type_value (str): Search string (type).
        book_name (str): Name of the current book.
        config_data (dict): Config settings.
        only_figure (str | None): Figure filter (optional).
        output_target (str): "console" or "csv".
        output_path (str | None): File path if saving is enabled.
    """
    json_path = os.path.join("data", book_name, f"categorization_{book_name}.json")
    entries = safe_read_json(json_path, default=[])
    lemma_map = safe_read_json("data/lemma_normalization.json", default={})

    # Filter entries by figure if given
    if only_figure:
        entries = [e for e in entries if e.get("Benannte Figur") == only_figure]

    # Load Excel with fallback and sheet/column check
    df = load_collocation_sheet(config_data, book_name)
    if df is None:
        print("⚠️ Could not load the Excel sheet with 'Kollokationen'.")
        print("🔄 Falling back to TEI to reconstruct collocations.")
        df = build_fallback_collocation_df_from_tei(data["xml"])

    results = []

    for entry in entries:
        all_type_fields = [
            entry.get(f"Bezeichnung {i}") for i in range(1, 5)
        ] + [
            entry.get(f"Epitheta {i}") for i in range(1, 6)
        ]

        if not any(t == type_value for t in all_type_fields if isinstance(t, str)):
            continue

        vers = entry.get("Vers")
        figur = entry.get("Benannte Figur")
        original_text = get_first_valid_text(
            entry.get("Erzähler"),
            entry.get("Bezeichnung"),
            entry.get("Eigennennung")
        )

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

        kollokation = match.iloc[0].get("Kollokationen")
        if not isinstance(kollokation, str) or not kollokation.strip():
            continue

                # Hole alle zugehörigen Varianten aus dem Lemma-Mapping
        raw_variants = lemma_map.get(type_value, [])

        # Sicherheit: Nur gültige Strings verwenden
        variants: List[str] = [v.strip() for v in raw_variants if isinstance(v, str) and v.strip()]

        # Fallback: Original-Keyword selbst auch aufnehmen (kleingeschrieben)
        if isinstance(type_value, str) and type_value.strip():
            variants.append(type_value.strip().lower())

        left, hit, right = format_kwic(kollokation, variants)
        results.append((vers, figur, left, hit, right))

    # Output formatting
    if output_target == "console":
        for _, _, left, hit, right in results:
            print(f"{left.strip():>40}  \033[1m\033[93m{hit}\033[0m  {right.strip():<40}")
    elif output_target == "csv" and output_path:
        with open(output_path, "w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["Vers", "Benannte Figur", "Left", "Hit", "Right"])
            for row in results:
                writer.writerow(row)
        print(f"✅ Collocation results saved to: {output_path}")

def format_kwic(context: str, variants: list[str]) -> tuple[str, str, str]:
    """
    Splits a collocation string into KWIC format: left, hit, and right part.

    Finds the first match of any variant and extracts the surrounding context.

    Parameters:
        context (str): The full collocation string.
        variants (list[str]): Lowercase variant tokens to match.

    Returns:
        tuple[str, str, str]: Left, hit, and right segments of the string.
    """
    context_lower = context.lower()

    for variant in variants:
        index = context_lower.find(variant.lower())
        if index != -1:
            left = context[:index].strip()
            hit = context[index:index + len(variant)]
            right = context[index + len(variant):].strip()
            return left, hit, right

    # No variant matched
    return context.strip(), "", ""

def run_visualization_menu(paths, book_name, data):
    while True:
        print("\n📈 Which visualization do you want to run?")
        print("[1] Verse-based naming distribution")
        print("[2] Intra-Figure Co-Occurrence Heatmap")
        print("[3] Sunburst Visualization")
        print("[4] Back to analysis menu")

        choice = ask_user_choice("> ", ["1", "2", "3", "4"])

        if choice == "1":
            visualize_verse_naming_distribution(paths, book_name)
        elif choice == "2":
            visualize_intra_figure_cooccurrence_heatmap(paths, book_name)
        elif choice == "3":
            run_sunburst_visualization(paths, book_name, data)

        elif choice == "4":
            print("↩️ Returning to analysis menu.")
            break

# ============================================================
# GLOBAL VISUAL STYLE CONFIGURATION (FAIR-compliant)
# Applies to all visualization functions defined below
# ============================================================

GLOBAL_VISUAL_STYLE: dict[str, Any] = {
    "typography": {
        # "font_family": "Times New Roman",
        "font_family": "Cormorant Garamond Medium",
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
            "Eigenname": "#EFE4D4",
            "Antonomasie": "#F9C691",

            # Counterpole
            "Epitheta": "#2F4A6D",

            # Epitheta (leaf/lexeme level) — optional but useful for consistent leaf styling
            "Epitheta_lexeme": "#6A97B8",

            # --- aliases (UI terminology, English) ---
            # Keep UI labels compatible with the same semantic categories
            "Name": "#EFE4D4",
            "Antonomasia": "#F9C691",
            "Epithet": "#2F4A6D",
            "Epithets": "#2F4A6D",
        },

        "levels": {
            "CORE": "#0D1E26",
            "STRUCTURE": "#A6B4A0",
            "NEUTRAL_TEXT": "#2D2926",
            "AUXILIARY": "#A3A39A",
        },
    },
}

def apply_global_visual_style(fig, *, tick_font_size=None, show_grid=None, has_axes: bool = True):
    """
    Apply global visual defaults to a Plotly figure.

    This function sets only presentation defaults (typography, background, margins, legend styling).
    It must not encode plot-specific semantics or data logic.
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
            title_font={"family": typography["font_family"], "size": typography["base_size"], "color": levels["NEUTRAL_TEXT"]},
            showgrid=show_grid,
            gridcolor=levels["AUXILIARY"],
            zeroline=False,
        )
        fig.update_yaxes(
            tickfont={"family": typography["font_family"], "size": tick_size, "color": levels["NEUTRAL_TEXT"]},
            title_font={"family": typography["font_family"], "size": typography["base_size"], "color": levels["NEUTRAL_TEXT"]},
            showgrid=show_grid,
            gridcolor=levels["AUXILIARY"],
            zeroline=False,
        )

    return fig

def apply_global_visual_visibility(fig, *, show_title=None, show_legend=None, show_axis_labels=None):
    """
    Toggle visibility of title, legend, and axis labels in a consistent way.

    If an argument is None, the global default from GLOBAL_VISUAL_STYLE["layout"] is used.
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
    Default export filename stub for Modebar exports (no extension).
    Format: viz_YYYY_MM_DD_HHMM
    """
    return f"{prefix}_{datetime.now():%Y_%m_%d_%H%M}"


def apply_global_visual_modebar_export(
    fig,
    output_path: str | Path,
    *,
    filename_stub: str | None = None,
) -> None:
    """
    Write an interactive Plotly HTML file with a customized export overlay.

    - Keeps Plotly interactivity + standard Modebar.
    - Adds two export buttons (SVG + PNG A4@300dpi).
    - Ensures updatemenus (dropdowns etc.) are NOT included in exported images by
      temporarily hiding them during export.
    - Positions the overlay buttons dynamically so they won't cover the Modebar
      (independent of desktop size / zoom). Repositions on resize + redraw.

    Notes:
    - This affects only exports inside the HTML.
    - The HTML remains interactive (updatemenus restored after export).
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

    var btnPng = makeBtn("Download plot as PNG (A4, 300 dpi)");
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
    
      hideMenusThenDownload(gd, {{
        format: "png",
        filename: "{filename_stub}",
        scale: 3
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

def visualize_verse_naming_distribution(paths, book_name):
    """
    Interactive CLI interface for visualizing naming variants and epithets using Plotly.

    The user is prompted to:
    - Select a figure to visualize,
    - Choose a token type (naming variants, epithets, or both),
    - Select specific tokens to include,
    - Define output mode (save as HTML or open in browser).

    The visualization is a scatter plot of tokens by verse, with optional category coloring.

    Parameters:
        paths (dict): Dictionary of file paths including 'categorization_json'.
        book_name (str): Name of the current book for labeling and output folder generation.
    """
    entries = safe_read_json(paths["categorization_json"], default=[])
    if not entries:
        print("❌ No categorization data available.")
        return

    df = pd.DataFrame(entries)

    # Step 1 – Ask for figure name
    figure_name = ask_valid_figure_name(paths["categorization_json"])
    if figure_name is None:
        return

    # Step 2 – Choose the visualization type
    print("\n📌 What should be visualized?")
    print("[1] Naming variants")
    print("[2] Epithets")
    print("[3] Combined")
    variant_type = ask_user_choice("> ", ["1", "2", "3"])

    variant_label = {
        "1": "Naming variants",
        "2": "Epithets",
        "3": "Naming variants & epithets"
    }[variant_type]

    # Step 3 – Prepare long-format DataFrame
    df_figure = df[df["Benannte Figur"] == figure_name].copy()
    naming_cols = [f"Bezeichnung {i}" for i in range(1, 5)]
    epithet_cols = [f"Epitheta {i}" for i in range(1, 6)]

    if variant_type == "1":
        selected_cols = naming_cols
    elif variant_type == "2":
        selected_cols = epithet_cols
    else:
        selected_cols = naming_cols + epithet_cols

    meta_cols = ["Nennende Figur", "Erzähler", "Eigennennung"]

    all_entries = []
    for col in selected_cols:
        keep_cols = ["Vers", col] + [c for c in meta_cols if c in df_figure.columns]
        temp = df_figure[keep_cols].dropna(subset=["Vers", col]).rename(columns={col: "Token"})
        all_entries.append(temp)

    df_combined = pd.concat(all_entries, ignore_index=True)

    df_combined["Token"] = df_combined["Token"].astype(str).str.strip()
    df_combined["Vers"] = pd.to_numeric(df_combined["Vers"], errors="coerce")

    # normalize meta cols (safe even if some cols are missing)
    for c in meta_cols:
        if c in df_combined.columns:
            df_combined[c] = df_combined[c].astype(str).str.strip()

    # Step 4 – Count frequencies
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

    # Step 5 – Token selection (user input with validation)
    if variant_type in ("1", "3"):
        print(f"\n📁 Available naming variants for {figure_name}:")
        for i, (token, freq) in enumerate(naming_list, 1):
            print(f"{i}. {token} – {freq}")
        while True:
            input_str = input(
                "\n✍ Which naming variants should be included? (e.g., 1–3, 5)\n"
                "Note: Selecting more than 14 entries in total may reduce visual clarity.\n> "
            ).strip()
            indices = parse_token_selection(input_str, len(naming_list))
            if indices:
                selected_naming = [naming_list[i - 1][0] for i in indices]
                break
            print("⚠️ Invalid input – please try again.")

    if variant_type in ("2", "3"):
        print(f"\n📁 Available epithets for {figure_name}:")
        for i, (token, freq) in enumerate(epithet_list, 1):
            print(f"{i}. {token} – {freq}")
        while True:
            input_str = input(
                "\n✍ Which epithets should be included? (e.g., 1–3, 5)\n"
                "Note: Selecting more than 14 entries in total may reduce visual clarity.\n> "
            ).strip()
            indices = parse_token_selection(input_str, len(epithet_list))
            if indices:
                selected_epithets = [epithet_list[i - 1][0] for i in indices]
                break
            print("⚠️ Invalid input – please try again.")

    # Combine selected tokens
    tokens_to_plot = selected_naming + selected_epithets
    if not tokens_to_plot:
        print("⚠️ No tokens selected – aborting.")
        return

    # Step 6 – Filter for plot and prepare HTML display labels
    df_plot = df_combined[df_combined["Token"].isin(tokens_to_plot)].copy()

    # Ensure renderable points: drop rows without a valid verse number
    df_plot = df_plot.dropna(subset=["Vers"])

    plot_token_counts = Counter(df_plot["Token"])
    sorted_tokens = [token for token, _ in plot_token_counts.most_common()]

    df_plot["Token_html"] = df_plot["Token"].apply(lambda x: f"<i>{x}</i>")
    df_plot["Token_html"] = pd.Categorical(
        df_plot["Token_html"],
        categories=[f"<i>{t}</i>" for t in sorted_tokens],
        ordered=True
    )

    if variant_type == "3":
        df_plot["Category"] = df_plot["Token"].apply(
            lambda x: "Naming variant" if x in selected_naming else "Epithet"
        )

    # Step 7 – Create interactive plot
    trace_tokens: list[str | None] = []

    if variant_type == "3":
        fig = go.Figure()

        color_map = {
            "Naming variant": GLOBAL_VISUAL_STYLE["colors"]["categories"]["Naming variants"],
            "Epithet": GLOBAL_VISUAL_STYLE["colors"]["categories"]["Epitheta"],
        }

        # Ensure stable ordering by frequency
        token_order = sorted_tokens

        # Legend dummies (keep legend compact: 2 entries only)
        fig.add_trace(
            go.Scatter(
                x=[None],
                y=[None],
                mode="markers",
                marker={"size": 18, "opacity": 0.7, "color": color_map["Naming variant"]},
                name="Naming variant",
                legendgroup="Naming variant",
                showlegend=True,
                hoverinfo="skip",
            )
        )
        trace_tokens.append(None)

        fig.add_trace(
            go.Scatter(
                x=[None],
                y=[None],
                mode="markers",
                marker={"size": 18, "opacity": 0.7, "color": color_map["Epitheton"]},
                name="Epitheton",
                legendgroup="Epitheton",
                showlegend=True,
                hoverinfo="skip",
            )
        )
        trace_tokens.append(None)

        # One trace per token (enables correct Show-N token visibility)
        for token in token_order:
            df_token = df_plot[df_plot["Token"] == token].copy()
            if df_token.empty:
                continue

            token_html = f"<i>{token}</i>"
            category = df_token["Category"].iloc[0]
            marker_color = color_map.get(category, GLOBAL_VISUAL_STYLE["colors"]["levels"]["AUXILIARY"])

            fig.add_trace(
                go.Scatter(
                    x=df_token["Vers"],
                    y=[token_html] * len(df_token),
                    mode="markers",
                    marker={"opacity": 0.7, "color": marker_color},
                    name=token_html,  # token identifier for Show-N toggles
                    legendgroup=category,
                    showlegend=False,  # legend handled by dummy traces above
                    meta=token_html,
                    hovertemplate="Vers: %{x}<br>Token: %{meta}<extra></extra>",
                )
            )
            trace_tokens.append(token)  # raw token id for visibility control

        fig.update_layout(title=f"{variant_label} for '{figure_name}'")
    else:
        fig = go.Figure()
        trace_tokens = []  # keep consistent type; not used in this branch

        base_color = (
            GLOBAL_VISUAL_STYLE["colors"]["categories"]["Naming variants"]
            if variant_type == "1"
            else GLOBAL_VISUAL_STYLE["colors"]["categories"]["Epitheta"]
        )

        token_order = sorted_tokens
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
                    name=token_html,  # keep for Show-N visibility logic if needed later
                    showlegend=False,  # IMPORTANT: no token legend
                    meta=token_html,
                    hovertemplate="Vers: %{x}<br>Token: %{meta}<extra></extra>",
                )
            )

    # Plot-specific axis/category ordering and labels (A4-readable default: up to 14)
    max_n = len(sorted_tokens)
    default_show_n = min(14, max_n)

    top_tokens = sorted_tokens[:default_show_n]
    top_categories = [f"<i>{t}</i>" for t in top_tokens]

    fig.update_yaxes(
        type="category",
        categoryorder="array",
        categoryarray=top_categories,
        tickmode="array",
        tickvals=top_categories,
        ticktext=top_categories,
        title_text=variant_label,
    )

    fig.update_xaxes(title_text="Verses")

    # Apply global visual defaults and visibility rules
    apply_global_visual_style(fig)
    apply_global_visual_visibility(fig, show_legend=(variant_type == "3"))

    # Keep plot-specific height (optional)
    fig.update_layout(height=800)

    # --- Interactive "Show N" toggles (14, 28, 42, ... up to max selected) ---
    def compute_tick_size(token_count: int) -> int:
        # 14 tokens are calibrated for tick size 36; reduce proportionally beyond that
        if token_count <= 14:
            return 36
        size = int(round(36 * 14 / token_count))
        return max(12, size)

    def compute_marker_size(token_count: int) -> int:
        """
        Scale marker size proportionally to tick font size.
        Calibrated for A4 readability at 14 tokens.
        """
        if token_count <= 14:
            return 18
        size = int(round(18 * 14 / token_count))
        return max(6, size)

    show_steps = list(range(14, max_n + 1, 14))

    if not show_steps:
        show_steps = [max_n]
    elif show_steps[-1] != max_n:
        show_steps.append(max_n)

    # Hide traces beyond the default Top-N view (robust: use trace_tokens if available)
    if variant_type == "3":
        # use trace_tokens
        initial_visible = []
        top_set = set(top_tokens)

        for tok in trace_tokens:
            if tok is None:
                initial_visible.append(True)
            else:
                initial_visible.append(tok in top_set)

        for i, tr in enumerate(fig.data):
            tr.visible = initial_visible[i]
    else:
        # px.scatter branch → name-based logic
        visible_by_name = {f"<i>{t}</i>": (t in top_tokens) for t in sorted_tokens}

        initial_visible = []
        for tr in fig.data:
            trace_name = str(getattr(tr, "name", None))
            initial_visible.append(visible_by_name.get(trace_name, True))

        for i, tr in enumerate(fig.data):
            tr.visible = initial_visible[i]

    buttons = []
    for n in show_steps:
        current_tokens = sorted_tokens[:n]
        current_categories = [f"<i>{t}</i>" for t in current_tokens]
        current_tick_size = compute_tick_size(n)
        current_marker_size = compute_marker_size(n)

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

    # Ensure the initial tick size matches the default Top-14 view
    fig.update_yaxes(tickfont={"size": compute_tick_size(default_show_n)})

    # Ensure marker size matches the default Top-N view
    default_marker_size = compute_marker_size(default_show_n)

    fig.update_traces(
        marker={"size": default_marker_size},
        selector=dict(mode="markers")
    )

    # Step 8 – Ask for output mode
    print("\n📅 How should the output be handled?")
    print("[1] Save as HTML file")
    print("[2] Show plot in browser")
    print("[3] Both")
    output_mode = ask_user_choice("> ", ["1", "2", "3"])

    output_dir = os.path.join("data", book_name, "visualization")
    os.makedirs(output_dir, exist_ok=True)
    variant_label = "combined" if variant_type == "3" else "epithets" if variant_type == "2" else "naming"
    filename = f"viz_{variant_label}_{figure_name}.html"
    output_path = os.path.join(output_dir, filename)

    # Step 9 – Output
    if output_mode == "1":
        # Save only
        apply_global_visual_modebar_export(fig, output_path)
        print(f"\n✅ Visualization completed.")
        print(f"📂 File saved at:\n{output_path}")

    elif output_mode == "2":
        # Display only → use the temporary file
        tmp_filename = f"viz_{uuid.uuid4().hex[:8]}.html"
        tmp_path = os.path.join(paths["tmp_dir"], tmp_filename)
        apply_global_visual_modebar_export(fig, tmp_path)
        webbrowser.open_new_tab(f"file://{os.path.abspath(tmp_path)}")
        print(f"🌐 The plot has been opened in your browser.")
        print(f"🧾 Temporary file created at: {tmp_path}")

    elif output_mode == "3":
        # Save and display → use the saved file
        apply_global_visual_modebar_export(fig, output_path)
        print(f"\n✅ Visualization completed.")
        print(f"📂 File saved at:\n{output_path}")
        webbrowser.open_new_tab(f"file://{os.path.abspath(output_path)}")
        print(f"🌐 The plot has been opened in your browser.")

def parse_token_selection(input_str: str, max_value: int) -> list[int] | None:
    """
    Parses a user input string of selected token indices into a list of valid integers.

    Accepted formats:
        - Single numbers: "3"
        - Ranges: "1-3"
        - Mixed: "1-3,5,7"
        - Unicode dashes (–) are normalized to hyphens (-)

    Returns None if input is invalid or out of range.

    Parameters:
        input_str (str): Raw input string provided by the user.
        max_value (int): Maximum allowed index (e.g., length of the list of options).

    Returns:
        list[int] | None: Sorted list of valid indices, or None if validation fails.
    """
    if not input_str.strip():
        return None

    input_str = input_str.replace("–", "-").replace(" ", "")
    parts = input_str.split(",")

    result = set()
    for part in parts:
        if "-" in part:
            try:
                start_str, end_str = part.split("-", 1)
                start = int(start_str)
                end = int(end_str)
                if start > end or start < 1 or end > max_value:
                    return None
                result.update(range(start, end + 1))
            except ValueError:
                return None
        else:
            try:
                value = int(part)
                if 1 <= value <= max_value:
                    result.add(value)
                else:
                    return None
            except ValueError:
                return None

    return sorted(result)

def collect_tokens_for_cooccurrence(row: dict, include_naming_variants: bool, include_epithets: bool) -> list[str]:
    """
    Collects all relevant tokens (naming variants and/or epithets) from a single entry.

    Parameters:
        row (dict): One entry from the categorization JSON.
        include_naming_variants (bool): Whether to include 'Bezeichnung 1–4'.
        include_epithets (bool): Whether to include 'Epitheta 1–5'.

    Returns:
        list[str]: A sorted, de-duplicated list of tokens extracted from the row.
    """
    tokens: list[str] = []

    if include_naming_variants:
        for i in range(1, 5):
            v = row.get(f"Bezeichnung {i}", "")
            if isinstance(v, str) and v.strip():
                tokens.append(v.strip())

    if include_epithets:
        for i in range(1, 6):
            v = row.get(f"Epitheta {i}", "")
            if isinstance(v, str) and v.strip():
                tokens.append(v.strip())

    # remove duplicates within one entry and sort alphabetically for stability
    return sorted(set(tokens))


def visualize_intra_figure_cooccurrence_heatmap(paths: dict, book_name: str) -> None:
    """
    Intra-Figure Co-Occurrence Heatmap (CLI + rendering).
    Scope: within-entry co-occurrence (row-based), no verse window.
    Source: categorization JSON (Bezeichnung 1–4, Epitheta 1–5).
    Defaults (silent): min_pair_count=2, top_n=30.
    """
    # Step 1 – Figure selection
    figure_name = ask_valid_figure_name(paths["categorization_json"])

    # Step 2 – Labeling types
    print("Which labeling types should be included?")
    print("[1] Both")
    print("[2] Only naming variants")
    print("[3] Only epithets")
    variant_type = ask_user_choice("> ", ["1", "2", "3"])
    include_naming_variants = (variant_type in ("1", "2"))
    include_epithets = (variant_type in ("1", "3"))

    # Fixed defaults (silent)
    min_pair_count = 2
    top_n = 30

    # Load data
    entries = safe_read_json(paths["categorization_json"], default=[])
    rows = [e for e in entries if isinstance(e, dict) and e.get("Benannte Figur") == figure_name]

    # Collect tokens per row
    token_rows = [
        collect_tokens_for_cooccurrence(r, include_naming_variants, include_epithets)
        for r in rows
    ]
    token_rows = [t for t in token_rows if len(t) >= 2]

    # Count unordered pairs per row
    pair_counter: Counter = Counter()
    for toks in token_rows:
        for a, b in combinations(toks, 2):
            pair = tuple(sorted((a, b)))
            pair_counter[pair] += 1

    # Apply threshold
    pair_counter = Counter({p: c for p, c in pair_counter.items() if c >= min_pair_count})
    if not pair_counter:
        print("\nℹ️ No co-occurring pairs met the minimum threshold.")
        return

    # Top-N selection
    top_pairs = pair_counter.most_common(top_n)
    tokens = sorted(set([t for p, _ in top_pairs for t in p]))
    index = {t: i for i, t in enumerate(tokens)}

    # Build symmetric full matrix (both halves)
    size = len(tokens)
    matrix = np.zeros((size, size), dtype=float)

    for (a, b), c in top_pairs:
        i, j = index[a], index[b]
        if i == j:
            continue  # diagonal stays 0
        matrix[i, j] = c
        matrix[j, i] = c

    # convert absolute counts to percent shares
    # NOTE: With full symmetry, each undirected pair is represented twice in the matrix.
    # Using sum(matrix) keeps the heatmap normalized to 100% over all displayed cells.
    total = float(np.sum(matrix))
    if total > 0:
        matrix_pct = (matrix / total) * 100.0
    else:
        matrix_pct = matrix.copy()

    levels = GLOBAL_VISUAL_STYLE["colors"]["levels"]

    # Cool, fine-grained sequential scale derived purely from STRUCTURE → AUXILIARY → CORE
    # No categorical semantics attached (tonal interpolation only)

    c_structure = hex_color_to_rgb_tuple(levels["STRUCTURE"])
    c_auxiliary = hex_color_to_rgb_tuple(levels["AUXILIARY"])
    c_core = hex_color_to_rgb_tuple(levels["CORE"])

    heatmap_colorscale = (
            [rgb_tuple_to_plotly_color(c) for c in n_colors(c_structure, c_auxiliary, 24, colortype="tuple")]
            + [rgb_tuple_to_plotly_color(c) for c in n_colors(c_auxiliary, c_core, 24, colortype="tuple")][1:]
    )
    # Robust upper cap to preserve contrast under skewed distributions
    # Same rule applied across all heatmaps (comparability preserved)
    z_cap = float(np.nanpercentile(matrix_pct, 99))
    z_cap = min(100.0, max(10.0, z_cap))  # safety bounds

    fig = px.imshow(
        matrix_pct,
        x=tokens,
        y=tokens,
        labels=dict(x="Label", y="Label", color="Co-occurrence share (%)"),
        aspect="auto",
        range_color=(0, z_cap),
        color_continuous_scale=heatmap_colorscale,
    )

    # hover: percent + absolute value
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

    apply_global_visual_style(fig, tick_font_size=22, show_grid=False)
    apply_global_visual_visibility(fig, show_legend=False)

    # Heatmap-specific: ensure enough top margin for title after global style is applied
    fig.update_layout(margin={**GLOBAL_VISUAL_STYLE["layout"]["margins"], "t": 100})

    # Step 3 – Output handling (identical to existing viz)
    print("\n📅 How should the output be handled?")
    print("[1] Save as HTML file")
    print("[2] Show plot in browser")
    print("[3] Both")
    output_mode = ask_user_choice("> ", ["1", "2", "3"])

    output_dir = os.path.join("data", book_name, "visualization")
    os.makedirs(output_dir, exist_ok=True)
    variant_label = "cooccurrence"
    filename = f"viz_{variant_label}_{figure_name}.html"
    output_path = os.path.join(output_dir, filename)

    filename_stub = os.path.splitext(filename)[0]

    if output_mode == "1":
        apply_global_visual_modebar_export(fig, output_path, filename_stub=filename_stub)
        print("\n✅ Visualization completed.")
        print(f"📂 File saved at:\n{output_path}")

    elif output_mode == "2":
        tmp_dir = paths.get("tmp_dir", os.path.join("data", book_name, "tmp"))
        os.makedirs(tmp_dir, exist_ok=True)
        tmp_filename = f"viz_{uuid.uuid4().hex[:8]}.html"
        tmp_path = os.path.join(tmp_dir, tmp_filename)
        apply_global_visual_modebar_export(fig, tmp_path, filename_stub=filename_stub)
        webbrowser.open_new_tab(f"file://{os.path.abspath(tmp_path)}")
        print("🌐 The plot has been opened in your browser.")
        print(f"🧾 Temporary file created at: {tmp_path}")

    elif output_mode == "3":
        apply_global_visual_modebar_export(fig, output_path, filename_stub=filename_stub)
        print("\n✅ Visualization completed.")
        print(f"📂 File saved at:\n{output_path}")
        webbrowser.open_new_tab(f"file://{os.path.abspath(output_path)}")
        print("🌐 The plot has been opened in your browser.")

def run_sunburst_visualization(paths, book_name, data):
    """
    Entry point for all Sunburst visualizations.
    Step 1: Ask user for Sunburst type
    Step 2: Delegate to figure-centered or work-centered view
    """

    print("\n🌞 Sunburst Visualization")
    print("[1] Figure-centered view")
    print("[2] Work-centered overview")
    print("[3] Back to visualization menu")

    choice = ask_user_choice("> ", ["1", "2", "3"])

    if choice == "1":
        visualize_sunburst_figure_view(paths, book_name, data)

    elif choice == "2":
        visualize_sunburst_work_overview(paths, book_name, data)

    elif choice == "3":
        print("↩️ Returning.")
        return

def visualize_sunburst_figure_view(paths, book_name, data):
    """
    Figure-centered sunburst visualization.

    Modes:
    [1] Figure centred mode: Types → Lemma
        path = center_figure → type → lemma
    [2] Namer centered mode: Naming figures and/or narrator → Lemma
        path = center_figure → namer → lemma
    """

    # --- 1) Load naming data (JSON + Excel fallback) ---
    df_json, df_excel = load_naming_sources_with_excel_fallback(paths, data)
    source, df, cols = prepare_naming_data(book_name, df_json, df_excel)

    if df is None or df.empty:
        print("⚠️ No naming data available after prepare_naming_data.")
        return

    # fig will be created conditionally below

    # --- 2) Ask for figure name ---
    figure_name = ask_valid_figure_name(paths["categorization_json"])

    # --- 3) Ask for mode ---
    print()
    print("[1] Figure centred mode: Types → Lemma")
    print("[2] Namer centered mode: Naming figures and/or narrator → Lemma")
    mode_choice = ask_user_choice("> ", ["1", "2"])

    # --- 4) Build aggregated data ---
    if mode_choice == "1":
        sunburst_df = build_sunburst_data_types_lemma(df, cols, figure_name)
    else:
        sunburst_df = build_sunburst_data_namer_lemma(df, cols, figure_name)

    if sunburst_df is None or sunburst_df.empty:
        print("⚠️ No data available for figure-centred sunburst.")
        return

    # --- 5) Compute percentages relative to the center figure ---
    total_freq = sunburst_df["frequency"].sum()
    if total_freq > 0:
        sunburst_df["pct_of_center"] = sunburst_df["frequency"] / total_freq
    else:
        sunburst_df["pct_of_center"] = 0.0

    # --- 6) Basic sorting for our own reference (Plotly behält seine Struktur) ---
    sunburst_df["type_group"] = pd.Categorical(
        sunburst_df["type_group"],
        categories=["Naming variants", "Epithets"],
        ordered=True,
    )

    sunburst_df = sunburst_df.sort_values(
        ["type_group", "lemma"],
        ascending=[True, True],
    ).reset_index(drop=True)

    if mode_choice == "1":
        # center_figure → type_group → lemma (types-centered view)

        categories = GLOBAL_VISUAL_STYLE["colors"]["categories"]
        levels = GLOBAL_VISUAL_STYLE["colors"]["levels"]

        # Defensive category lookups (keeps refactors stable)
        name_color = categories.get("Name", "#EFE4D4")
        naming_variants_ring = categories.get("Naming variants", "#8C6A4A")
        epithets_ring = categories.get("Epithets", "#2F4A6D")
        epithets_lexeme = categories.get("Epitheta_lexeme", epithets_ring)

        fig = px.sunburst(
            sunburst_df,
            path=["center_figure", "type_group", "lemma"],
            values="frequency",
            color="color_group",
            color_discrete_map={
                "Name": categories.get("Name", "#EFE4D4"),
                "Antonomasia": categories.get("Antonomasia", "#F9C691"),
                "Epithet": categories.get("Epithet", "#2F4A6D"),
            },
        )

        if fig.data:
            trace = fig.data[0]

            # Segment borders (lines)
            trace.update(
                marker=dict(
                    line=dict(
                        color=levels.get("AUXILIARY", "#A3A39A"),
                        width=1.0,
                    )
                )
            )

            labels = list(trace["labels"])
            parents = list(trace["parents"])
            colors = list(trace["marker"]["colors"])

            # --- Color patching: root + group rings + epithet leaves ---
            for i, (lab, par) in enumerate(zip(labels, parents)):
                # Root node (center figure)
                if lab == figure_name and (par is None or par == ""):
                    colors[i] = name_color
                    continue

                # Middle ring groups
                if lab == "Naming variants":
                    colors[i] = naming_variants_ring
                    continue
                if lab == "Epithets":
                    colors[i] = epithets_ring
                    continue

                # Leaves under Epithets
                if par == "Epithets":
                    colors[i] = epithets_lexeme

            trace["marker"]["colors"] = colors

            # --- WCAG-ish text color per segment (expects shared.py to handle hex + rgba) ---
            text_colors = [
                pick_accessible_text_color(
                    bg,
                    dark_text_hex=levels.get("NEUTRAL_TEXT", "#2D2926"),
                    light_text_hex="#FFFFFF",
                )
                for bg in colors
            ]
            trace["textfont"] = {"color": text_colors}

            # --- Hover data ---
            group_totals = (
                sunburst_df.groupby("type_group", observed=False)["frequency"]
                .sum()
                .to_dict()
            )

            lemma_map: dict[str, list[object]] = {}
            for _, row in sunburst_df.iterrows():
                lemma_map[str(row["lemma"])] = [
                    str(row["type_group"]),  # Group
                    str(row["color_group"]),  # Category
                    str(row["lemma"]),  # Lemma
                    float(row["frequency"]),
                    float(row["pct_of_center"]),
                ]

            customdata: list[list[object]] = []
            for lab in labels:
                if lab == figure_name:
                    customdata.append(["", "", str(lab), float(total_freq), 1.0])
                elif lab in group_totals:
                    cnt = float(group_totals.get(lab, 0.0))
                    share = (cnt / total_freq) if total_freq > 0 else 0.0
                    customdata.append([str(lab), "", str(lab), cnt, share])
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

    else:
        # center_figure → namer_display → lemma (namer-centered view)

        sunburst_df["namer_display"] = sunburst_df["namer"].apply(
            lambda v: (
                v[max(v.rfind("/"), v.rfind("#")) + 1:].strip()
                if isinstance(v, str) and max(v.rfind("/"), v.rfind("#")) != -1
                else v
            )
        )

        per_namer_raw = (
            sunburst_df.groupby("namer_display", dropna=False)["frequency"]
            .sum()
            .to_dict()
        )
        total_all = float(sum(per_namer_raw.values())) or 0.0

        categories = GLOBAL_VISUAL_STYLE["colors"]["categories"]
        levels = GLOBAL_VISUAL_STYLE["colors"]["levels"]

        fig = px.sunburst(
            sunburst_df,
            path=["center_figure", "namer_display", "lemma"],
            values="frequency",
            color="color_group",
            color_discrete_map={
                "Name": categories.get("Name", "#EFE4D4"),
                "Antonomasia": categories.get("Antonomasia", "#F9C691"),
                "Epithet": categories.get("Epithet", "#2F4A6D"),
            },
        )

        if fig.data:
            trace = fig.data[0]

            # Segment borders (same visual language as other sunburst)
            trace.update(
                marker=dict(
                    line=dict(
                        color="rgba(45,41,38,0.25)",
                        width=0.7,
                    )
                )
            )

            labels = list(trace["labels"])
            parents = list(trace["parents"])
            values = list(trace["values"])
            colors = list(trace["marker"]["colors"])

            # Leaf-type lookup (robust + typed)
            df_leaf_type = (
                sunburst_df.loc[:, ["namer_display", "lemma", "color_group"]]
                .dropna(subset=["namer_display", "lemma"])
                .drop_duplicates(subset=["namer_display", "lemma"])
            )
            leaf_type_by_pair: dict[tuple[str, str], str] = {
                (str(r.namer_display), str(r.lemma)): str(r.color_group)
                for r in df_leaf_type.itertuples(index=False)
            }

            # --- Customdata ---
            customdata: list[list[object]] = []
            for lab, par, val in zip(labels, parents, values):
                # Root
                if lab == figure_name and (par is None or par == ""):
                    customdata.append(["", "", str(lab), float(total_all), 1.0])
                    continue

                # Namer ring
                if par == figure_name:
                    freq_namer = float(per_namer_raw.get(lab, val) or 0.0)
                    share_namer = (freq_namer / total_all) if total_all > 0 else 0.0
                    customdata.append([str(lab), "", "", freq_namer, share_namer])
                    continue

                # Leaf
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

            # --- Minimal color patching: center + namer ring (STRUCTURE with alpha) ---
            structure_rgba = hex_color_to_rgba(levels.get("STRUCTURE", "#A6B4A0"), 0.55)
            name_color = categories.get("Name", "#EFE4D4")

            for i, (lab, par) in enumerate(zip(labels, parents)):
                if lab == figure_name and (par is None or par == ""):
                    colors[i] = name_color
                elif par == figure_name:
                    colors[i] = structure_rgba

            trace["marker"]["colors"] = colors

            # WCAG-ish text colors
            text_colors = [
                pick_accessible_text_color(
                    bg,
                    dark_text_hex=levels.get("NEUTRAL_TEXT", "#2D2926"),
                    light_text_hex="#FFFFFF",
                )
                for bg in colors
            ]
            trace["textfont"] = {"color": text_colors}

    if "fig" not in locals() or fig is None:
        print("⚠️ No figure could be created for this configuration.")
        return

    apply_global_visual_style(fig, has_axes=False)
    apply_global_visual_visibility(fig, show_axis_labels=False)

    fig.update_layout(
        title={
        "text": f"Sunburst – {figure_name} ({book_name})",
        "pad": {"t": 10},        },
        margin = {"t": 80, "l": 20, "r": 20, "b": 20},
    )

    # --- 7) Output mode (save / show / both) ---
    print()
    print("📅 How should the output be handled?")
    print("[1] Save as HTML file")
    print("[2] Show plot in browser")
    print("[3] Both")
    output_mode = ask_user_choice("> ", ["1", "2", "3"])

    output_dir = os.path.join("data", book_name, "visualization")
    os.makedirs(output_dir, exist_ok=True)

    # sanitize figure_name for filename
    safe_figure = "".join(
        c if c.isalnum() or c in ("_", "-") else "_" for c in str(figure_name)
    )
    variant_label = "sunburst_figure_types" if mode_choice == "1" else "sunburst_figure_namers"

    filename_stub = f"viz_{variant_label}_{safe_figure}"
    output_path = os.path.join(output_dir, f"{filename_stub}.html")

    if output_mode == "1":
        apply_global_visual_modebar_export(
            fig,
            output_path,
            filename_stub=filename_stub,
        )
        print("\n✅ Visualization completed.")
        print(f"📂 File saved at:\n{output_path}")

    elif output_mode == "2":
        # Display only → use the temporary file
        tmp_stub = f"viz_{uuid.uuid4().hex[:8]}"

        tmp_dir = paths.get("tmp_dir") or output_dir
        os.makedirs(tmp_dir, exist_ok=True)

        tmp_path = os.path.join(tmp_dir, f"{tmp_stub}.html")
        apply_global_visual_modebar_export(
            fig,
            tmp_path,
            filename_stub=tmp_stub,
        )
        webbrowser.open_new_tab(f"file://{os.path.abspath(tmp_path)}")
        print("🌐 The plot has been opened in your browser.")
        print(f"🧾 Temporary file created at: {tmp_path}")

    elif output_mode == "3":
        # Save and display
        apply_global_visual_modebar_export(
            fig,
            output_path,
            filename_stub=filename_stub,
        )
        print("\n✅ Visualization completed.")
        print(f"📂 File saved at:\n{output_path}")
        webbrowser.open_new_tab(f"file://{os.path.abspath(output_path)}")
        print("🌐 The plot has been opened in your browser.")

def visualize_sunburst_work_overview(paths, book_name, data):
    """
    Work-centered sunburst visualization.

    Steps:
    1. Load raw naming data (JSON + Excel fallback) via load_naming_sources_with_excel_fallback.
    2. Normalize via prepare_naming_data(book_name, df_json, df_excel).
    3. Ask for Top-K (default: 12).
    4. Build aggregated DataFrame (root = work, ring 1 = figure, ring 2 = type).
    5. Create Plotly sunburst.
    6. Ask for output mode (save / show / both) and handle export.
    """

    # --- 1) load naming data via central loader ---
    df_json, df_excel = load_naming_sources_with_excel_fallback(paths, data)

    # --- 2) normalization via prepare_naming_data ---
    _source, df, cols = prepare_naming_data(book_name, df_json, df_excel)
    if df is None or df.empty:
        print("⚠️ No naming data available after prepare_naming_data.")
        return

    # --- 3) Top-K selection (CLI) ---
    print()
    print("ℹ Top figures will be selected automatically based on total naming frequency.")
    default_top_k = 12
    top_k = default_top_k

    while True:
        raw = input(
            f"Enter number of top figures to include  [Press Enter to use default: {default_top_k}]:\n> "
        ).strip()

        if raw == "":
            # keep default_top_k
            break

        try:
            value = int(raw)
            if value <= 0:
                print("Please enter a positive integer or press Enter for the default.")
                continue

            top_k = value
            break

        except ValueError:
            print("Please enter a valid integer or press Enter for the default.")

    # --- 4) build aggregated data for work-centered sunburst ---
    sunburst_df = build_sunburst_data_work_overview(df, cols, top_k, book_name)
    if sunburst_df is None or sunburst_df.empty:
        print("⚠️ No data available for work-centered sunburst overview.")
        return

    # --- 5) Create Plotly sunburst (work overview) ---

    categories = GLOBAL_VISUAL_STYLE["colors"]["categories"]
    levels = GLOBAL_VISUAL_STYLE["colors"]["levels"]

    # Use your poster palette (defensive: falls Keys fehlen)
    eigenname_color = categories.get("Eigenname", "#EFE4D4")
    antonomasie_color = categories.get("Antonomasie", "#F9C691")
    epitheton_color = categories.get("Epitheta", categories.get("Epitheton", "#2F4A6D"))

    type_color_map = {
        # German dataset labels
        "Eigenname": eigenname_color,
        "Antonomasie": antonomasie_color,
        "Epitheton": epitheton_color,
        "Epitheta": epitheton_color,

        # optional English aliases (falls später mal genutzt)
        "Name": eigenname_color,
        "Antonomasia": antonomasie_color,
        "Epithet": epitheton_color,
        "Epithets": epitheton_color,
    }

    fig = px.sunburst(
        sunburst_df,
        path=["root", "figure", "type"],
        values="value",
        color="type",
        color_discrete_map=type_color_map,
    )

    apply_global_visual_style(fig, has_axes=False)
    apply_global_visual_visibility(fig, show_axis_labels=False)

    # --- Optional: enforce colors on the trace (Plotly can fall back to defaults) ---
    if fig.data:
        trace = fig.data[0]

        labels = list(trace["labels"])
        parents = list(trace["parents"])
        colors = list(trace["marker"]["colors"])

        root_label = str(book_name)

        # "Abschwächung" wie Namer-Ring: STRUCTURE + Alpha
        fig_ring_color = hex_color_to_rgba(levels["STRUCTURE"], 0.55)

        # Root eher ruhig (z.B. Eigenname/Background)
        root_color = eigenname_color

        for i, (lab, par) in enumerate(zip(labels, parents)):
            # Root node (work)
            if str(lab) == root_label and (par is None or par == ""):
                colors[i] = root_color

            # Ring 1: figures (children of root)
            elif str(par) == root_label:
                colors[i] = fig_ring_color

            # Ring 2: types -> already handled by color_discrete_map
            # (Optional fallback, falls Plotly doch Default setzt)
            elif str(lab) in type_color_map:
                colors[i] = type_color_map[str(lab)]

        trace["marker"]["colors"] = colors

        # separators / "lines between fields"
        trace.update(
            marker=dict(
                line=dict(
                    color="rgba(45,41,38,0.22)",  # NEUTRAL_TEXT w/ alpha
                    width=0.8,
                )
            )
        )

    apply_global_visual_style(fig, has_axes=False)
    apply_global_visual_visibility(fig, show_axis_labels=False)

    fig.update_layout(
        title={
        "text": f"Sunburst – {book_name} (work overview)",
        "pad": {"t": 10},        },
        margin = {"t": 80, "l": 20, "r": 20, "b": 20},
    )

    # --- 6) output mode (save / show / both) ---
    print()
    print("📅 How should the output be handled?")
    print("[1] Save as HTML file")
    print("[2] Show plot in browser")
    print("[3] Both")
    output_mode = ask_user_choice("> ", ["1", "2", "3"])

    output_dir = os.path.join("data", book_name, "visualization")
    os.makedirs(output_dir, exist_ok=True)

    variant_label = "sunburst_work"
    filename_stub = f"viz_{variant_label}_{book_name}"
    output_path = os.path.join(output_dir, f"{filename_stub}.html")

    if output_mode == "1":
        apply_global_visual_modebar_export(
            fig,
            output_path,
            filename_stub=filename_stub,
        )
        print("\n✅ Visualization completed.")
        print(f"📂 File saved at:\n{output_path}")

    elif output_mode == "2":
        tmp_stub = f"viz_{uuid.uuid4().hex[:8]}"

        tmp_dir = paths.get("tmp_dir") or output_dir
        os.makedirs(tmp_dir, exist_ok=True)

        tmp_path = os.path.join(tmp_dir, f"{tmp_stub}.html")

        apply_global_visual_modebar_export(
            fig,
            tmp_path,
            filename_stub=tmp_stub,
        )
        webbrowser.open_new_tab(f"file://{os.path.abspath(tmp_path)}")
        print("🌐 The plot has been opened in your browser.")
        print(f"🧾 Temporary file created at: {tmp_path}")

    elif output_mode == "3":
        apply_global_visual_modebar_export(
            fig,
            output_path,
            filename_stub=filename_stub,
        )
        print("\n✅ Visualization completed.")
        print(f"📂 File saved at:\n{output_path}")
        webbrowser.open_new_tab(f"file://{os.path.abspath(output_path)}")
        print("🌐 The plot has been opened in your browser.")

def resolve_name_lemmas_for_figure(df, cols, figure_name):
    """
    Resolve which lemmas should be treated as the proper name (Eigenname)
    of a given figure for visualization purposes.

    Logic:
    1. Collect all lemmas (designations and epithets) for the target figure.
    2. Use match_name_to_lemma to find any direct name-based matches.
    3. If none found:
        - Suggest the closest lemma via difflib
        - Ask the user whether this is a variant of the name.
    4. Returns a set of lemmas treated as 'Eigenname' ONLY for this visualization run.
    """

    target_col = cols.get("target")
    designation_cols_all = cols.get("designation_cols", [])
    designation_cols = [
        c for c in designation_cols_all
        if str(c).strip().lower() != "bezeichnung"
    ]
    epithet_cols = cols.get("epithet_cols", [])

    if not target_col:
        return set()

    dff = df[df[target_col].astype(str).str.strip() == str(figure_name).strip()].copy()
    dff = dff.reset_index(drop=True)

    lemmas_all = set()
    name_lemmas = set()

    # 1) Direct name matching
    for _, row in dff.iterrows():
        # designations
        for col in designation_cols:
            val = row.get(col)
            if not isinstance(val, str):
                continue
            lemma = val.strip()
            if not lemma:
                continue
            lemmas_all.add(lemma)
            try:
                if match_name_to_lemma(figure_name, lemma, aliases=None):
                    name_lemmas.add(lemma)
            except (TypeError, ValueError, AttributeError):
                pass

        # epithets (optional for matching)
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
                    name_lemmas.add(lemma)
            except (TypeError, ValueError, AttributeError):
                pass

    # 2) If direct matches exist: done
    if name_lemmas:
        return name_lemmas

    # 3) Suggest a close match if none found
    if not lemmas_all:
        return set()

    lowered = {lm.lower(): lm for lm in lemmas_all}
    best = difflib.get_close_matches(figure_name.lower(), list(lowered.keys()), n=1, cutoff=0.6)

    if not best:
        return set()

    suggested = lowered[best[0]]

    print(f'{figure_name} could not be found as a name-based lemma.')
    yn = input(f'Could "{suggested}" be a variant of the name? (y/n) ').strip().lower()

    if yn == "y":
        return {suggested}

    return set()

def build_sunburst_data_types_lemma(df, cols, figure_name):
    """
    Figure-centered: Type → Lemma.
    Center = figure_name.
    """
    target_col = cols.get("target")

    # use only normalized designation columns (Bezeichnung 1–4)
    designation_cols_all = cols.get("designation_cols", [])
    designation_cols = [
        c for c in designation_cols_all
        if str(c).strip().lower() != "bezeichnung"
    ]

    epithet_cols = cols.get("epithet_cols", [])

    if target_col is None:
        raise ValueError("Column mapping 'target' is missing in cols.")

    # Determine which lemmas count as proper names
    name_lemmas = resolve_name_lemmas_for_figure(df, cols, figure_name)

    dff = df[df[target_col].astype(str).str.strip() == str(figure_name).strip()].copy()
    dff = dff.reset_index(drop=True)

    counts = Counter()

    for _, row in dff.iterrows():
        used_lemmas = set()

        # --- designations ---
        for col in designation_cols:
            val = row.get(col)
            if not isinstance(val, str):
                continue
            lemma = val.strip()
            if not lemma or lemma in used_lemmas:
                continue
            used_lemmas.add(lemma)

            # Internal type logic
            if lemma in name_lemmas:
                type_label = "Eigenname"
            else:
                type_label = "Antonomasie"

            counts[(type_label, lemma)] += 1

        # --- epithets ---
        for col in epithet_cols:
            val = row.get(col)
            if not isinstance(val, str):
                continue
            lemma = val.strip()
            if not lemma or lemma in used_lemmas:
                continue
            used_lemmas.add(lemma)

            counts[("Epitheton", lemma)] += 1

    # --- BUILD OUTPUT ROWS ---
    rows = []
    for (type_label, lemma), freq in counts.items():

        # Map to type_group (ring 1)
        if type_label in ("Eigenname", "Antonomasie"):
            type_group = "Naming variants"
        else:
            type_group = "Epithets"

        # Map to color_group (visual category)
        if type_label == "Eigenname":
            color_group = "Name"
        elif type_label == "Antonomasie":
            color_group = "Antonomasia"
        else:
            color_group = "Epithet"

        rows.append(
            {
                "center_figure": figure_name,
                "type_group": type_group,
                "color_group": color_group,
                "lemma": lemma,
                "frequency": freq,
            }
        )

    return pd.DataFrame(rows)

def build_sunburst_data_namer_lemma(df, cols, figure_name):
    """
    Figure-centered: Namer → Type → Lemma.
    Center = figure_name.
    """
    target_col = cols.get("target")
    namer_col = cols.get("namer")
    designation_cols_all = cols.get("designation_cols", [])
    designation_cols = [
        c for c in designation_cols_all
        if str(c).strip().lower() != "bezeichnung"
    ]
    epithet_cols = cols.get("epithet_cols", [])

    if target_col is None:
        raise ValueError("Column mapping 'target' is missing in cols.")
    if namer_col is None:
        raise ValueError("Column mapping 'namer' is missing in cols.")

    # Try to find an Erzähler column
    narrator_col = None
    for key in ("narrator", "narrator_col"):
        if key in cols:
            narrator_col = cols[key]
            break
    if narrator_col is None:
        for c in df.columns:
            if str(c).strip().lower() in ("erzähler", "erzaehler", "narrator"):
                narrator_col = c
                break

    # NEW: Name-matching logic
    name_lemmas = resolve_name_lemmas_for_figure(df, cols, figure_name)

    dff = df[df[target_col].astype(str).str.strip() == str(figure_name).strip()].copy()
    dff = dff.reset_index(drop=True)

    counts = Counter()

    for _, row in dff.iterrows():
        raw_namer = row.get(namer_col)
        namer = raw_namer.strip() if isinstance(raw_namer, str) else ""

        if not namer and narrator_col is not None:
            raw_narr = row.get(narrator_col)
            namer = raw_narr.strip() if isinstance(raw_narr, str) else ""

        if not namer:
            continue

        used_lemmas = set()

        # --- designations ---
        for col in designation_cols:
            val = row.get(col)
            if not isinstance(val, str):
                continue
            lemma = val.strip()
            if not lemma or lemma in used_lemmas:
                continue
            used_lemmas.add(lemma)

            type_label = "Antonomasie"

            if lemma in name_lemmas:
                type_label = "Eigenname"
            else:
                try:
                    if match_name_to_lemma(figure_name, lemma, aliases=None):
                        type_label = "Eigenname"
                except (TypeError, ValueError, AttributeError):
                    pass

            counts[(namer, type_label, lemma)] += 1

        # --- epithets ---
        for col in epithet_cols:
            val = row.get(col)
            if not isinstance(val, str):
                continue
            lemma = val.strip()
            if not lemma or lemma in used_lemmas:
                continue

            used_lemmas.add(lemma)
            counts[(namer, "Epitheton", lemma)] += 1

    rows = []
    for (namer, type_label, lemma), freq in counts.items():
        # Map internal type_label (Eigenname / Antonomasie / Epitheton)
        # to English group labels for the sunburst
        if type_label in ("Eigenname", "Antonomasie"):
            type_group = "Naming variants"
        else:
            type_group = "Epithets"

        if type_label == "Eigenname":
            color_group = "Name"
        elif type_label == "Antonomasie":
            color_group = "Antonomasia"
        else:
            color_group = "Epithet"

        rows.append(
            {
                "center_figure": figure_name,
                "namer": namer,  # ← neu: nennende Figur
                "type_group": type_group,  # ring 1 (Naming variants / Epithets)
                "color_group": color_group,  # Name / Antonomasia / Epithet
                "type": type_label,  # ← neu: interne Typ-Bezeichnung
                "lemma": lemma,
                "frequency": freq,
            }
        )

    return pd.DataFrame(rows)

def build_sunburst_data_work_overview(df, cols, top_k, book_name):
    """
    Work-centered: Work → Figure → Type.
    """
    target_col = cols.get("target")
    designation_cols = cols.get("designation_cols", [])
    epithet_cols = cols.get("epithet_cols", [])

    if target_col is None:
        raise ValueError("Column mapping 'target' is missing in cols.")

    total_counts = Counter()
    type_counts = Counter()

    for _, row in df.iterrows():
        raw_target = row.get(target_col)
        figure = raw_target.strip() if isinstance(raw_target, str) else ""
        if not figure:
            continue

        total_counts[figure] += 1
        types_in_row = set()

        # designations
        for col in designation_cols:
            val = row.get(col)
            if not isinstance(val, str):
                continue
            lemma = val.strip()
            if not lemma:
                continue
            type_label = "Antonomasie"
            try:
                if match_name_to_lemma(figure, lemma, aliases=None):
                    type_label = "Eigenname"
            except (TypeError, ValueError, AttributeError):
                pass
            types_in_row.add(type_label)

        # epithets
        for col in epithet_cols:
            val = row.get(col)
            if not isinstance(val, str):
                continue
            lemma = val.strip()
            if not lemma:
                continue
            types_in_row.add("Epitheton")

        for t in types_in_row:
            type_counts[(figure, t)] += 1

    if not total_counts:
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

    sorted_figures = sorted(
        total_counts.items(),
        key=lambda item: (-item[1], str(item[0]).lower()),
    )

    if top_k is not None and top_k > 0:
        top_figures = {name for name, _ in sorted_figures[:top_k]}
    else:
        top_figures = {name for name, _ in sorted_figures}

    per_figure_sums = Counter()
    for (figure, t), count in type_counts.items():
        if figure in top_figures:
            per_figure_sums[figure] += count

    rows = []
    for (figure, t), count in type_counts.items():
        if figure not in top_figures:
            continue

        total = per_figure_sums.get(figure, 0)
        pct = count / total if total > 0 else 0.0

        rows.append(
            {
                "root": book_name,
                "figure": figure,
                "type": t,
                "value": count,
                "total_for_figure": total,
                "pct_of_figure": pct,
            }
        )

    return pd.DataFrame(rows)