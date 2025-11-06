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
import os
import math
import csv
import difflib
import numpy as np
import uuid
import webbrowser
from collections import Counter
from typing import List

import pandas as pd
import plotly.express as px

from naming_analysis.shared import (
    ask_user_choice,
    get_first_valid_text
)
from naming_analysis.io_utils import safe_read_json
from naming_analysis.loaders import load_collocation_sheet, build_fallback_collocation_df_from_tei
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
            run_visualization_menu(paths, book_name)
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

    # --- load categorization JSON safely ---
    df_json = None
    try:
        json_path = paths.get("categorization_json")
        js = safe_read_json(json_path, default=[])
        if isinstance(js, list):
            df_json = pd.DataFrame(js)
        elif isinstance(js, dict):
            for key in ("entries", "data", "items"):
                if key in js and isinstance(js[key], list):
                    df_json = pd.DataFrame(js[key])
                    break
    except Exception as e:
        print(f"⚠️ Could not load categorization JSON: {e}")

    # --- load Excel fallback into DataFrame ---
    df_excel = None
    try:
        # identical to collocation loader
        for k in ("excel", "excel_df", "df_excel"):
            if k in data and isinstance(data[k], pd.DataFrame):
                df_excel = data[k]
                break
        # ensure the in-memory Excel DataFrame actually comes from the 'lemmatisiert' sheet if available
        excel_path = paths.get("excel_path") or data.get("excel_path")
        if excel_path and os.path.exists(excel_path):
            try:
                xls = pd.ExcelFile(excel_path)
                sheets_lower = [s.strip().lower() for s in xls.sheet_names]
                if "lemmatisiert" in sheets_lower:
                    # check whether current df_excel columns look like the lemmatized structure
                    looks_lemmatized = any("bezeichnung 1" in c.lower() for c in df_excel.columns)
                    if not looks_lemmatized:
                        df_excel = pd.read_excel(excel_path, sheet_name="lemmatisiert", dtype=str)
            except Exception as e:
                print(f"⚠️ Could not verify or switch to 'lemmatisiert' sheet: {e}")

        if df_excel is None:
            excel_path = paths.get("excel_path") or data.get("excel_path")
            if excel_path and os.path.exists(excel_path):
                try:
                    # prefer the sheet 'lemmatisiert' if present
                    xls = pd.ExcelFile(excel_path)
                    if "lemmatisiert" in [s.strip().lower() for s in xls.sheet_names]:
                        df_excel = pd.read_excel(excel_path, sheet_name="lemmatisiert", dtype=str)
                        print(f"ℹ️ Excel sheet 'lemmatisiert' loaded from: {excel_path} ({len(df_excel)} rows).")
                    else:
                        df_excel = pd.read_excel(excel_path, dtype=str)
                        print(f"ℹ️ Excel default sheet loaded from: {excel_path} ({len(df_excel)} rows).")
                except Exception as e:
                    print(f"⚠️ Could not load Excel file: {e}")
            else:
                print("⚠️ Excel path not found or missing.")
    except Exception as e:
        print(f"⚠️ Could not load Excel fallback: {e}")

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

def run_visualization_menu(paths, book_name):
    while True:
        print("\n📈 Which visualization do you want to run?")
        print("[1] Verse-based naming distribution")
        print("[2] Intra-Figure Co-Occurrence Heatmap")
        print("[3] Back to analysis menu")

        choice = ask_user_choice("> ", ["1", "2", "3"])

        if choice == "1":
            visualize_verse_naming_distribution(paths, book_name)
        elif choice == "2":
            visualize_intra_figure_cooccurrence_heatmap(paths, book_name)
        elif choice == "3":
            print("↩️ Returning to analysis menu.")
            break

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

    all_entries = []
    for col in selected_cols:
        temp = df_figure[["Vers", col]].dropna().rename(columns={col: "Token"})
        all_entries.append(temp)

    df_combined = pd.concat(all_entries)
    df_combined["Token"] = df_combined["Token"].astype(str).str.strip()
    df_combined["Vers"] = pd.to_numeric(df_combined["Vers"], errors="coerce")

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
            input_str = input("\n✍ Which naming variants should be included? (e.g., 1–3, 5)\n> ").strip()
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
            input_str = input("\n✍ Which epithets should be included? (e.g., 1–3, 5)\n> ").strip()
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
            lambda x: "Naming variant" if x in selected_naming else "Epitheton"
        )
        color_column = "Category"
    else:
        color_column = "Token_html"

    # Step 7 – Create interactive plot
    fig = px.scatter(
        df_plot,
        x="Vers",
        y="Token_html",
        color=color_column,
        title=f"Visualization for '{figure_name}'",
        hover_data=["Vers", "Token"]
    )

    fig.update_traces(marker=dict(size=18, opacity=0.7))

    fig.update_layout(
        title=dict(
            text=f"Visualization for {variant_label} '{figure_name}'",
            x=0.5,
            xanchor="center"
        ),
        xaxis_title="Verses",
        yaxis_title=variant_label,
        font=dict(
            family="Times New Roman",
            size=18
        ),
        xaxis=dict(
            tickfont=dict(size=36)
        ),
        yaxis=dict(
            tickfont=dict(size=36),
            categoryorder="array",
            categoryarray=[f"<i>{t}</i>" for t in sorted_tokens]
        ),
        height=800,
        margin=dict(l=100, r=40, t=60, b=60),
        showlegend=(variant_type == "3")
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
        fig.write_html(output_path)
        print(f"\n✅ Visualization completed.")
        print(f"📂 File saved at:\n{output_path}")

    elif output_mode == "2":
        # Display only → use the temporary file
        tmp_filename = f"viz_{uuid.uuid4().hex[:8]}.html"
        tmp_path = os.path.join(paths["tmp_dir"], tmp_filename)
        fig.write_html(tmp_path)
        webbrowser.open_new_tab(f"file://{os.path.abspath(tmp_path)}")
        print(f"🌐 The plot has been opened in your browser.")
        print(f"🧾 Temporary file created at: {tmp_path}")

    elif output_mode == "3":
        # Save and display → use the saved file
        fig.write_html(output_path)
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
    from itertools import combinations
    from collections import Counter
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
    top_pairs = pair_counter.most_common(top_n) if top_n and top_n > 0 else pair_counter.most_common()
    tokens = sorted(set([t for p, _ in top_pairs for t in p]))
    index = {t: i for i, t in enumerate(tokens)}

    # Build symmetric matrix, but display only one half
    size = len(tokens)
    matrix = np.zeros((size, size), dtype=float)

    for (a, b), c in top_pairs:
        i, j = index[a], index[b]
        if i > j:  # only fill the lower half
            matrix[i, j] = c
        elif i < j:
            matrix[j, i] = c
        # diagonal remains 0

    # set all cells above the diagonal to NaN → invisible in Plotly
    matrix[np.triu_indices(size, k=1)] = np.nan

    fig = px.imshow(
        matrix,
        x=tokens,
        y=tokens,
        labels=dict(x="Token", y="Token", color="Co-occurrences"),
        aspect="auto"
    )
    fig.update_traces(hovertemplate="Token %{y} × %{x}<br>Co-occurrences: %{z}<extra></extra>")

    # Plot
    fig = px.imshow(
        matrix,
        x=tokens,
        y=tokens,
        labels=dict(x="Token", y="Token", color="Co-occurrences"),
        aspect="auto"
    )

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

    if output_mode == "1":
        fig.write_html(output_path)
        print("\n✅ Visualization completed.")
        print(f"📂 File saved at:\n{output_path}")

    elif output_mode == "2":
        tmp_dir = paths.get("tmp_dir", os.path.join("data", book_name, "tmp"))
        os.makedirs(tmp_dir, exist_ok=True)
        tmp_filename = f"viz_{uuid.uuid4().hex[:8]}.html"
        tmp_path = os.path.join(tmp_dir, tmp_filename)
        fig.write_html(tmp_path)
        webbrowser.open_new_tab(f"file://{os.path.abspath(tmp_path)}")
        print("🌐 The plot has been opened in your browser.")
        print(f"🧾 Temporary file created at: {tmp_path}")

    elif output_mode == "3":
        fig.write_html(output_path)
        print("\n✅ Visualization completed.")
        print(f"📂 File saved at:\n{output_path}")
        webbrowser.open_new_tab(f"file://{os.path.abspath(output_path)}")
        print("🌐 The plot has been opened in your browser.")