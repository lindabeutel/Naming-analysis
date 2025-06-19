"""
collection.py

This module contains the core data collection and annotation logic
for the naming analysis pipeline. It includes:
- TEI-based extraction of missing naming variants,
- interactive collection of collocations,
- and lemma-based categorization into naming variants (a) and epithets (e).

All functionality in this module is designed for interactive use during
the data collection phase of the project.
"""
import re
import json
import pandas as pd
from xml.etree.ElementTree import Element

from naming_analysis.shared import (
    normalize_text,
    get_first_valid_text,
    clean_cell_value,
    sanitize_cell_value,
    ask_user_choice
)
from naming_analysis.tei_utils import tei_ns, get_valid_verse_number, get_verse_context
from naming_analysis.io_utils import safe_write_json
from naming_analysis.loaders import (
    load_lemma_normalization,
    load_ignored_lemmas,
    load_lemma_categories,
    load_json_annotations
)
from naming_analysis.savers import (
    save_progress,
    save_lemma_normalization,
    save_ignored_lemmas,
    save_lemma_categories
)

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
    Runs the data collection process depending on active modes:
    - If check_naming_variants is True → TEI-based loop.
    - If only collocations and/or categorization are active → Excel-based loop.
    Returns updated (missing_naming_variants, collocation_data, categorized_entries).
    """

    # --- TEI-based loop (only when check_naming_variants is active)
    if check_naming_variants:
        if root is None:
            print("⚠️ No TEI root found – cannot perform TEI-based iteration.")
            return missing_naming_variants, collocation_data, categorized_entries

        verse = root.findall('.//tei:l', tei_ns)
        if not verse:
            print("⚠️ No verses found in TEI.")
            return missing_naming_variants, collocation_data, categorized_entries

        start_index = next(
            (i for i, line in enumerate(verse) if get_valid_verse_number(line.get("n")) > last_verse),
            0
        )

        print(f"🔁 Starting TEI iteration from verse {int(verse[start_index].get('n'))} (Index {start_index})")

        for line in verse[start_index:]:
            verse_number = get_valid_verse_number(line.get("n"))

            verse_text = ' '.join([seg.text for seg in line.findall(".//tei:seg", tei_ns) if seg.text])
            normalized_verse = normalize_text(verse_text)

            # Naming detection
            missing_naming_variants = check_and_extend_namings(
                verse_number,
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

            # Collocations
            if perform_collocations:
                rows = df[df["Vers"] == verse_number]

                for _, row in rows.iterrows():
                    check_and_add_collocations(
                        verse_number, collocation_data, root, paths, row=row
                    )

            # Categorization
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
                        if int(e.get("Vers", -1)) != verse_number:
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

            # Save progress after each verse
            save_progress(
                missing_naming_variants=missing_naming_variants,
                last_processed_verse=verse_number,
                paths=paths,
                check_naming_variants=check_naming_variants,
                perform_collocations=perform_collocations,
                perform_categorization=perform_categorization
            )

    # --- Excel-based loop (when only collocations and/or categorization are active)
    elif perform_collocations or perform_categorization:
        print("🔁 Starting EXCEL-based iteration over 'Vers' list.")

        # Extract and sort valid verse numbers from Excel
        vers_list = sorted(set(
            v for v in df["Vers"] if str(v).strip().isdigit()
        ))

        for v in vers_list:
            verse_number = get_valid_verse_number(v)

            # Collocations
            if perform_collocations:
                rows = df[df["Vers"] == verse_number]

                for _, row in rows.iterrows():
                    check_and_add_collocations(
                        verse_number, collocation_data, root, paths, row=row
                    )
            # Categorization
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
                        if int(e.get("Vers", -1)) != verse_number:
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

            # Save progress after each verse
            save_progress(
                missing_naming_variants=missing_naming_variants,
                last_processed_verse=verse_number,
                paths=paths,
                check_naming_variants=check_naming_variants,
                perform_collocations=perform_collocations,
                perform_categorization=perform_categorization
            )

    # Return updated data
    return missing_naming_variants, collocation_data, categorized_entries

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
    Identifies potential naming variants in the current TEI verse that are missing
    in the Excel file and not yet confirmed or rejected.

    The user is prompted interactively to confirm, assign a role (e.g., Eigennennung),
    and optionally define collocation lines. Confirmed entries can also be immediately
    categorized (if enabled).

    Returns:
        list: The updated list of missing naming variants (with confirmed/rejected entries).
    """
    # 1. Extract naming variants from Excel for the current verse
    existing_naming_variants = set()
    if "Vers" in df.columns:
        df_verse = df[df["Vers"] == verse_number]
        for column in ["Eigennennung", "Bezeichnung", "Erzähler"]:
            if column in df_verse.columns:
                values = df_verse[column].dropna().tolist()
                existing_naming_variants.update(
                    normalize_text(str(value).strip()) for value in values if str(value).strip()
                )

    # 2. Extract and normalize naming variants from dict
    dict_naming_variants = set()
    for book_list in naming_variants_dict.get("Namings", {}).values():
        dict_naming_variants.update(
            normalize_text(name.strip()) for name in book_list if name.strip()
        )

    # 3. Match check and user interaction
    for naming_variant in dict_naming_variants:
        if not naming_variant:
            continue

        # skip if already handled in Excel (auch als Token-Menge)
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

        # skip if already handled in JSON
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

        if not re.search(rf'\b{re.escape(naming_variant)}\b', normalized_verse):
            continue

        print("\n" + "-" * 60)
        print(f"❗ New naming variant found that is not listed in the Excel file!")
        print(f"🔍 Detected naming variant: \"{naming_variant}\"")

        # 📖 Show context
        prev_line = root.find(f'.//tei:l[@n="{verse_number - 1}"]', tei_ns)
        if prev_line is not None:
            prev_text = ' '.join([seg.text for seg in prev_line.findall('.//tei:seg', tei_ns) if seg.text])
            print(f"📖 Previous verse ({verse_number - 1}): {prev_text}")

        highlighted = verse_text.replace(naming_variant, f"\033[1m\033[93m{naming_variant}\033[0m")
        print(f"📖 Verse ({verse_number}): {highlighted}")

        next_line = root.find(f'.//tei:l[@n="{verse_number + 1}"]', tei_ns)
        if next_line is not None:
            next_text = ' '.join([seg.text for seg in next_line.findall('.//tei:seg', tei_ns) if seg.text])
            print(f"📖 Next verse ({verse_number + 1}): {next_text}")

        # 🧍 Confirm with user
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
            print("✅ Rejection saved.")
            continue

        extend = ask_user_choice("💡 Would you like to shorten or lengthen the naming variant (y/n): ", ["y", "n"])
        if extend == "y":
            naming_variant = input("✍ Enter the adapted naming variant: ").strip()

        print("Please choose the correct category:")
        print("[1] Eigennennung")
        print("[2] Bezeichnung")
        print("[3] Erzähler")
        print("[4] Skip")

        choice = input("👉 Your selection: ").strip()
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

        # 📌 Optional: add collocation
        wants_collocation = ask_user_choice("📌 Do you want to add a collocation (context lines)? (y/n): ", ["y", "n"])
        if wants_collocation == "y":
            print("\n📖 Extended context (1–13):")
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

            selection = input("\n👉 Please enter the line number(s) (e.g., '5-7' or '6'): ").strip()
            selected = []

            try:
                if "-" in selection:
                    start, end = map(int, selection.split("-"))
                    selected = [context_lines[i] for i in range(start, end + 1) if i in context_lines]
                else:
                    idx = int(selection)
                    selected = [context_lines[idx]]
            except (ValueError, KeyError):
                print("⚠️ Invalid input – no collocation saved.")

            if selected:
                entry["Kollokation"] = ' / '.join(selected)

        missing_naming_variants.append(entry)
        save_progress(missing_naming_variants, verse_number, paths)
        print("✅ Entry saved.")

        # 🆕 Sofortige Kategorisierung, falls aktiviert und bestätigt
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

def check_and_add_collocations(verse_number, collocation_data, root, paths, row):
    """
    Interactively collects a collocation context for a given naming variant
    if the Excel field is currently empty and no prior entry exists in the JSON data.

    Uses verse context from the TEI tree and prompts the user to select
    relevant lines.

    Returns:
        bool | None: True if a new collocation was added, None otherwise.
    """
    # Check if already handled via Excel
    if sanitize_cell_value(row.get("Kollokationen")) != "":
        return None

    naming_variant = clean_cell_value(row.get("Eigennennung")) \
             or clean_cell_value(row.get("Bezeichnung")) \
             or clean_cell_value(row.get("Erzähler"))

    named_entity = clean_cell_value(row.get("Benannte Figur"))

    context = get_verse_context(verse_number, root)

    collocations = ask_for_collocations(verse_number, named_entity, naming_variant, context)

    # Check if already handled via JSON
    if any(
            int(entry.get("Vers", -1)) == verse_number
            and entry.get("Benannte Figur", "") == named_entity
            and entry.get("Naming", "") == naming_variant
            and str(entry.get("Kollokationen", "")).strip()
            for entry in collocation_data
    ):
        return None

    collocation_data.append({
        "Vers": verse_number,
        "Benannte Figur": named_entity,
        "Naming": naming_variant,
        "Kollokationen": collocations
    })

    # 📝 Immediately save progress
    with open(paths["collocations_json"], "w", encoding="utf-8") as f:
        json.dump(collocation_data, f, indent=4, ensure_ascii=False)

    return True

def ask_for_collocations(verse_number, named_entity, naming_variant, context):
    """
    Displays ±6 lines of verse context around the current verse and asks
    the user to select relevant lines for the collocation.

    Returns:
        str: The concatenated collocation context selected by the user.
    """
    print(f"\n🟡 Empty collocation field detected in verse {verse_number}!")
    if named_entity or naming_variant:
        print(f"👤 {named_entity}: {naming_variant}\n")

    for number, text in context:
        if naming_variant:
            # Highlight naming_variant
            highlighted = text.replace(str(naming_variant), f"\033[1;33m{naming_variant}\033[0m")
        else:
            highlighted = text
        print(f"{number}. {highlighted}")

    while True:
        user_input = input("\n👉 Please enter the number(s) of the relevant lines (e.g., '5' or '5-7'): ")

        selected = []

        try:
            if "-" in user_input:
                start, end = map(int, user_input.split("-"))
                selected = [text for number, text in context if start <= number <= end]
            else:
                number = int(user_input)
                selected = [text for num, text in context if num == number]
            break

        except (ValueError, KeyError):
            print("⚠️ Invalid input. Please enter a single number or a range.")

    return " / ".join(selected)

def lemmatize_and_categorize_entry(entry, lemma_normalization, paths, ignored_lemmas=None, lemma_categories=None):
    """
    Annotates a naming entry with lemmatized components and categorizes
    each token as a naming variant ('a') or epithet ('e').

    The user is prompted for any unknown tokens and can revise each categorization interactively.

    Returns:
        dict | None: The annotated entry, or None if skipped.
    """
    if lemma_normalization is None:
        lemma_normalization = load_lemma_normalization(paths["lemma_normalization_json"])

    if ignored_lemmas is None:
        ignored_lemmas = load_ignored_lemmas(paths["ignored_lemmas_json"])

    if lemma_categories is None:
        lemma_categories = load_lemma_categories(paths["lemma_categories_json"])

    text = get_first_valid_text(
        entry.get("Erzähler"),
        entry.get("Bezeichnung"),
        entry.get("Eigennennung")
    )

    if not text:
        print("⚠ No text to annotate – entry skipped.\n")
        return None

    print("\n" + "=" * 60)
    print(f"▶ Verse: {entry.get('Vers')}")
    print(f"▶ Named Entity: {entry.get('Benannte Figur')}")

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

    print(f"▶ Type: {typ}")

    print(f"\n▶ Original text: {text}")

    tokens = [t for t in tokenize(text.lower()) if t.isalpha()]

    # Filter only real word tokens
    missing = [
        t for t in tokens
        if t.isalpha() and not any(t in v or t == k for k, v in lemma_normalization.items())
    ]

    if missing:
        while True:
            print(f"\n▶ Please add lemma(ta) for {', '.join(missing)} (comma-separated):")
            user_input = input("> ").strip()
            new_lemmata = [l.strip() for l in user_input.split(",") if l.strip()]
            if len(new_lemmata) == len(missing):
                break
            print(
                f"⚠ Number of lemmata ({len(new_lemmata)}) doesn't match number of tokens ({len(missing)}). Please try again.")

        for token, lemma in zip(missing, new_lemmata):
            lemma_normalization.setdefault(lemma, [])
            if token not in lemma_normalization[lemma]:
                lemma_normalization[lemma].append(token)

        # 🔤 Sort alphabetically
        for lemma in lemma_normalization:
            lemma_normalization[lemma] = sorted(set(lemma_normalization[lemma]))

        save_lemma_normalization(lemma_normalization, path=paths["lemma_normalization_json"])

    lemmata = [resolve_lemma(t, lemma_normalization) for t in tokens]
    print(f"\n▶ Lemma: {', '.join(lemmata)}\n")

    while True:
        naming_variants, epithets = run_categorization(
            lemmata, lemma_categories, ignored_lemmas, paths
        )

        if not naming_variants and not epithets:
            print("⚠ No entry – please review and confirm again.")
            confirm = ask_user_choice("Really skip this entry? [y = yes / n = no]: ", ["y", "n"])
            if confirm == "y":
                print("⏭ Entry skipped.\n")
                return None
            else:
                return run_categorization(lemmata, lemma_categories, ignored_lemmas, paths)
        else:
            break

    if lemma_normalization:
        save_lemma_normalization(lemma_normalization, path=paths["lemma_normalization_json"])

    if ignored_lemmas:
        save_ignored_lemmas(ignored_lemmas, path=paths["ignored_lemmas_json"])

    if lemma_categories:
        save_lemma_categories(lemma_categories, path=paths["lemma_categories_json"])

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

    # 💾 Kategorisierung direkt speichern
    existing = load_json_annotations(paths["categorization_json"])
    existing.append(annotated_entry)
    safe_write_json(existing, paths["categorization_json"], merge=True)

    print("✅ Entry saved.\n")
    return annotated_entry

def run_categorization(lemmata, lemma_categories, ignored_lemmas, paths):
    """
    Interactive helper for assigning lemma categories.

    For each token, the user can:
    - assign a = naming variant
    - assign e = epithet
    - ignore
    - go back and revise the previous step

    Returns:
        tuple[list[str], list[str]]: Two lists containing categorized naming variants and epithets.
    """
    while True:
        naming_variants = []
        epithets = []
        history = []
        i = 0

        while i < len(lemmata):
            lemma = lemmata[i]

            if lemma in ignored_lemmas:
                i += 1
                continue

            default = f"[{lemma_categories.get(lemma, '')}]" if lemma in lemma_categories else ""
            print(f"{lemma:<12} → {default} ", end="")
            user_input = input().strip()

            if user_input == "<":
                if i == 0 or not history:
                    print("↩️  Already at beginning – can't step back.")
                    continue
                i -= 1
                last_action = history.pop()
                if last_action["type"] == "a":
                    naming_variants.pop()
                elif last_action["type"] == "e":
                    epithets.pop()
                elif last_action["type"] == "ignore":
                    ignored_lemmas.discard(last_action["lemma"])
                    save_ignored_lemmas(ignored_lemmas, path=paths["ignored_lemmas_json"])
                elif last_action["type"] == "override":
                    del lemma_categories[last_action["lemma"]]
                    save_lemma_categories(lemma_categories, path=paths["lemma_categories_json"])
                continue

            if user_input == "" and default:
                if default == "[a]":
                    naming_variants.append(lemma)
                    history.append({"type": "a", "lemma": lemma})
                elif default == "[e]":
                    epithets.append(lemma)
                    history.append({"type": "e", "lemma": lemma})
                i += 1
                continue

            if user_input == "":
                confirm_ignore = ask_user_choice(f"⚠️ Really ignore lemma “{lemma}”? [y/n]: ", ["y", "n"])
                if confirm_ignore == "y":
                    ignored_lemmas.add(lemma)
                    save_ignored_lemmas(ignored_lemmas, path=paths["ignored_lemmas_json"])
                    print(f"ℹ️ Lemma “{lemma}” added to ignore list.")
                    history.append({"type": "ignore", "lemma": lemma})
                    i += 1
                    continue
                else:
                    print("↩️  Skipped ignoring – please choose a category or go back.\n")
                    continue

            if user_input in ("a", "e"):
                if user_input == "a":
                    naming_variants.append(lemma)
                else:
                    epithets.append(lemma)
                lemma_categories[lemma] = user_input
                save_lemma_categories(lemma_categories, path=paths["lemma_categories_json"])
                history.append({"type": user_input, "lemma": lemma})
                i += 1
                continue

            correction = user_input
            cat = ""
            while cat not in ("a", "e"):
                cat = input(f'Define category for “{correction}” [a/e]: ').strip().lower()

            if cat == "a":
                naming_variants.append(correction)
            else:
                epithets.append(correction)

            lemma_categories[correction] = cat
            save_lemma_categories(lemma_categories)
            history.append({"type": "override", "lemma": correction})
            i += 1

        return naming_variants, epithets

def tokenize(text):
    """
    Splits a string into individual tokens using a simple regular expression.

    Tokens include:
    - Words (alphanumeric sequences)
    - Individual punctuation characters

    Used during lemma categorization to break a naming string into discrete components.

    Parameters:
        text (str): The input string to tokenize.

    Returns:
        list[str]: A list of tokens extracted from the input text.
    """
    return re.findall(r'\w+|[^\w\s]', text, re.UNICODE)

def resolve_lemma(token: str, lemma_dict: dict[str, list[str]]) -> str:
    """
    Resolves a token to its corresponding lemma using a mapping dictionary.

    The dictionary must follow the structure:
        {lemma: [variant1, variant2, ...]}

    If the token matches any listed variant, the corresponding lemma is returned.
    If no match is found, the token is returned as-is (fallback).

    Parameters:
        token (str): The word form to resolve.
        lemma_dict (dict): A mapping of lemma → variant list.

    Returns:
        str: The resolved lemma or the original token if not found.
    """
    for lemma, variants in lemma_dict.items():
        if token in variants:
            return lemma
    return token  # fallback if no variant matches