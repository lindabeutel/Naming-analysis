"""
controller.py

This module coordinates the execution flow of the naming analysis pipeline.
It organizes project setup, data collection, progress saving, and optional export or analysis.

The functions here are designed to be used directly from the main entry point (run.py).
"""

from naming_analysis.project_setup import initialize_project
from naming_analysis.loaders import load_or_extend_naming_variants_dict
from naming_analysis.loaders import (
    load_lemma_normalization,
    load_ignored_lemmas,
    load_lemma_categories
)
from naming_analysis.collection import run_data_collection
from naming_analysis.savers import save_progress
from naming_analysis.io_utils import load_missing_naming_variants
from naming_analysis.exporter import export_all_data_to_new_excel
from naming_analysis.analysis import run_analysis_menu
from naming_analysis.shared import ask_user_choice
from naming_analysis.loaders import (
    load_collocations_json,
    load_json_annotations
)

def setup_project_session():
    """
    Initializes the session by selecting a book, loading paths, config and data.

    Returns:
        tuple:
            book_name (str)
            config_data (dict)
            data (dict)
            paths (dict)
            active_last_verse (int)
            mode_flags (dict): contains the mode booleans for each operation
    """
    book_name, naming_variants_last_verse, collocations_last_verse, categorization_last_verse, paths = initialize_project()
    naming_variants_dict = load_or_extend_naming_variants_dict()

    from naming_analysis.config import ask_config_interactively
    config_data, data = ask_config_interactively(paths["config_json"])
    paths["original_excel"] = data.get("excel_path")

    mode_flags = {
        "check_naming_variants": config_data.get("check_naming_variants", False),
        "perform_collocations": config_data.get("fill_collocations", False),
        "perform_categorization": config_data.get("do_categorization", False)
    }

    if config_data.get("modus") == "analyze":
        run_analysis_menu(config_data, paths, data, book_name)
        exit(0)

    elif config_data.get("modus") == "export":
        options = {
            "benennungen": config_data.get("check_naming_variants", True),
            "kollokationen": config_data.get("fill_collocations", True),
            "kategorisierung": config_data.get("do_categorization", True)
        }
        paths["original_excel"] = data.get("excel_path")
        export_all_data_to_new_excel(book_name, paths, options)

        analyze = ask_user_choice("Do you want to run an analysis now? (y/n): ", ["y", "n"])
        if analyze == "y":
            run_analysis_menu(config_data, paths, data, book_name)

        exit(0)

    if mode_flags["check_naming_variants"]:
        last_verse = naming_variants_last_verse
    elif mode_flags["perform_collocations"]:
        last_verse = collocations_last_verse
    elif mode_flags["perform_categorization"]:
        last_verse = categorization_last_verse
    else:
        last_verse = 0

    return book_name, config_data, data, paths, last_verse, mode_flags, naming_variants_dict

def run_data_workflow(book_name, config_data, data, paths, last_verse, mode_flags, naming_variants_dict):
    """
    Executes the data collection workflow using the selected mode(s).

    Returns:
        tuple: (missing_naming_variants, collocation_data, categorized_entries)
    """
    df = data.get("excel")
    root = data.get("xml")

    missing_naming_variants = load_missing_naming_variants(paths["missing_naming_variants_json"])
    collocation_data = load_collocations_json(paths["collocations_json"])
    categorized_entries = load_json_annotations(paths["categorization_json"])

    previous_naming_variants = missing_naming_variants.copy()
    previous_collocations = collocation_data.copy()
    previous_categorized_entries = categorized_entries.copy()

    lemma_normalization = None
    ignored_lemmas = None
    lemma_categories = None

    if mode_flags["perform_categorization"]:
        lemma_normalization = load_lemma_normalization(paths["lemma_normalization_json"])
        ignored_lemmas = load_ignored_lemmas(paths["ignored_lemmas_json"])
        lemma_categories = load_lemma_categories(paths["lemma_categories_json"])

    missing_naming_variants, collocation_data, categorized_entries = run_data_collection(
        df=df,
        root=root,
        naming_variants_dict=naming_variants_dict,
        last_verse=last_verse,
        paths=paths,
        missing_naming_variants=missing_naming_variants,
        collocation_data=collocation_data,
        check_naming_variants=mode_flags["check_naming_variants"],
        perform_collocations=mode_flags["perform_collocations"],
        perform_categorization=mode_flags["perform_categorization"],
        lemma_normalization=lemma_normalization,
        ignored_lemmas=ignored_lemmas,
        lemma_categories=lemma_categories,
        categorized_entries=categorized_entries
    )

    save_progress(
        missing_naming_variants=missing_naming_variants,
        last_processed_verse=last_verse,
        paths=paths,
        previous_verse=last_verse,
        previous_naming_variants=previous_naming_variants,
        collocation_data=collocation_data,
        previous_collocations=previous_collocations,
        categorized_entries=categorized_entries,
        previous_categorized_entries=previous_categorized_entries,
        check_naming_variants=mode_flags["check_naming_variants"],
        perform_collocations=mode_flags["perform_collocations"],
        perform_categorization=mode_flags["perform_categorization"]
    )

    return missing_naming_variants, collocation_data, categorized_entries

def finalize_and_prompt(results, data, paths, book_name, config_data):
    """
    Offers the user optional export and analysis steps after data collection is complete.
    """
    print("\n📤 Export results:")
    print(" [1] Naming variants")
    print(" [2] Collocations")
    print(" [3] Categorizations")
    print(" [4] All of the above")
    print(" [0] No export")

    export = ask_user_choice("👉 Please select one or more (e.g., '1,3' or '4'): ",
                             ["0", "1", "2", "3", "4", "1,2", "1,3", "2,3", "1,2,3"])

    if export != "0":
        selected = export.split(",") if export != "4" else ["1", "2", "3"]
        options = {
            "benennungen": "1" in selected,
            "kollokationen": "2" in selected,
            "kategorisierung": "3" in selected
        }
        paths["original_excel"] = data.get("excel_path")
        export_all_data_to_new_excel(book_name, paths, options)

    analyze = ask_user_choice("Do you want to run an analysis now? (y/n): ", ["y", "n"])
    if analyze == "y":
        run_analysis_menu(config_data, paths, data, book_name)
