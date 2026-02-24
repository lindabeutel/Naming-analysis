"""
controller.py

Orchestrates the naming-analysis workflow.

Responsibilities:
- Initialize a project session (book selection, paths, config, and input data)
- Execute the data-collection workflow (naming variants, collocations, categorization)
- Persist intermediate progress
- Optionally trigger export and/or analysis menus

This module is called from the main entry point (run.py).
"""
import sys

# Project setup / configuration
from naming_analysis.project_setup import initialize_project
from naming_analysis.config import ask_config_interactively

# Data loading
from naming_analysis.loaders import (
    load_or_extend_naming_variants_dict,
    load_lemma_normalization,
    load_ignored_lemmas,
    load_lemma_categories,
    load_collocations_json,
    load_json_annotations,
)

# Data collection and persistence
from naming_analysis.collection import run_data_collection
from naming_analysis.savers import save_progress

# Export and analysis
from naming_analysis.exporter import export_all_data_to_new_excel
from naming_analysis.analysis import run_analysis_menu

# Shared utilities
from naming_analysis.shared import ask_user_choice
from naming_analysis.io_utils import load_missing_naming_variants

def setup_project_session() -> tuple[str, dict, dict, dict, int, dict, dict]:
    """
    Initialize a project session: select a book, load paths, read configuration, and load input data.

    Notes:
        Depending on the configured mode ("analyze" / "export"), this function may run the
        corresponding menu and terminate the program early.

    Returns:
        tuple:
            book_name (str)
            config_data (dict)
            data (dict)
            paths (dict)
            last_verse (int): last processed verse for the selected workflow
            mode_flags (dict): boolean flags controlling the selected operations
            naming_variants_dict (dict): reference dictionary for naming-variant lookup/extension
    """
    book_name, naming_variants_last_verse, collocations_last_verse, categorization_last_verse, paths = initialize_project()
    naming_variants_dict = load_or_extend_naming_variants_dict()

    config_data, data = ask_config_interactively(paths["config_json"])
    paths["original_excel"] = data.get("excel_path")

    mode_flags = {
        "check_naming_variants": config_data.get("check_naming_variants", False),
        "perform_collocations": config_data.get("fill_collocations", False),
        "perform_categorization": config_data.get("do_categorization", False)
    }

    if config_data.get("modus") == "analyze":
        run_analysis_menu(config_data, paths, data, book_name)
        sys.exit(0)

    elif config_data.get("modus") == "export":
        options = {
            "benennungen": config_data.get("check_naming_variants", True),
            "kollokationen": config_data.get("fill_collocations", True),
            "kategorisierung": config_data.get("do_categorization", True)
        }
        export_all_data_to_new_excel(book_name, paths, options)

        analyze = ask_user_choice("Do you want to run an analysis now? (y/n): ", ["y", "n"])
        if analyze == "y":
            run_analysis_menu(config_data, paths, data, book_name)

        sys.exit(0)

    if mode_flags["check_naming_variants"]:
        last_verse = naming_variants_last_verse
    elif mode_flags["perform_collocations"]:
        last_verse = collocations_last_verse
    elif mode_flags["perform_categorization"]:
        last_verse = categorization_last_verse
    else:
        last_verse = 0

    return book_name, config_data, data, paths, last_verse, mode_flags, naming_variants_dict

def run_data_workflow(
    data,
    paths,
    last_verse,
    mode_flags,
    naming_variants_dict
) -> tuple[dict, dict, dict]:
    """
    Execute the data-collection workflow for the selected mode(s).

    Depending on the active flags, this function updates naming variants,
    collocations, and/or categorization data and persists intermediate progress.

    Returns:
        tuple:
            missing_naming_variants (dict)
            collocation_data (dict)
            categorized_entries (dict)
    """
    # Input handles: tabular data (Excel) and parsed TEI/XML root
    df = data.get("excel")
    root = data.get("xml")

    # Core inputs, mode flags, and optional resources for the data-collection step
    missing_naming_variants = load_missing_naming_variants(paths["missing_naming_variants_json"])
    collocation_data = load_collocations_json(paths["collocations_json"])
    categorized_entries = load_json_annotations(paths["categorization_json"])

    # Create snapshots of the current state before running the data-collection step
    previous_naming_variants = missing_naming_variants.copy()
    previous_collocations = collocation_data.copy()
    previous_categorized_entries = categorized_entries.copy()

    # Optional resources, only required for categorization mode
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
        # Mode flags control which sub-workflows are executed in this run
        check_naming_variants=mode_flags["check_naming_variants"],
        perform_collocations=mode_flags["perform_collocations"],
        perform_categorization=mode_flags["perform_categorization"],
        lemma_normalization=lemma_normalization,
        ignored_lemmas=ignored_lemmas,
        lemma_categories=lemma_categories,
        categorized_entries=categorized_entries
    )

    # Persist updated data and compare it against the pre-run snapshots
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

def finalize_and_prompt(results, data, paths, book_name, config_data) -> None:
    """
    Prompt the user for optional export and analysis steps after data collection.

    Parameters:
        results:
            Placeholder for the return value of `run_data_workflow(...)`.
            This parameter is intentionally unused here and exists to keep the
            call signature compatible with the main execution flow in `run.py`,
            where the workflow result is assigned but not required for prompting.
        data (dict): Loaded input data (Excel, XML).
        paths (dict): Project paths.
        book_name (str): Name of the current text.
        config_data (dict): Active configuration.
    """
    print("\nExport results:")
    print(" [1] Naming variants")
    print(" [2] Collocations")
    print(" [3] Categorizations")
    print(" [4] All of the above")
    print(" [0] No export")

    export = ask_user_choice("Please select one or more (e.g., '1,3' or '4'): ",
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
