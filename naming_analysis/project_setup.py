"""
project_setup.py

This module provides project initialization logic for the naming analysis pipeline.
It defines functions to:
- interactively select the target corpus,
- create all necessary folder structures and file paths,
- and initialize required JSON files if they do not exist.

Used at the beginning of each processing session.
"""

import os
import json

from naming_analysis.io_utils import safe_read_json


def initialize_project():
    """
    Interactively selects a book to work on, creates directory structure and
    initializes project-specific file paths.

    Also loads verse-related progress information if available and the triggers
    creation of any missing JSON files.

    Returns:
        tuple:
            book_name (str): Capitalized name of the selected book
            naming_variants_last_verse (int): Last processed verse for naming variants
            collocations_last_verse (int): Last processed verse for collocations
            categorization_last_verse (int): Last processed verse for categorization
            paths (dict): Dictionary of all file paths used in this session
    """
    book_name = input("Which book are we working on today? (e.g., Rolandslied): ").strip()
    book_name = book_name[0].upper() + book_name[1:]

    project_dir = os.path.join("data", book_name)
    os.makedirs(project_dir, exist_ok=True)
    os.makedirs("data", exist_ok=True)  # Für globale Dateien, falls nicht vorhanden

    config_path = os.path.join(project_dir, f"config_{book_name}.json")
    progress_path = os.path.join(project_dir, f"progress_{book_name}.json")

    paths = {
        # Projektbezogene Dateien
        "missing_naming_variants_json": os.path.join(project_dir, f"missing_naming_variants_{book_name}.json"),
        "progress_json": progress_path,
        "collocations_json": os.path.join(project_dir, f"collocations_{book_name}.json"),
        "categorization_json": os.path.join(project_dir, f"categorization_{book_name}.json"),
        "config_json": config_path,

        # Globale Dateien
        "lemma_normalization_json": os.path.join("data", "lemma_normalization.json"),
        "ignored_lemmas_json": os.path.join("data", "ignored_lemmas.json"),
        "lemma_categories_json": os.path.join("data", "lemma_categories.json")
    }

    # Fortschritt laden oder initialisieren
    naming_variants_last_verse = 0
    collocations_last_verse = 0
    categorization_last_verse = 0

    if os.path.exists(progress_path):
        progress_data = safe_read_json(progress_path, default={})
        naming_variants_last_verse = progress_data.get("naming_variants_last_verse", 0)
        collocations_last_verse = progress_data.get("collocations_last_verse", 0)
        categorization_last_verse = progress_data.get("categorization_last_verse", 0)

    # Fehlende Dateien anlegen
    initialize_files(paths)

    return (
        book_name,
        naming_variants_last_verse,
        collocations_last_verse,
        categorization_last_verse,
        paths
    )

def initialize_files(paths):
    """
    Creates all required project-specific JSON files if they do not already exist.

    This includes:
    - progress file (verse tracking),
    - missing naming variants,
    - collocations,
    - categorization results.

    Parameters:
        paths (dict): Dictionary containing target file paths.
    """
    def create_if_missing(path, content):
        if not os.path.exists(path):
            with open(path, "w", encoding="utf-8") as f:
                json.dump(content, f, indent=4, ensure_ascii=False)

    create_if_missing(paths["progress_json"], {
        "naming_variants_last_verse": 0,
        "collocations_last_verse": 0,
        "categorization_last_verse": 0
    })

    create_if_missing(paths["missing_naming_variants_json"], [])
    create_if_missing(paths["collocations_json"], [])
    create_if_missing(paths["categorization_json"], [])