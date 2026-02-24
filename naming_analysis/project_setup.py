"""
project_setup.py

Project bootstrap module for the naming-analysis pipeline.

This module contains the interactive initialization logic that prepares
a corpus-specific working environment at the beginning of a processing session.

Responsibilities:
- Interactive corpus selection (CLI-based).
- Deterministic construction of project-specific directory and file paths.
- Creation of required directory structures if missing.
- Initialization of mandatory corpus-specific JSON session files (with empty/default
  skeleton structures) when they do not yet exist.

Scope:
This module performs filesystem side effects (directory and file creation)
but does not execute analytical logic. It acts as a thin orchestration layer
between user input and the project’s data/IO utilities.

Intended usage:
Called once at the start of a session (typically via run.py).
"""
# Standard library
import json
import os

# Internal project imports
from naming_analysis.io_utils import safe_read_json

def initialize_project() -> tuple[str, int, int, int, dict]:
    """
    Initialize a corpus-specific project session via interactive CLI input.

    This function:
    - prompts the user for a book name,
    - creates required directory structures if they do not exist,
    - constructs all project-specific and global file paths,
    - loads persisted progress information (if available),
    - ensures required JSON resource files exist (via initialize_files).

    Side effects:
        - Reads from stdin (interactive prompt).
        - Creates directories on the filesystem.
        - May create JSON files if they are missing.

    Progress loading behavior:
        If a progress JSON file exists, verse counters are loaded from it.
        Missing keys default to 0.
        If the file does not exist, all counters are initialized to 0.

    Returns:
        tuple[str, int, int, int, dict]:
            book_name:
                Capitalized corpus name as entered by the user.
            naming_variants_last_verse:
                Last processed verse index for naming variants.
            collocations_last_verse:
                Last processed verse index for collocations.
            categorization_last_verse:
                Last processed verse index for categorization.
            paths:
                Dictionary mapping semantic path identifiers to absolute
                or relative filesystem paths used in the session.

    Notes:
        - No input validation is performed on the book name (Beta state).
        - No structural validation of progress JSON content is performed.
    """
    # Interactive corpus selection (no validation; Beta state)
    book_name = input("Which book are we working on today? (e.g., Trojanerkrieg): ").strip()

    # Normalize capitalization (first letter upper-case, rest unchanged)
    book_name = book_name[0].upper() + book_name[1:]

    # Create corpus-specific project directory inside /data
    project_dir = os.path.join("data", book_name)
    os.makedirs(project_dir, exist_ok=True)

    # Ensure shared temporary workspace directory exists
    tmp_dir = os.path.join("data", "tmp")
    os.makedirs(tmp_dir, exist_ok=True)

    # Define corpus-specific configuration and progress file paths
    config_path = os.path.join(project_dir, f"config_{book_name}.json")
    progress_path = os.path.join(project_dir, f"progress_{book_name}.json")

    paths = {
        # ------------------------------------------------------------------
        # Corpus-specific resource files (scoped to selected book)
        # These files live inside: data/<BookName>/
        # ------------------------------------------------------------------

        # JSON: Naming variants container (may contain prefilled/manual entries;
        # legacy filename retained for compatibility)
        "missing_naming_variants_json": os.path.join(project_dir, f"missing_naming_variants_{book_name}.json"),

        # JSON: Stores persistent progress counters (verse-based)
        "progress_json": progress_path,

        # JSON: Stores collocation extraction results
        "collocations_json": os.path.join(project_dir, f"collocations_{book_name}.json"),

        # JSON: Stores categorization data (figure-based naming structure)
        "categorization_json": os.path.join(project_dir, f"categorization_{book_name}.json"),

        # JSON: Configuration file for the selected corpus
        "config_json": config_path,

        # Excel: Human-readable progress tracking / export file
        "progress_excel": os.path.join(project_dir, f"{book_name}_progress.xlsx"),


        # ------------------------------------------------------------------
        # Global resource files (shared across corpora)
        # These files live inside: data/
        # ------------------------------------------------------------------

        "lemma_normalization_json": os.path.join("data", "lemma_normalization.json"),
        "ignored_lemmas_json": os.path.join("data", "ignored_lemmas.json"),
        "lemma_categories_json": os.path.join("data", "lemma_categories.json"),


        # ------------------------------------------------------------------
        # Temporary workspace (session-level)
        # ------------------------------------------------------------------

        # Directory for temporary intermediate artifacts
        "tmp_dir": tmp_dir
    }

    # Initialize verse counters (default: 0)
    naming_variants_last_verse = 0
    collocations_last_verse = 0
    categorization_last_verse = 0

    # Load progress JSON if it exists (missing keys default to 0)
    if os.path.exists(progress_path):
        progress_data = safe_read_json(progress_path, default={})
        naming_variants_last_verse = progress_data.get("naming_variants_last_verse", 0)
        collocations_last_verse = progress_data.get("collocations_last_verse", 0)
        categorization_last_verse = progress_data.get("categorization_last_verse", 0)

    # Ensure required JSON files exist (creates missing skeleton files)
    initialize_files(paths)

    # Return session state (positional API contract)
    return (
        book_name,
        naming_variants_last_verse,
        collocations_last_verse,
        categorization_last_verse,
        paths
    )

def initialize_files(paths: dict) -> None:
    """
    Ensure required corpus-specific JSON files exist.

    For each expected JSON resource file, this function creates a default
    skeleton file if it is missing. Existing files are left untouched.

    Created skeleton structures:

        progress_json:
            {
                "naming_variants_last_verse": 0,
                "collocations_last_verse": 0,
                "categorization_last_verse": 0
            }

        missing_naming_variants_json: []
        collocations_json: []
        categorization_json: []

    Parameters:
        paths (dict):
            Dictionary containing semantic path keys defined in initialize_project().

    Notes:
        - No schema validation is performed.
        - No overwrite of existing files.
        - This function performs filesystem side effects.
    """
    # Create JSON file with default content if it does not yet exist
    def create_if_missing(path: str, content) -> None:
        """
        Create a JSON file with the given content if it does not exist.

        Parameters:
            path (str):
                Target file path.
            content:
                JSON-serializable default content written as file skeleton.

        Notes:
            - Existing files are not modified.
            - No validation of content structure is performed.
        """
        if not os.path.exists(path):
            with open(path, "w", encoding="utf-8") as f:
                json.dump(content, f, indent=4, ensure_ascii=False)

    # Initialize progress tracking JSON (verse counters)
    create_if_missing(paths["progress_json"], {
        "naming_variants_last_verse": 0,
        "collocations_last_verse": 0,
        "categorization_last_verse": 0
    })

    # Initialize empty result containers
    create_if_missing(paths["missing_naming_variants_json"], [])
    create_if_missing(paths["collocations_json"], [])
    create_if_missing(paths["categorization_json"], [])