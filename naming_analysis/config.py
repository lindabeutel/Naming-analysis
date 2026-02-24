"""
config.py

Configuration management module for the naming-analysis pipeline.

This module handles interactive configuration setup and persistence
of session-specific settings. It supports:

- Loading existing configuration data from JSON files,
- Interactive modification or reuse of previous settings,
- Conditional activation of data sources (e.g., Excel, TEI),
- Saving updated configuration states.

Scope:
This module manages configuration state only. It does not execute
analysis logic but determines which data sources and workflow
branches are activated during a session.

Side effects:
- Reads from stdin (interactive prompts).
- Reads from and writes to JSON configuration files.
"""
# Standard library
import os
import shutil
import xml.etree.ElementTree as ET

# Third-party libraries
import pandas as pd

# Internal project imports
from naming_analysis.io_utils import safe_read_json, safe_write_json
from naming_analysis.tei_utils import normalize_tei_text
from naming_analysis.shared import ask_user_choice
from naming_analysis.project_types import DataType
from naming_analysis.validation import check_required_columns
from naming_analysis.loaders import load_data

def save_config(path, config_data):
    """
    Persist configuration data to a JSON file.

    Parameters:
        path (str):
            Target file path.
        config_data (dict):
            JSON-serializable configuration dictionary.

    Behavior:
        - Delegates JSON writing to safe_write_json().
        - Prints a status message to stdout.
        - Catches all exceptions and reports them via CLI output.

    Notes:
        - Exceptions are not re-raised (CLI-oriented behavior).
        - No validation of config_data structure is performed.
        - This function performs filesystem side effects.
    """
    try:
        safe_write_json(config_data, path)
        print(f"Settings saved to: {path}")
    except Exception as e:
        # Broad catch: prevents CLI crash; error is reported but not re-raised
        print(f"Failed to save config: {e}")

def ask_config_interactively(config_path: str) -> tuple[dict, DataType]:
    """
    Interactive configuration loader and session setup for one corpus.

    If a configuration JSON exists, the user can choose to reuse it. When reusing,
    this function attempts to reload previously referenced Excel/TEI resources
    from the stored file paths. If reuse is declined or no config exists, the user
    is guided through a new interactive configuration setup.

    Side effects:
        - Reads from stdin (interactive prompts).
        - Reads from and writes to the configuration JSON (via save_config).
        - May access the filesystem to load Excel/TEI files and copy templates.
        - Assumes execution from the project root directory (required resources
          such as template_excel.xlsx must be available in the working directory).

    Reload behavior (reuse branch):
        - Excel is reloaded if config_data["load_excel"] is true and excel_path exists.
          PermissionError triggers an optional manual reselection (load_data fallback).
        - TEI is reloaded if config_data["load_tei"] is true and tei_path exists.

    Returns:
        tuple[dict, DataType]:
            config_data (dict):
                Updated configuration state (may be reused or newly created).
            data (DataType):
                Container for loaded resources and associated paths.
                Base keys:
                    - "excel"
                    - "excel_path"
                    - "xml"
                Additional keys (e.g., "tei_path") may be added dynamically
                depending on user choices and reload outcomes.
                Keys are populated conditionally and may remain None
                if the corresponding resource was not loaded.

    Notes (Beta state):
        - No structural validation of config_data is performed beyond key presence checks.
        - No strict validation of file contents is performed beyond minimal column checks
          for Excel via check_required_columns().
        - The function is CLI-oriented and reports failures via print statements.
        - TEI loading is optional by design; sessions may run without XML input.
    """
    config_data = {}
    data: DataType = {"excel": None, "excel_path": None, "xml": None}

    # Reuse branch: load existing config JSON and attempt to reload referenced files
    if os.path.exists(config_path):
        reuse = ask_user_choice(
            "A configuration for this book was found. Do you want to reuse the previous settings? (y/n): ",
            ["y", "n"]
        )
        if reuse == "y":
            config_data = safe_read_json(config_path, default={})

            # Excel reload (from persisted path, if enabled in config)
            excel_path = config_data.get("excel_path")
            if config_data.get("load_excel") and excel_path and os.path.exists(excel_path):
                try:
                    df = pd.read_excel(excel_path)
                    df = check_required_columns(df)
                    data["excel"] = df
                    data["excel_path"] = excel_path
                    print(f"Excel file reloaded: {os.path.basename(excel_path)}")
                except PermissionError:
                    # Excel file locked/open: optionally fall back to manual selection
                    print(f"Excel file is currently open or locked: {excel_path}")
                    retry = ask_user_choice("Retry file selection? (y/n): ", ["y", "n"])
                    if retry == "y":
                        partial = load_data(load_excel=True, load_tei=False)
                        if partial.get("excel") is not None:
                            data["excel"] = partial["excel"]
                            data["excel_path"] = partial.get("excel_path")
                            config_data["excel_path"] = partial.get("excel_path")
                except Exception as e:
                    print(f"Failed to reload Excel: {e}")
            elif config_data.get("load_excel"):
                print(f"Excel file not found at saved path: {excel_path}")

            # TEI reload (from persisted path, if enabled in config)
            tei_path = config_data.get("tei_path")
            if config_data.get("load_tei") and tei_path and os.path.exists(tei_path):
                try:
                    tree = ET.parse(tei_path)
                    root_elem = tree.getroot()
                    root_elem = normalize_tei_text(root_elem)
                    data["xml"] = root_elem
                    data["tei_path"] = tei_path
                    print(f"TEI file reloaded: {os.path.basename(tei_path)}")
                except Exception as e:
                    print(f"Failed to reload TEI: {e}")
            elif config_data.get("load_tei"):
                print(f"TEI file not found at saved path: {tei_path}")

            # Early return: reused config + best-effort resource reload
            return config_data, data
        else:
            print("Reusing declined – please define new settings.")
    else:
        print("No existing config found – please define new settings.")

    # Manual setup: Excel (optional)
    config_data["load_excel"] = ask_user_choice(
        "Do you want to load an Excel file with existing naming data? (y/n): ",
        ["y", "n"]
    ) == "y"
    if config_data["load_excel"]:
        partial = load_data(load_excel=True, load_tei=False)

        if partial.get("excel") is not None:
            data["excel"] = partial["excel"]
            data["excel_path"] = partial.get("excel_path")
            config_data["excel_path"] = partial.get("excel_path")
        else:
            print("No Excel file was loaded.")
            config_data["load_excel"] = False

    else:
        # Optional template-based Excel creation (progress workbook)
        print("No Excel file selected.")
        create_new = ask_user_choice("Would you like to create a new Excel file instead? (y/n): ", ["y", "n"])
        if create_new == "y":
            book_name = os.path.basename(config_path).replace("config_", "").replace(".json", "")
            project_dir = os.path.join("data", book_name)
            os.makedirs(project_dir, exist_ok=True)

            new_excel_path = os.path.join(project_dir, f"{book_name}_progress.xlsx")
            # Assumes template_excel.xlsx is located in the current working directory (project root)
            template_path = os.path.join(os.getcwd(), "template_excel.xlsx")

            try:
                shutil.copy(template_path, new_excel_path)
                df = pd.read_excel(new_excel_path)
                df = check_required_columns(df)

                data["excel"] = df
                data["excel_path"] = new_excel_path
                config_data["excel_path"] = new_excel_path
                config_data["load_excel"] = True

                print(f"New Excel file created at: {new_excel_path}")
            except Exception as e:
                print(f"Failed to create new Excel file: {e}")
                config_data["load_excel"] = False
        else:
            config_data["load_excel"] = False

    # Manual setup: TEI (optional)
    config_data["load_tei"] = (
            ask_user_choice("Do you want to load the corresponding TEI file? (y/n): ", ["y", "n"]) == "y"
    )
    if config_data["load_tei"]:
        partial = load_data(load_excel=False, load_tei=True)

        if partial.get("xml") is not None:
            data["xml"] = partial["xml"]
            data["tei_path"] = partial.get("tei_path") or partial.get("xml_path")
            config_data["tei_path"] = data["tei_path"]
        else:
            print("No TEI file was loaded. Disabling TEI-related processing.")
            config_data["load_tei"] = False

    # Select session mode (collect/analyze/export)
    print("What would you like to do today?")
    print("[1] Collect new data")
    print("[2] Analyze existing data")
    print("[3] Export current results")
    mode = ask_user_choice("> ", ["1", "2", "3"])

    if mode == "1":
        config_data["modus"] = "collect"
    elif mode == "2":
        config_data["modus"] = "analyze"
    else:
        config_data["modus"] = "export"

    # For analyze/export, persist config and return (no collection toggles needed)
    if config_data["modus"] in {"analyze", "export"}:
        save_config(config_path, config_data)
        return config_data, data

    # Collection-mode toggles (fine-grained workflow switches)
    config_data["check_naming_variants"] = ask_user_choice("Should namings be checked and added? (y/n): ", ["y", "n"]) == "y"
    config_data["fill_collocations"] = ask_user_choice("Should empty collocations be filled? (y/n): ", ["y", "n"]) == "y"
    config_data["do_categorization"] = ask_user_choice("Should namings be lemmatized and categorized? (y/n): ", ["y", "n"]) == "y"

    save_config(config_path, config_data)

    return config_data, data
