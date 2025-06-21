"""
config.py

Handles interactive configuration setup and loading/saving configuration data.
Supports reuse of previous settings and conditional loading of Excel and TEI files.
"""

import os
import pandas as pd
import xml.etree.ElementTree as ET

from naming_analysis.io_utils import safe_read_json, safe_write_json
from naming_analysis.tei_utils import normalize_tei_text
from naming_analysis.shared import ask_user_choice
from naming_analysis.project_types import DataType
from naming_analysis.validation import check_required_columns
from naming_analysis.loaders import load_data

def save_config(path, config_data):
    """
    Saves the configuration dictionary to a JSON file.

    Parameters:
        path (str): Destination path.
        config_data (dict): Configuration dictionary.
    """
    try:
        safe_write_json(config_data, path)
        print(f"✓ Settings saved to: {path}")
    except Exception as e:
        print(f"✗ Failed to save config: {e}")

def ask_config_interactively(config_path: str) -> tuple[dict, DataType]:
    """
    Prompts the user to either reuse an existing configuration or define a new one.
    Handles loading of Excel and TEI files based on user choices and configuration state.

    Parameters:
        config_path (str): Path to the configuration JSON file.

    Returns:
        tuple[dict, DataType]: A tuple containing the loaded configuration and data:
            - config_data (dict): The updated or reused configuration settings.
            - data (DataType): A dictionary with loaded Excel/TEI objects and file paths.
    """
    config_data = {}
    data: DataType = {"excel": None, "excel_path": None, "xml": None}

    if os.path.exists(config_path):
        reuse = ask_user_choice(
            "⚙️ A configuration for this book was found. Do you want to reuse the previous settings? (y/n): ",
            ["y", "n"]
        )
        if reuse == "y":
            config_data = safe_read_json(config_path, default={})

            # Excel reload
            excel_path = config_data.get("excel_path")
            if config_data.get("load_excel") and excel_path and os.path.exists(excel_path):
                try:
                    df = pd.read_excel(excel_path)
                    df = check_required_columns(df)
                    data["excel"] = df
                    data["excel_path"] = excel_path
                    print(f"✅ Excel file reloaded: {os.path.basename(excel_path)}")
                except PermissionError:
                    print(f"❌ Excel file is currently open or locked: {excel_path}")
                    retry = ask_user_choice("🔁 Retry file selection? (y/n): ", ["y", "n"])
                    if retry == "y":
                        partial = load_data(load_excel=True, load_tei=False)
                        if partial.get("excel") is not None:
                            data["excel"] = partial["excel"]
                            data["excel_path"] = partial.get("excel_path")
                            config_data["excel_path"] = partial.get("excel_path")
                except Exception as e:
                    print(f"❌ Failed to reload Excel: {e}")
            elif config_data.get("load_excel"):
                print(f"⚠️ Excel file not found at saved path: {excel_path}")

            # TEI reload
            tei_path = config_data.get("tei_path")
            if config_data.get("load_tei") and tei_path and os.path.exists(tei_path):
                try:
                    tree = ET.parse(tei_path)
                    root_elem = tree.getroot()
                    root_elem = normalize_tei_text(root_elem)
                    data["xml"] = root_elem
                    data["tei_path"] = tei_path
                    print(f"✅ TEI file reloaded: {os.path.basename(tei_path)}")
                except Exception as e:
                    print(f"❌ Failed to reload TEI: {e}")
            elif config_data.get("load_tei"):
                print(f"⚠️ TEI file not found at saved path: {tei_path}")

            return config_data, data
        else:
            print("🛠 Reusing declined – please define new settings.")
    else:
        print("🛠 No existing config found – please define new settings.")

    # Excel: manual load
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
            print("❌ No Excel file was loaded. Disabling Excel-related processing.")
            config_data["load_excel"] = False

    # TEI: manual load
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
            print("❌ No TEI file was loaded. Disabling TEI-related processing.")
            config_data["load_tei"] = False

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

    if config_data["modus"] in {"analyze", "export"}:
        save_config(config_path, config_data)
        return config_data, data

    config_data["check_naming_variants"] = ask_user_choice("Should namings be checked and added? (y/n): ", ["y", "n"]) == "y"
    config_data["fill_collocations"] = ask_user_choice("Should empty collocations be filled? (y/n): ", ["y", "n"]) == "y"
    config_data["do_categorization"] = ask_user_choice("Should namings be lemmatized and categorized? (y/n): ", ["y", "n"]) == "y"

    save_config(config_path, config_data)

    return config_data, data
