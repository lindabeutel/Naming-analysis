"""
loaders.py

Handles all file loading operations used throughout the project.
Includes Excel, JSON, and TEI data, as well as cached variant dictionaries.
"""

import os
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element
from openpyxl import load_workbook

from naming_analysis.io_utils import safe_read_json, safe_write_json
from naming_analysis.shared import ask_user_choice
from naming_analysis.tei_utils import tei_ns, normalize_tei_text
from naming_analysis.validation import check_required_columns,has_collocations_column
from naming_analysis.project_types import DataType

def load_data(load_excel: bool = False, load_tei: bool = False) -> DataType:
    """
    Interactively loads an Excel file and/or TEI XML file using file dialogs.
    Applies normalization and structural validation where necessary.

    Parameters:
        load_excel (bool): If True, the function will prompt the user to select or create an Excel file.
        load_tei (bool): If True, prompts the user to select a TEI-encoded XML file.

    Returns:
        DataType: A dictionary containing the loaded data and associated file paths.
    """
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    data: DataType = {"excel": None, "excel_path": None, "xml": None}

    # 1. Load or create an Excel file
    if load_excel:
        excel_path = filedialog.askopenfilename(
            title="Select the Excel file with naming data",
            initialdir=os.getcwd(),
            filetypes=[("Excel files", "*.xlsx")]
        )

        if excel_path:
            while True:
                try:
                    df = pd.read_excel(excel_path)
                    df = check_required_columns(df)
                    data["excel"] = df
                    data["excel_path"] = excel_path
                    print(f"✅ Excel file loaded: {os.path.basename(excel_path)}")
                    break
                except PermissionError:
                    print("❌ The Excel file is currently open or locked. Please close it and try again.")
                    retry = ask_user_choice("🔁 Retry file selection? (y/n): ", ["y", "n"])
                    if retry == "y":
                        excel_path = filedialog.askopenfilename(
                            title="Re-select the Excel file",
                            initialdir=os.getcwd(),
                            filetypes=[("Excel files", "*.xlsx")]
                        )
                        if not excel_path:
                            print("⚠️ No file selected – aborting.")
                            break
                    else:
                        break
                except Exception as e:
                    print(f"❌ Error loading Excel file: {e}")
                    break
        else:
            print("⚠️ No Excel file selected.")
            create_new = ask_user_choice("Would you like to create a new Excel file instead? (y/n): ", ["y", "n"])
            if create_new == "y":
                save_path = filedialog.asksaveasfilename(
                    title="Choose save location for the new Excel file",
                    defaultextension=".xlsx",
                    initialdir=os.getcwd(),
                    filetypes=[("Excel files", "*.xlsx")]
                )
                if save_path:
                    try:
                        template_path = os.path.join(os.getcwd(), "template_excel.xlsx")
                        wb = load_workbook(template_path)
                        wb.save(save_path)
                        df = pd.read_excel(save_path)
                        df = check_required_columns(df)
                        data["excel"] = df
                        data["excel_path"] = save_path
                        print(f"✅ New Excel file created: {os.path.basename(save_path)}")
                    except Exception as e:
                        print(f"❌ Error while creating the new file: {e}")
                else:
                    print("⚠️ No save location selected.")

    # 2. Load TEI-XML file
    if load_tei:
        xml_path = filedialog.askopenfilename(
            title="Select the TEI-XML file",
            initialdir=os.getcwd(),
            filetypes=[("XML files", "*.xml")]
        )
        if xml_path:
            try:
                tree = ET.parse(xml_path)
                root_elem = tree.getroot()
                root_elem = normalize_tei_text(root_elem)
                data["xml"] = root_elem
                print(f"✅ XML file loaded: {os.path.basename(xml_path)}")
                data["tei_path"] = xml_path

            except Exception as e:
                print(f"❌ Error loading XML file: {e}")
        else:
            print("⚠️ No XML file selected.")

    return data

def load_collocations_json(file_path):
    """
    Loads collocation entries from a JSON file.

    Parameters:
        file_path (str): Path to the JSON file.

    Returns:
        list: A list of collocation strings, or an empty list if the file is missing or invalid.
    """

    return safe_read_json(file_path, default=[])

def load_lemma_categories(path="data/lemma_categories.json"):
    """
    Loads lemma categories from a JSON file.

    Parameters:
        path (str): Path to the lemma category JSON file.

    Returns:
        dict: A mapping of lemma strings to category labels, or an empty dict on failure.
    """

    return safe_read_json(path, default={})

def load_json_annotations(path):
    """
    Loads annotation data from a JSON file.

    Parameters:
        path (str): Path to the annotation file.

    Returns:
        list: A list of annotation objects, or an empty list if the file cannot be loaded.
    """

    return safe_read_json(path, default=[])

def load_lemma_normalization(path="lemma_normalization.json"):
    """
    Loads lemma normalization rules from a JSON file.

    Parameters:
        path (str): Path to the normalization file.

    Returns:
        dict: A mapping from raw lemma variants to normalized forms.
    """

    return safe_read_json(path, default={})

def load_collocation_sheet(config_data: dict, book_name: str) -> pd.DataFrame | None:
    """
    Loads the collocation sheet 'Gesamt' from the finalized Excel file, or from a fallback path.

    Parameters:
        config_data (dict): The current configuration dictionary.
        book_name (str): Name of the text used to construct the primary file path.

    Returns:
        pd.DataFrame | None: The loaded DataFrame, or None if loading fails.
    """
    primary_path = os.path.join("data", book_name, f"{book_name}_final.xlsx")
    fallback_path = config_data.get("excel_path")

    if os.path.exists(primary_path):
        try:
            df = pd.read_excel(primary_path, sheet_name="Gesamt", engine="openpyxl")
            if not has_collocations_column(df):
                print(f"⚠️ Sheet 'Gesamt' in file '{primary_path}' has no 'Kollokationen' column.")
                return None
            return df
        except PermissionError as e:
            raise e  # lassen wir weiter oben abfangen

    # Nur wenn Datei nicht existiert: Fallback
    if fallback_path and os.path.exists(fallback_path):
        try:
            df = pd.read_excel(fallback_path, sheet_name="Gesamt", engine="openpyxl")
            if not has_collocations_column(df):
                print(f"⚠️ Sheet 'Gesamt' in file '{primary_path}' has no 'Kollokationen' column.")
                return None
            return df
        except Exception as e:
            print(f"⚠️ Could not load fallback file: {e}")
            return None

    print("❌ No valid Excel file found for collocations.")
    return None

def build_fallback_collocation_df_from_tei(root_tei: Element) -> pd.DataFrame:
    """
    Constructs a fallback collocation table from a TEI XML document.
    Uses ±3 lines of verse context around each line to build composite segments.

    Parameters:
        root_tei (Element): The TEI XML root element.

    Returns:
        pd.DataFrame: A DataFrame with columns 'Vers' and 'Kollokationen'.
    """
    context_data = []
    verses = root_tei.findall('.//tei:l', tei_ns)

    for idx, line in enumerate(verses):
        n_attr = line.get("n")
        if not n_attr or not n_attr.isdigit():
            continue
        verse_num = int(n_attr)

        # Kontext: Verse v-3 bis v+3
        segment_texts = []
        for offset in range(-3, 4):
            target_idx = idx + offset
            if 0 <= target_idx < len(verses):
                target_line = verses[target_idx]
                segments = [seg.text for seg in target_line.findall('.//tei:seg', tei_ns) if seg.text]
                segment_texts.append(" ".join(segments))

        full_context = " / ".join(segment_texts)
        context_data.append({"Vers": verse_num, "Kollokationen": full_context})

    return pd.DataFrame(context_data)

def load_or_extend_naming_variants_dict():
    """
    Loads or extends the central naming variants dictionary.
    Allows users to add Excel-based character naming lists interactively.

    Returns:
        dict: The updated naming variants dictionary, including all included books.
    """
    os.makedirs("data", exist_ok=True)
    dict_path = os.path.join("data", "naming_variants_dict.json")

    # Load existing dict or create new one
    if os.path.exists(dict_path):
        naming_variants_dict = safe_read_json(dict_path, default={"Included Books": [], "Namings": {}})
        print(f"📚 A naming dictionary was found.")
        book_list = naming_variants_dict.get("Included Books", [])
        if book_list:
            print(f"👉 Included books: {', '.join(book_list)}")
        else:
            print("👉 Included books: [empty]")
        extend = ask_user_choice("Do you want to add a file? (y/n): ", ["y", "n"])
    else:
        naming_variants_dict = {"Included Books": [], "Namings": {}}
        print("❗ No naming dictionary found.")
        extend = "y"

    while extend == "y":
        print("📂 Please select an Excel file with naming data.")
        tk.Tk().withdraw()
        file_path = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            print("⚠️ No file selected. Operation cancelled.")
            break

        book_name = input("What is the name of the book? (e.g., Eneasroman): ").strip()

        namings = []

        try:
            df = pd.read_excel(file_path)
            relevant_columns = ["Eigennennung", "Bezeichnung", "Erzähler"]
            namings = []

            for column in relevant_columns:
                if column in df.columns:
                    namings.extend(df[column].dropna().tolist())

            # Remove duplicates and normalize
            namings = list(set(str(f).strip().lower() for f in namings if str(f).strip()))


        except PermissionError:
            print("❌ The Excel file is currently open or locked.")
            print("🔁 Please close the file and select it again.")
            file_path = filedialog.askopenfilename(
                title="Re-select the Excel file with namings",
                initialdir=os.getcwd(),
                filetypes=[("Excel files", "*.xlsx")]
            )
            if not file_path:
                print("⚠️ No file selected – aborting.")
                break

        except Exception as e:
            print(f"❌ Error while reading the file: {e}")
            break

        naming_variants_dict["Included Books"].append(book_name)
        naming_variants_dict["Namings"][book_name] = namings
        print(f"✅ Book '{book_name}' added with {len(namings)} naming variants.")

        extend = ask_user_choice("Do you want to add another file? (y/n): ", ["y", "n"])

        safe_write_json(naming_variants_dict, dict_path)
        print(f"💾 Current dictionary saved at: {dict_path}")

    return naming_variants_dict

def load_ignored_lemmas(path="ignored_lemmas.json"):
    """
    Loads the list of ignored lemmas from a JSON file and returns them as a set.

    If the file contains a list, it is converted directly.
    If the file contains a dictionary (legacy format), the keys are used as lemma entries.

    Parameters:
        path (str): Path to the JSON file containing ignored lemmas. Defaults to 'ignored_lemmas.json'.

    Returns:
        set: A set of lemma strings to be excluded from categorization.
    """
    data = safe_read_json(path, default=[])
    return set(data) if isinstance(data, list) else set(data.keys())