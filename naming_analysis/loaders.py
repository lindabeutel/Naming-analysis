"""
loaders.py

Load and validate project inputs (Excel, JSON, TEI/XML) and cached resources.
Includes interactive file selection where applicable.
"""
# Standard library
import os
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element
import tkinter as tk
from tkinter import filedialog

# Third-party libraries
import pandas as pd
from openpyxl import load_workbook

# Project modules
from naming_analysis.io_utils import safe_read_json, safe_write_json
from naming_analysis.shared import ask_user_choice
from naming_analysis.tei_utils import tei_ns, normalize_tei_text
from naming_analysis.validation import (
    check_required_columns,
    has_collocations_column,
)
from naming_analysis.project_types import DataType

# Module entry point
def load_data(load_excel: bool = False, load_tei: bool = False) -> DataType:
    """
    Interactively load an Excel file and/or a TEI-encoded XML file via file dialogs.

    The returned mapping includes:
    - "excel": pandas.DataFrame | None
    - "excel_path": str | None
    - "xml": xml.etree.ElementTree.Element | None
    - "tei_path": str | None

    Parameters:
        load_excel (bool): If True, the function will prompt the user to select or create an Excel file.
        load_tei (bool): If True, prompts the user to select a TEI-encoded XML file.

    Returns:
        DataType: A dictionary containing the loaded data and associated file paths.
    """
    # Initialize a hidden Tk root to enable native file dialogs without showing a GUI window
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    data: DataType = {
        "excel": None,
        "excel_path": None,
        "xml": None,
        "tei_path": None,
    }

    # 1. Load or create an Excel file
    if load_excel:
        excel_path = filedialog.askopenfilename(
            title="Select the Excel file with naming data",
            initialdir=os.getcwd(),
            filetypes=[("Excel files", "*.xlsx")]
        )

        if excel_path:
            # Attempt to load and validate the selected Excel file;
            # retry is offered if the file is locked or inaccessible.
            while True:
                try:
                    df = pd.read_excel(excel_path)
                    df = check_required_columns(df)
                    data["excel"] = df
                    data["excel_path"] = excel_path
                    print(f"Excel file loaded: {os.path.basename(excel_path)}")
                    break
                except PermissionError:
                    # Handle file locks (e.g., file open in Excel) by offering a re-selection
                    print("The Excel file is currently open or locked. Please close it and try again.")
                    retry = ask_user_choice("Retry file selection? (y/n): ", ["y", "n"])
                    if retry == "y":
                        excel_path = filedialog.askopenfilename(
                            title="Re-select the Excel file",
                            initialdir=os.getcwd(),
                            filetypes=[("Excel files", "*.xlsx")]
                        )
                        if not excel_path:
                            print("No file selected – aborting.")
                            break
                    else:
                        break
                except Exception as e:
                    print(f"Error loading Excel file: {e}")
                    break
        else:
            print("No Excel file selected.")
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
                        # Use the project-internal Excel template (located in the project root)
                        # to ensure the expected sheets and column structure.
                        template_path = os.path.join(os.getcwd(), "template_excel.xlsx")
                        wb = load_workbook(template_path)
                        wb.save(save_path)
                        df = pd.read_excel(save_path)
                        df = check_required_columns(df)
                        data["excel"] = df
                        data["excel_path"] = save_path
                        print(f"New Excel file created: {os.path.basename(save_path)}")
                    except Exception as e:
                        print(f"Error while creating the new file: {e}")
                else:
                    print("No save location selected.")

    # 2. Load TEI-XML file
    if load_tei:
        xml_path = filedialog.askopenfilename(
            title="Select the TEI-XML file",
            initialdir=os.getcwd(),
            filetypes=[("XML files", "*.xml")]
        )
        if xml_path:
            try:
                # Parse TEI-XML, normalize textual content, and store the root element
                tree = ET.parse(xml_path)
                root_elem = tree.getroot()
                root_elem = normalize_tei_text(root_elem)
                data["xml"] = root_elem
                print(f"XML file loaded: {os.path.basename(xml_path)}")
                data["tei_path"] = xml_path

            except Exception as e:
                print(f"Error loading XML file: {e}")
        else:
            print("No XML file selected.")

    return data

def load_collocations_json(file_path):
    """
    Load collocation data from a JSON file.

    Parameters:
        file_path (str): Path to the JSON file.

    Returns:
        dict: Collocation data mapping, or an empty mapping if the file is missing or invalid.
    """
    return safe_read_json(file_path, default={})

def load_json_annotations(path):
    """
    Load annotation data from a JSON file.

    Parameters:
        path (str): Path to the annotation file.

    Returns:
        dict: Annotation data mapping, or an empty mapping if the file cannot be loaded.
    """
    return safe_read_json(path, default={})

def load_lemma_categories(path="data/lemma_categories.json"):
    """
    Load lemma-category mappings from a JSON file.

    Parameters
    ----------
    path : str
        Path to the lemma category JSON file.

    Returns
    -------
    dict
        Mapping of lemma strings to category labels.
        Returns an empty dict if the file cannot be read or parsed.

    Notes
    -----
    Delegates file handling and error management to `safe_read_json(...)`.
    No exceptions are raised on read/parse failure.
    """
    return safe_read_json(path, default={})

def load_lemma_normalization(path="lemma_normalization.json"):
    """
    Loads lemma normalization rules from a JSON file.

    Parameters:
        path (str): Path to the normalization file.

    Returns:
        dict: A mapping from raw lemma variants to normalized forms.
    """
    return safe_read_json(path, default={})

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

def load_or_extend_naming_variants_dict() -> dict:
    """
    Load the central naming-variants dictionary and optionally extend it interactively.

    Side effects:
        - Prompts the user to select Excel files and enter book names
        - Writes updates to 'data/naming_variants_dict.json'

    Returns:
        dict: Naming variants dictionary with the following structure:
            {
                "Included Books": list[str],
                "Namings": dict[str, list[str]]
            }
    """
    os.makedirs("data", exist_ok=True)
    dict_path = os.path.join("data", "naming_variants_dict.json")

    # Load existing dict or create new one
    if os.path.exists(dict_path):
        naming_variants_dict = safe_read_json(dict_path, default={"Included Books": [], "Namings": {}})
        print(f"A naming dictionary was found.")
        book_list = naming_variants_dict.get("Included Books", [])
        if book_list:
            print(f"Included books: {', '.join(book_list)}")
        else:
            print("Included books: [empty]")
        extend = ask_user_choice("Do you want to add a file? (y/n): ", ["y", "n"])
    else:
        naming_variants_dict = {"Included Books": [], "Namings": {}}
        print("No naming dictionary found.")
        extend = "y"

    # Initialize a hidden Tk root once for repeated file dialogs
    tk_root = tk.Tk()
    tk_root.withdraw()

    while extend == "y":
        print("Please select an Excel file with naming data.")
        file_path = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            print("No file selected. Operation cancelled.")
            break

        book_name = input("What is the name of the book? (e.g., Eneasroman): ").strip()

        namings = []

        try:
            # Read the selected Excel file and extract naming variants
            df = pd.read_excel(file_path)
            relevant_columns = ["Eigennennung", "Bezeichnung", "Erzähler"]
            namings = []

            for column in relevant_columns:
                if column in df.columns:
                    namings.extend(df[column].dropna().tolist())

            # Normalize naming variants to a canonical form (lowercase, stripped strings)
            # to avoid duplicate entries caused by casing or surrounding whitespace.
            namings = list(set(str(f).strip().lower() for f in namings if str(f).strip()))


        except PermissionError:
            # Handle file locks by prompting the user to re-select the file
            print("The Excel file is currently open or locked.")
            print("Please close the file and select it again.")
            file_path = filedialog.askopenfilename(
                title="Re-select the Excel file with namings",
                initialdir=os.getcwd(),
                filetypes=[("Excel files", "*.xlsx")]
            )
            if not file_path:
                print("No file selected – aborting.")
                break

        except Exception as e:
            # Abort the extension step on unrecoverable read errors
            print(f"Error while reading the file: {e}")
            break

        naming_variants_dict["Included Books"].append(book_name)
        naming_variants_dict["Namings"][book_name] = namings
        print(f"Book '{book_name}' added with {len(namings)} naming variants.")

        extend = ask_user_choice("Do you want to add another file? (y/n): ", ["y", "n"])

        safe_write_json(naming_variants_dict, dict_path)
        print(f"Current dictionary saved at: {dict_path}")

    return naming_variants_dict

def load_collocation_sheet(config_data: dict, book_name: str) -> pd.DataFrame | None:
    """
    Load the collocation sheet ('Gesamt') from an Excel file.

    The function first attempts to load the finalized project file
    located at 'data/{book_name}/{book_name}_final.xlsx'.
    If this file does not exist, a fallback path specified in the
    configuration is used.

    Returns None if the sheet is missing, lacks a 'Kollokationen' column,
    or cannot be read.
    """
    primary_path = os.path.join("data", book_name, f"{book_name}_final.xlsx")
    fallback_path = config_data.get("excel_path")

    if os.path.exists(primary_path):
        try:
            df = pd.read_excel(primary_path, sheet_name="Gesamt", engine="openpyxl")
            if not has_collocations_column(df):
                print(f"Sheet 'Gesamt' in file '{primary_path}' has no 'Kollokationen' column.")
                return None
            return df
        except PermissionError as e:
            # Re-raise permission errors to be handled at a higher level
            raise e

    # Use fallback path only if the primary file does not exist
    if fallback_path and os.path.exists(fallback_path):
        try:
            df = pd.read_excel(fallback_path, sheet_name="Gesamt", engine="openpyxl")
            if not has_collocations_column(df):
                print(f"Sheet 'Gesamt' in file '{fallback_path}' has no 'Kollokationen' column.")
                return None
            return df
        except Exception as e:
            print(f"Could not load fallback file: {e}")
            return None

    print("No valid Excel file found for collocations.")
    return None

def build_fallback_collocation_df_from_tei(root_tei: Element) -> pd.DataFrame:
    """
    Build a fallback collocation table from a TEI/XML document.

    This function is intended as a fallback mechanism when no curated collocation
    sheet is available. The resulting table is derived automatically and should be
    interpreted as heuristic rather than manually validated data.

    For each TEI line element (<l>) with a numeric @n attribute, the function collects
    the textual content of descendant <seg> elements from the current line and its
    surrounding context (index-based window: ±3 <l> elements). The collected segments
    are concatenated using " / " as a separator.

    Returns:
        pd.DataFrame: A DataFrame with columns 'Vers' (int) and 'Kollokationen' (str).
    """
    context_data = []
    verses = root_tei.findall('.//tei:l', tei_ns)

    for idx, line in enumerate(verses):
        n_attr = line.get("n")
        if not n_attr or not n_attr.isdigit():
            continue
        verse_num = int(n_attr)

        # Context window: collect segments from the current line and ±3 surrounding lines
        segment_texts = []
        for offset in range(-3, 4):
            target_idx = idx + offset
            if 0 <= target_idx < len(verses):
                target_line = verses[target_idx]
                segments = [seg.text for seg in target_line.findall('.//tei:seg', tei_ns) if seg.text]
                segment_texts.append(" ".join(segments))

        # Join contextual segments using a visible separator to preserve line boundaries
        full_context = " / ".join(segment_texts)
        context_data.append({"Vers": verse_num, "Kollokationen": full_context})

    return pd.DataFrame(context_data)

def load_naming_sources_with_excel_fallback(paths, data):
    """
    Load tabular naming sources from JSON and Excel with automatic fallback handling.

    - JSON is loaded from 'paths["categorization_json"]' (preferred source).
    - Excel data is taken from an in-memory DataFrame in 'data' if available; otherwise
      it is loaded from 'paths["excel_path"]' or 'data["excel_path"]'. If a sheet named
      'lemmatisiert' exists, it is preferred.

    Returns:
        tuple[pd.DataFrame | None, pd.DataFrame | None]:
            df_json: DataFrame created from the categorization JSON, or None if unavailable.
            df_excel: DataFrame created from Excel input (in-memory or file), or None if unavailable.
    """
    df_json = None
    df_excel = None

    # --- 1) Load JSON (preferred source) ---
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
        print(f"Could not load categorization JSON: {e}")
        df_json = None

    # --- 2) Excel Fallback (automatic) ---
    try:
        # 2.1 try in-memory Excel first
        for k in ("excel", "excel_df", "df_excel"):
            if k in data and isinstance(data[k], pd.DataFrame):
                df_excel = data[k]
                break

        excel_path = paths.get("excel_path") or data.get("excel_path")

        if excel_path and os.path.exists(excel_path):
            try:
                xls = pd.ExcelFile(excel_path)
                sheets_lower = [s.strip().lower() for s in xls.sheet_names]

                if "lemmatisiert" in sheets_lower and df_excel is not None:
                    looks_lemmatized = any(
                        "bezeichnung 1" in str(c).lower() for c in df_excel.columns
                    )
                    if not looks_lemmatized:
                        df_excel = pd.read_excel(
                            excel_path, sheet_name="lemmatisiert", dtype=str
                        )

            except Exception as e:
                print(f"Could not verify or switch to 'lemmatisiert' sheet: {e}")

        # 2.2 If no valid in-memory Excel → load from file
        if df_excel is None:
            excel_path = paths.get("excel_path") or data.get("excel_path")
            if excel_path and os.path.exists(excel_path):
                try:
                    xls = pd.ExcelFile(excel_path)
                    sheets_lower = [s.strip().lower() for s in xls.sheet_names]

                    if "lemmatisiert" in sheets_lower:
                        df_excel = pd.read_excel(
                            excel_path, sheet_name="lemmatisiert", dtype=str
                        )
                        print(
                            f"Excel sheet 'lemmatisiert' loaded from: "
                            f"{excel_path} ({len(df_excel)} rows)."
                        )
                    else:
                        df_excel = pd.read_excel(excel_path, dtype=str)
                        print(
                            f"Excel default sheet loaded from: "
                            f"{excel_path} ({len(df_excel)} rows)."
                        )

                except Exception as e:
                    print(f"Could not load Excel file: {e}")
            else:
                print("Excel path not found or missing.")

    except Exception as e:
        print(f"Could not load Excel fallback: {e}")
        df_excel = None

    # JSON and Excel sources are loaded independently; either return value may be None
    return df_json, df_excel
