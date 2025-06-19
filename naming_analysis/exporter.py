"""
exporter.py

This module provides all export-related functionality for the naming analysis project.
It includes routines to:
- insert confirmed naming variants into Excel,
- update collocation fields,
- export categorized lemmata into a new worksheet,
- and generate a final export Excel file for further analysis or archiving.

All formatting is preserved or replicated using openpyxl tools.
"""

import os
import shutil
from copy import copy

import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border
from openpyxl.utils import get_column_letter
from naming_analysis.shared import ask_user_choice

from naming_analysis.io_utils import safe_read_json

def export_all_data_to_new_excel(book_name, paths, options):
    """
    Creates a final Excel file that integrates all collected data:
    - Confirmed naming variants
    - Collocation lines
    - Categorized lemmata

    The function copies the original Excel file and populates or modifies its contents
    based on the current session's JSON exports.

    Parameters:
        book_name (str): Name of the current book or corpus
        paths (dict): Dictionary of file paths used in this session
        options (dict): Dictionary with boolean flags for:
                        'benennungen', 'kollokationen', 'kategorisierung'
    """
    print("🟢 Starting export of all naming variant data...")

    # Support alternate keys for export paths
    paths = {
        **paths,
        "json_benennungen": paths.get("json_benennungen") or paths.get("missing_naming_variants_json"),
        "json_kollokationen": paths.get("json_kollokationen") or paths.get("collocations_json"),
        "json_kategorisierung": paths.get("json_kategorisierung") or paths.get("categorization_json"),
    }

    # 🔧 Ensure output directory exists
    project_dir = os.path.join("data", book_name)
    os.makedirs(project_dir, exist_ok=True)
    target_path = os.path.join(project_dir, f"{book_name}_final.xlsx")

    # Prevent overwriting if source and target are identical – warn and allow user to choose
    if os.path.abspath(paths["original_excel"]) == os.path.abspath(target_path):
        print(f"⚠️ The export target file is the same as the original: {target_path}")
        decision = ask_user_choice("❓ Do you want to overwrite it? (y = overwrite / n = enter new name): ", ["y", "n"])
        if decision != "y":
            new_name = input("📝 Enter a new filename (e.g. 'Eneasroman_final_v2.xlsx'): ").strip()
            target_path = os.path.join(project_dir, new_name)

    # Copy the Excel file
    while True:
        try:
            shutil.copy(paths["original_excel"], target_path)
            break  # Erfolgreich
        except PermissionError:
            print("❌ The Excel file is currently open or locked.")
            print("🔁 Please close the file and try again.")
            retry = ask_user_choice("🔁 Retry export? (y/n): ", ["y", "n"])
            if retry != "y":
                return

    wb = openpyxl.load_workbook(target_path)
    sheet = wb["Gesamt"]

    if options.get("benennungen", False):
        print("📤 Exporting confirmed naming variants...")
        insert_naming_variants(sheet, paths["missing_naming_variants_json"])

    if options.get("kollokationen", False):
        print("📤 Exporting collocations...")
        update_collocations(sheet, paths["collocations_json"])

    if options.get("kategorisierung", False):
        print("📤 Exporting categorized lemmata (this may take a second)...")
        create_categorized_lemmas_sheet(wb, sheet, paths["categorization_json"])

    wb.save(target_path)
    print(f"✅ Export completed: {target_path}")

    # Optional open
    answer = ask_user_choice(
        f"📂 Do you want to open the Excel file '{os.path.basename(target_path)}' now? (y/n):", ["y", "n"]).strip().lower()
    if answer == "y":
        try:
            os.startfile(os.path.abspath(target_path))  # Only works on Windows
        except Exception as e:
            print(f"⚠️ Could not open file: {e}")

def get_format_template(sheet, column_index):
    """
    Extracts the formatting style of the first non-empty cell in the given column.

    Returns a tuple of (font, alignment, border, number_format), or all None if no style found.

    Parameters:
        sheet (Worksheet): The Excel sheet object
        column_index (int): Index of the column to inspect

    Returns:
        tuple: (Font, Alignment, Border, NumberFormat)
    """
    for row in range(2, sheet.max_row + 1):
        cell = sheet.cell(row=row, column=column_index)
        if cell.value:
            if cell.has_style:
                return copy(cell.font), copy(cell.alignment), copy(cell.border), cell.number_format
    return None, None, None, None

def insert_naming_variants(sheet, json_path):
    """
    Appends all confirmed naming variant entries from JSON into the Excel worksheet.

    Each new row is visually highlighted and formatted to match the existing style of the column.

    Parameters:
        sheet (Worksheet): The worksheet named 'Gesamt'
        json_path (str): Path to the naming variant JSON file
    """
    data = safe_read_json(json_path, default=[])

    confirmed_entries = [entry for entry in data if entry.get("Status") == "confirmed"]
    if not confirmed_entries:
        print("ℹ️ No confirmed naming variants to insert.")
        return

    last_line = sheet.max_row + 1
    fill_color = PatternFill(start_color="4BACC6", end_color="4BACC6", fill_type="solid")

    for entry in confirmed_entries:
        new_line = [
            entry.get("Benannte Figur", ""),
            entry.get("Vers", ""),
            entry.get("Eigennennung", ""),
            entry.get("Nennende Figur", ""),
            entry.get("Bezeichnung", ""),
            entry.get("Erzähler", ""),
            entry.get("Kollokation", "")
        ]

        for col_num, value in enumerate(new_line, start=1):
            cell = sheet.cell(row=last_line, column=col_num, value=value)

            font_tpl, alignment_tpl, border_tpl, number_format_tpl = get_format_template(sheet, col_num)
            if font_tpl:
                cell.font = font_tpl
                cell.alignment = alignment_tpl
                cell.border = border_tpl
                cell.number_format = number_format_tpl

            cell.fill = fill_color

        last_line += 1

    print("✅ Naming variants successfully added.")

def update_collocations(sheet, json_path):
    """
    Updates the 'Kollokationen' column in the Excel sheet based on the JSON data.

    Matches are performed by verse number, and formatting is preserved from the template column.

    Parameters:
        sheet (Worksheet): The worksheet to update
        json_path (str): Path to the collocations JSON file
    """
    data = safe_read_json(json_path, default=[])

    header = [cell.value for cell in sheet[1]]
    try:
        verse_col = header.index("Vers") + 1
        collocation_col = header.index("Kollokationen") + 1
    except ValueError:
        print("❌ Columns 'Vers' or 'Kollokationen' not found!")
        return

    verse_to_rows = {}
    for row in range(2, sheet.max_row + 1):
        verse_value = sheet.cell(row=row, column=verse_col).value
        if verse_value is not None:
            verse_to_rows.setdefault(int(verse_value), []).append(row)

    font_tpl, alignment_tpl, border_tpl, number_format_tpl = get_format_template(sheet, collocation_col)

    updated_count = 0
    for entry in data:
        verse = entry["Vers"]
        new_value = entry["Kollokationen"]
        matching_rows = verse_to_rows.get(verse, [])
        for row in matching_rows:
            cell = sheet.cell(row=row, column=collocation_col, value=new_value)
            if font_tpl:
                cell.font = font_tpl
                cell.alignment = alignment_tpl
                cell.border = border_tpl
                cell.number_format = number_format_tpl
            updated_count += 1

    print(f"✅ {updated_count} collocations successfully updated.")

def create_categorized_lemmas_sheet(wb, _, json_path):
    """
    Creates or replaces the worksheet 'lemmatisiert' containing all categorized entries.

    Data is written in structured columns and formatted consistently.
    The sheet is auto-filtered and frozen for easier navigation.

    Parameters:
        wb (Workbook): The Excel workbook object
        _ (unused): Placeholder for the original sheet (not needed)
        json_path (str): Path to the categorized entries JSON file
    """
    # Load JSON data
    annotations = safe_read_json(json_path, default=[])

    # Replace the existing sheet if necessary
    if "lemmatisiert" in wb.sheetnames:
        del wb["lemmatisiert"]
    ws_new = wb.create_sheet("lemmatisiert")

    # Define base styles
    regular_font = Font(name="Times New Roman", size=8, bold=False)
    bold_font = Font(name="Times New Roman", size=8, bold=True)

    default_alignment = Alignment(horizontal="left", vertical="bottom")  # bottom-aligned
    default_border = Border()

    # Define column headers
    headers = [
        "Benannte Figur", "Vers", "Eigennennung", "Nennende Figur", "Bezeichnung", "Erzähler",
        "Bezeichnung 1", "Bezeichnung 2", "Bezeichnung 3", "Bezeichnung 4",
        "Epitheta 1", "Epitheta 2", "Epitheta 3", "Epitheta 4", "Epitheta 5"
    ]

    df = pd.DataFrame(annotations)
    for col in headers:
        if col not in df.columns:
            df[col] = ""
    df = df[headers]

    # Write header row with bold formatting
    for col_idx, header in enumerate(headers, start=1):
        col_letter = get_column_letter(col_idx)
        cell = ws_new.cell(row=1, column=col_idx, value=header)
        ws_new.column_dimensions[col_letter].width = 20
        cell.font = bold_font
        cell.alignment = default_alignment
        cell.border = default_border
        cell.number_format = "General"

    # Write data rows with regular formatting
    for row_idx, row in df.iterrows():
        for col_idx, header in enumerate(headers, start=1):
            cell = ws_new.cell(row=row_idx + 2, column=col_idx, value=row[header])
            cell.font = regular_font
            cell.alignment = default_alignment
            cell.border = default_border
            cell.number_format = "General"

    # Freeze only the first row
    ws_new.freeze_panes = "A2"

    # Enable auto-filter on the header
    ws_new.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"

    # Move the sheet to second position
    wb._sheets.insert(1, wb._sheets.pop(wb._sheets.index(ws_new)))

    print("✅ Worksheet 'lemmatisiert' successfully created.")