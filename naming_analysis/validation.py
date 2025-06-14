"""
validation.py

Contains validation functions for input data such as Excel tables.
Used to ensure that all required columns are present before further processing.
"""

import pandas as pd

def check_required_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
     Validates whether all required columns for naming analysis are present in the given DataFrame.
     If any are missing, the user is prompted to add them interactively.

     Parameters:
         df (pd.DataFrame): The input DataFrame to validate.

     Returns:
         pd.DataFrame: The original or extended DataFrame with missing columns optionally added.
     """
    required_columns = [
        "benannte figur",
        "vers",
        "eigennennung",
        "nennende figur",
        "bezeichnung",
        "erzähler",
        "kollokationen"
    ]

    current_columns_lower = [col.lower() for col in df.columns]
    missing_columns = [col for col in required_columns if col not in current_columns_lower]

    if not missing_columns:
        print("✅ All required columns are present.")
        return df

    print("⚠️ The following required columns are missing:")
    for col in missing_columns:
        print(f"   – {col}")

    for col in missing_columns:
        answer = input(f"Do you want to add the column \"{col}\" automatically? (y/n): ").strip().lower()
        if answer == "y":
            df[col] = ""
            print(f"➕ Column \"{col}\" added (empty).")
        else:
            print(f"⚠️ Column \"{col}\" remains missing.")

    return df

def has_collocations_column(df: pd.DataFrame) -> bool:
    """
    Checks whether the DataFrame contains a 'Kollokationen' column.

    Parameters:
        df (pd.DataFrame): The DataFrame to check.

    Returns:
        bool: True if the column exists, False otherwise.
    """
    return "Kollokationen" in df.columns
