"""
validation.py

Validation helpers for input tables (primarily Excel-derived DataFrames).

Scope (BETA):
- Performs *structural* checks only (presence of required columns).
- Does not validate column *content* (types, formats, emptiness) in the BETA stage.
- `check_required_columns()` may prompt interactively and may *mutate* the passed DataFrame
  by adding missing columns (create-if-missing semantics).

These helpers are intended to fail early (or guide the user) before downstream processing.
"""
# Third-party libraries
import pandas as pd

def check_required_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Ensure that all required naming-analysis columns are present in the given DataFrame.

    Behavior (BETA):
    - Required column names are defined in lowercase and compared against the
      lowercased existing headers (case-insensitive structural check).
    - If one or more required columns are missing, the user is prompted
      interactively for each column.
    - If confirmed, the missing column is created with empty-string values.
    - The input DataFrame may be mutated in place and is returned (same object).

    Scope:
    - Only checks structural presence of columns.
    - Does not validate column content, types, or value completeness.

    Parameters:
        df (pd.DataFrame): The input table to validate.

    Returns:
        pd.DataFrame: The original DataFrame, possibly extended with newly
        created empty columns.
    """
    # Canonical list of required headers (Excel schema, canonical capitalization).
    # These correspond to the expected column names used throughout the pipeline.
    required_columns = [
        "Benannte Figur",
        "Vers",
        "Eigennennung",
        "Nennende Figur",
        "Bezeichnung",
        "Erzähler",
        "Kollokationen"
    ]

    # Normalize both sides for comparison (case-insensitive structural check)
    required_columns_lower = [col.lower() for col in required_columns]
    current_columns_lower = [col.lower() for col in df.columns]

    # Keep canonical names for user prompts, but compare via lowercase lists
    missing_columns = [
        required_columns[i]
        for i, req_lower in enumerate(required_columns_lower)
        if req_lower not in current_columns_lower
    ]

    # Early return if schema is complete.
    if not missing_columns:
        print("All required columns are present.")
        return df

    # Inform user about structural schema gaps (interactive CLI feedback).
    print("The following required columns are missing:")
    for col in missing_columns:
        print(f"   – {col}")

    # For each missing column, ask whether it should be created automatically.
    # Creation semantics: column is added with empty-string values.
    # This mutates the original DataFrame in place.
    for col in missing_columns:
        answer = input(f"Do you want to add the column \"{col}\" automatically? (y/n): ").strip().lower()
        if answer == "y":
            # Add new column with empty-string default values.
            # No type enforcement or content validation is performed (BETA stage).
            df[col] = ""
            print(f"Column \"{col}\" added (empty).")
        else:
            # Column intentionally remains absent.
            # Downstream functions may fail if this schema is required.
            print(f"Column \"{col}\" remains missing.")

    # Return the (possibly mutated) DataFrame.
    return df

def has_collocations_column(df: pd.DataFrame) -> bool:
    """
    Check whether the DataFrame contains a column named exactly "Kollokationen".

    Behavior (BETA):
    - The check is case-sensitive and matches the column header exactly as written.
    - No normalization (e.g. lowercasing) is performed.
    - Only structural presence is checked; column content is not inspected.

    Parameters:
        df (pd.DataFrame): The DataFrame to inspect.

    Returns:
        bool: True if a column named exactly "Kollokationen" exists,
              False otherwise.
    """
    return "Kollokationen" in df.columns
