"""
io_utils.py

Low-level input/output utilities for the Naming-analysis pipeline.

This module centralizes filesystem-based read/write operations that are
shared across collection, analysis, configuration, and export layers.

Provided functionality:

JSON utilities
--------------
- safe_read_json(...)
- safe_write_json(...)
- load_missing_naming_variants(...)

These helpers provide defensive JSON handling with controlled fallback
behavior (e.g. default return values, retry on PermissionError).
No structural schema validation is performed (BETA state).

CSV utilities
-------------
- write_csv_table(...)

Centralized helper for writing CSV exports in a consistent and
Excel-friendly format. This function standardizes:

- delimiter (default: ';' for DACH/Excel compatibility),
- encoding (default: UTF-8),
- newline handling,
- output directory creation.

All analysis-layer CSV exports should delegate to this function in order
to avoid inconsistent delimiter or encoding settings.

Visualization output utilities
-------------------------------
- export_visualization_output(...)

Centralized helper for handling interactive visualization output.
This function standardizes:

- output mode selection (save / show / both),
- output directory creation (data/<book>/visualization),
- temporary file handling with robust fallback,
- delegation to a provided export function (e.g. apply_global_visual_modebar_export),
- optional browser opening.

All analysis-layer visualizations should delegate their output handling
to this function in order to avoid duplicated CLI and filesystem logic.

Scope (BETA):
-------------
- This module performs direct filesystem side effects.
- It does not perform schema validation or semantic checks.
- It does not contain analytical logic.
- It is designed as a thin I/O abstraction layer to improve FAIR-readability
  and reduce duplication of file-handling code.
"""
# Standard library
import csv
import json
import math
import os
import time
import uuid
import webbrowser

# Shared utilities
from naming_analysis.shared import sorted_entries, standardize_verse_number, ask_user_choice

def safe_read_json(path, default=None):
    """
    Read a JSON file with controlled fallback behavior.

    Behavior:
    - If the file contains a list of dictionaries with key "Vers",
      each entry is normalized via `standardize_verse_number(...)`
      and the list is sorted using `sorted_entries(...)`.
    - If the file contains a single dictionary with key "Vers",
      it is normalized accordingly.
    - Otherwise, the parsed JSON content is returned unchanged.

    On failure (FileNotFoundError, JSONDecodeError, PermissionError),
    the function returns:
    - `default` if provided,
    - otherwise an empty dict `{}`.

    Parameters:
        path (str): Path to the JSON file.
        default: Optional fallback value on read failure.

    Returns:
        Parsed JSON content (possibly normalized/sorted).

        If a read error occurs:
        - returns `default` if `default` is not None,
        - otherwise returns an empty dict `{}`.

        The function does not raise read-related exceptions.
    """

    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)

            # verse-based record lists (normalize "Vers" and apply stable sorting)
            if isinstance(data, list) and all(isinstance(x, dict) and "Vers" in x for x in data):
                data = [standardize_verse_number(x) for x in data]
                return sorted_entries(data)

            # single verse-based record (normalize "Vers")
            elif isinstance(data, dict) and "Vers" in data:
                return standardize_verse_number(data)

            # Default: return parsed JSON unchanged
            return data

    except FileNotFoundError:
        # File missing: return configured fallback structure
        print(f"File not found: {path} – using fallback structure.")
        return default if default is not None else {}
    except json.JSONDecodeError:
        # Invalid JSON: return configured fallback structure
        print(f"Invalid JSON in file {path} – using empty fallback.")
        return default if default is not None else {}
    except PermissionError:
        # Permission denied: return configured fallback structure
        print(f"Access denied: {path} – read aborted.")
        return default if default is not None else {}

def safe_write_json(data, path, sort_keys=False, merge=False):
    """
    Write JSON data to disk with optional merge and verse-based normalization.

    Behavior:
    - If `merge=True` and the destination file exists, existing JSON content is loaded
      and combined with `data`:
        * list + list: merge and de-duplicate (special handling for lists of dicts)
        * dict + dict: `existing.update(data)`
        * other list-like: union via set conversion
    - If `data` is a set, it is converted to a list for JSON serialization.
    - If the final payload is verse-based (list of dicts containing "Vers" or a dict
      containing "Vers"), verse numbers are standardized and entries are sorted.

    Sorting:
    - If `sort_keys=True` and the final payload is a list, `sorted(data)` is written.
      (This sorts list elements, not JSON object keys.)

    Retry:
    - On PermissionError, the function waits 1 second and retries once.

    Parameters:
        data: Data to write (JSON-serializable; sets are converted to lists).
        path (str): Destination file path.
        sort_keys (bool): Sort list elements before writing (only applies if data is a list).
        merge (bool): Merge with existing file content if present.

    Returns:
        None

    Raises:
        PermissionError:
            If writing fails twice due to a locked file.

    Notes:
        The function suppresses file-read errors during merge loading,
        but does not suppress a second PermissionError during writing.
    """
    for attempt in range(2):
        try:
            if merge and os.path.exists(path):
                try:
                    with open(path, "r", encoding="utf-8") as f:
                        existing = json.load(f)
                except (FileNotFoundError, json.JSONDecodeError, PermissionError):
                    # If existing file cannot be read, initialize empty structure
                    existing = [] if isinstance(data, (list, set)) else {}

                # Ensure JSON-serializable type
                if isinstance(data, set):
                    data = list(data)

                # Case 1: list + list → merge and de-duplicate
                if isinstance(data, list) and isinstance(existing, list):
                    # Specialized handling for lists of dict entries
                    if all(isinstance(x, dict) for x in data + existing):
                        seen = set()
                        merged = []
                        for entry in existing + data:
                            # Deduplication key based on core identifying fields
                            # NaN values are normalized to None for stable comparison
                            key = tuple(
                                None if (isinstance(entry.get(k), float) and math.isnan(entry.get(k)))
                                else entry.get(k)
                                for k in ("Vers", "Benannte Figur", "Bezeichnung", "Erzähler", "Eigennennung")
                            )
                            if key not in seen:
                                merged.append(standardize_verse_number(entry))
                                seen.add(key)
                        data = merged
                    else:
                        # Generic list merge via set union
                        data = list(set(existing).union(set(data)))

                elif isinstance(data, dict) and isinstance(existing, dict):
                    existing.update(data)
                    data = existing

            # Non-merge path: ensure JSON-serializable type
            elif isinstance(data, set):
                data = list(data)

            # Standardize and sort verse-based structures before writing
            if isinstance(data, list) and all(isinstance(x, dict) and "Vers" in x for x in data):
                # Normalize verse numbers in each entry
                data = [standardize_verse_number(entry) for entry in data]
                # Apply project-defined stable sorting
                data = sorted_entries(data)
            elif isinstance(data, dict) and "Vers" in data:
                # Normalize single verse-based record
                data = standardize_verse_number(data)

            # Write JSON to disk (UTF-8, pretty-printed)
            with open(path, "w", encoding="utf-8") as f:
                json.dump(
                    # Optional list sorting prior to serialization
                    sorted(data) if sort_keys and isinstance(data, list) else data,
                    f,
                    ensure_ascii=False,
                    indent=2
                )

            # Successful write → exit function
            return

        except PermissionError as e:
            # Retry once if file is temporarily locked
            if attempt == 0:
                print(f"Access denied for {path}. Waiting 1 second and retrying...")
                time.sleep(1)
            else:
                # Second failure → propagate exception
                print(f"Second attempt failed. File remains locked: {path}")
                raise e

def load_missing_naming_variants(path: str) -> list:
    """
    Load naming-variant records from JSON with list fallback.

    Thin wrapper around `safe_read_json(...)` that enforces a list return type
    by providing `default=[]`.

    On read failure (missing file, invalid JSON, permission error),
    an empty list is returned.

    Parameters:
        path (str): Path to the JSON file.

    Returns:
        list: List of naming-variant entries (possibly empty).
    """
    return safe_read_json(path, default=[])

def write_csv_table(
    output_path,
    header,
    rows,
    *,
    delimiter=";",
    encoding="utf-8",
) -> None:
    """
    Write a CSV table (header + rows) with consistent defaults.

    This helper centralizes CSV output settings (delimiter, encoding, newline)
    to keep analysis exports consistent and Excel-friendly.

    Parameters:
        output_path (str | os.PathLike): Destination CSV path.
        header (list[str] | tuple[str, ...]): Column names written as the first row.
        rows (iterable): Data rows (iterables of values) written after the header.
        delimiter (str): CSV delimiter (default: ';' for DACH/Excel).
        encoding (str): File encoding (default: 'utf-8').

    Returns:
        None
    """
    # Ensure output directory exists (no-op if output_path has no directory part).
    out_dir = os.path.dirname(str(output_path))
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    with open(output_path, "w", encoding=encoding, newline="") as f:
        writer = csv.writer(f, delimiter=delimiter)
        writer.writerow(list(header))
        writer.writerows(rows)

    with open(output_path, "w", encoding=encoding, newline="") as f:
        writer = csv.writer(f, delimiter=delimiter)
        writer.writerow(list(header))
        writer.writerows(rows)


def export_visualization_output(
    fig,
    *,
    paths: dict,
    book_name: str,
    filename: str,
    export_func,
    filename_stub: str | None = None,
    output_dir: str | None = None,
    tmp_fallback_dir: str | None = None,
) -> None:
    """
    Export and/or display an interactive visualization HTML file.

    This helper centralizes the repeated output workflow used by analysis
    visualizations:
    - prompt for output mode (save / show / both),
    - ensure directories exist,
    - write HTML via the provided `export_func`,
    - optionally open the output in a browser.

    The function performs filesystem side effects and user interaction (CLI).
    No schema validation is performed (BETA state).
    """
    if output_dir is None:
        output_dir = os.path.join("data", book_name, "visualization")

    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, filename)

    if tmp_fallback_dir is None:
        tmp_fallback_dir = os.path.join("data", book_name, "tmp")

    tmp_dir = paths.get("tmp_dir", tmp_fallback_dir)
    os.makedirs(tmp_dir, exist_ok=True)

    print("\nHow should the output be handled?")
    print("[1] Save as HTML file")
    print("[2] Show plot in browser")
    print("[3] Both")
    output_mode = ask_user_choice("> ", ["1", "2", "3"])

    if output_mode == "1":
        export_func(fig, output_path, filename_stub=filename_stub)
        print("\nVisualization completed.")
        print(f"File saved at:\n{output_path}")
        return

    if output_mode == "2":
        tmp_filename = f"viz_{uuid.uuid4().hex[:8]}.html"
        tmp_path = os.path.join(tmp_dir, tmp_filename)
        export_func(fig, tmp_path, filename_stub=filename_stub)
        webbrowser.open_new_tab(f"file://{os.path.abspath(tmp_path)}")
        print("The plot has been opened in your browser.")
        print(f"Temporary file created at: {tmp_path}")
        return

    # output_mode == "3"
    export_func(fig, output_path, filename_stub=filename_stub)
    print("\nVisualization completed.")
    print(f"File saved at:\n{output_path}")
    webbrowser.open_new_tab(f"file://{os.path.abspath(output_path)}")
    print("The plot has been opened in your browser.")
