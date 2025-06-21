"""
io_utils.py

Utility functions for reading and writing JSON data with error handling.
Used throughout the project to persist progress, annotations, and configuration.
"""

import json
import time
import os

from naming_analysis.shared import sorted_entries, standardize_verse_number

def safe_write_json(data, path, sort_keys=False, merge=False):
    """
    Safely writes data to a JSON file.
    Optionally merges with existing content and sorts lists of dicts.
    Retries once if the file is temporarily locked.

    Parameters:
        data (any): Data to write.
        path (str): Destination file path.
        sort_keys (bool): Whether to sort top-level keys (only for lists).
        merge (bool): Whether to merge with existing file content if present.
    """
    for attempt in range(2):
        try:
            if merge and os.path.exists(path):
                try:
                    with open(path, "r", encoding="utf-8") as f:
                        existing = json.load(f)
                except (FileNotFoundError, json.JSONDecodeError, PermissionError):
                    existing = [] if isinstance(data, (list, set)) else {}

                if isinstance(data, set):
                    data = list(data)

                if isinstance(data, list) and isinstance(existing, list):
                    if all(isinstance(x, dict) for x in data + existing):
                        seen = set()
                        merged = []
                        for entry in existing + data:
                            key = (
                                entry.get("Vers"),
                                entry.get("Benannte Figur"),
                                entry.get("Eigennennung") or entry.get("Bezeichnung") or entry.get("Erzähler")
                            )
                            if key not in seen:
                                merged.append(standardize_verse_number(entry))
                                seen.add(key)
                        data = merged
                    else:
                        data = list(set(existing).union(set(data)))

                elif isinstance(data, dict) and isinstance(existing, dict):
                    existing.update(data)
                    data = existing

            elif isinstance(data, set):
                data = list(data)

            # Standardize and sort if applicable
            if isinstance(data, list) and all(isinstance(x, dict) and "Vers" in x for x in data):
                data = [standardize_verse_number(entry) for entry in data]
                data = sorted_entries(data)
            elif isinstance(data, dict) and "Vers" in data:
                data = standardize_verse_number(data)

            with open(path, "w", encoding="utf-8") as f:
                json.dump(
                    sorted(data) if sort_keys and isinstance(data, list) else data,
                    f,
                    ensure_ascii=False,
                    indent=2
                )
            return

        except PermissionError as e:
            if attempt == 0:
                print(f"⚠️ Access denied for {path}. Waiting 1 second and retrying...")
                time.sleep(1)
            else:
                print(f"❌ Second attempt failed. File remains locked: {path}")
                raise e

def safe_read_json(path, default=None):
    """
    Safely reads a JSON file and returns its content.
    If the content is a list of dicts with 'Vers', entries are normalized and sorted.

    Parameters:
        path (str): Path to the JSON file.
        default (any): Fallback value in case of read failure.

    Returns:
        any: Parsed and optionally sorted JSON content, or fallback structure.
    """
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)

            if isinstance(data, list) and all(isinstance(x, dict) and "Vers" in x for x in data):
                data = [standardize_verse_number(x) for x in data]
                return sorted_entries(data)

            elif isinstance(data, dict) and "Vers" in data:
                return standardize_verse_number(data)

            return data

    except FileNotFoundError:
        print(f"⚠️ File not found: {path} – using fallback structure.")
        return default if default is not None else {}
    except json.JSONDecodeError:
        print(f"⚠️ Invalid JSON in file {path} – using empty fallback.")
        return default if default is not None else {}
    except PermissionError:
        print(f"❌ Access denied: {path} – read aborted.")
        return default if default is not None else {}

def load_missing_naming_variants(path: str) -> list:
    """
    Loads missing or confirmed naming variants from a JSON file.
    Returns an empty list if the file is unavailable or invalid.

    Parameters:
        path (str): Path to the JSON file.

    Returns:
        list: List of naming variant entries.
    """
    return safe_read_json(path, default=[])
