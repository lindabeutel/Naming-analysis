"""
project_types.py

Central type definitions for the naming-analysis pipeline.

This module defines shared structural type contracts (e.g., TypedDicts,
aliases, or other typing constructs) that are used across multiple modules.

Scope:
Contains type-level declarations only. No operational logic,
I/O behavior, or data processing should be implemented here.
"""
# Standard library
from xml.etree.ElementTree import Element

# Third-party libraries
import pandas as pd

# Common session data container for loaded resources.
# Keys are string identifiers (e.g., "excel", "excel_path", "xml", "tei_path").
# Values may be DataFrame, XML Element, path string, or None.
DataType = dict[str, pd.DataFrame | Element | str | None]
