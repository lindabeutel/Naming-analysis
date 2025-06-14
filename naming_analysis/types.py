"""
types.py

Shared type definitions used throughout the naming_analysis package.
"""

from typing import Union
import pandas as pd
from xml.etree.ElementTree import Element

# Common data container used for loaded files (Excel, TEI)
DataType = dict[str, Union[pd.DataFrame, Element, str, None]]
