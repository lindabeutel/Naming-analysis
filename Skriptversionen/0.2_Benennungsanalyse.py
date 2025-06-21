import tkinter as tk
from tkinter import filedialog

import pandas as pd
import os
import xml.etree.ElementTree as ET

from typing import Optional, Dict, Union
from xml.etree.ElementTree import Element

DatenTyp = Dict[str, Optional[Union[pd.DataFrame, Element]]]


def lade_daten() -> DatenTyp:
    """Fragt interaktiv nach Excel- und TEI-Dateien, lädt sie bei Zustimmung und gibt sie gesammelt zurück."""
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    daten: DatenTyp = {"excel": None, "xml": None}

    # 1. Excel-Tabelle laden?
    antwort_excel = input("Möchtest du eine Excel-Tabelle mit bereits erhobenen Benennungen laden? (j/n): ").strip().lower()
    if antwort_excel == "j":
        excel_pfad = filedialog.askopenfilename(
            title="Wähle die Excel-Datei mit den Benennungen",
            initialdir=os.getcwd(),
            filetypes=[("Excel-Dateien", "*.xlsx")]
        )
        if excel_pfad:
            try:
                daten["excel"] = pd.read_excel(excel_pfad)
                print(f"✅ Excel-Datei geladen: {os.path.basename(excel_pfad)}")
            except Exception as e:
                print(f"❌ Fehler beim Laden der Excel-Datei: {e}")
        else:
            print("⚠️ Keine Excel-Datei ausgewählt.")

    # 2. TEI-XML-Datei laden?
    antwort_xml = input("Möchtest du die dazugehörige TEI-Datei laden? (j/n): ").strip().lower()
    if antwort_xml == "j":
        xml_pfad = filedialog.askopenfilename(
            title="Wähle die TEI-XML-Datei",
            initialdir=os.getcwd(),
            filetypes=[("XML-Dateien", "*.xml")]
        )
        if xml_pfad:
            try:
                tree = ET.parse(xml_pfad)
                daten["xml"] = tree.getroot()
                print(f"✅ XML-Datei geladen: {os.path.basename(xml_pfad)}")
            except Exception as e:
                print(f"❌ Fehler beim Laden der XML-Datei: {e}")
        else:
            print("⚠️ Keine XML-Datei ausgewählt.")

    return daten


def main():
    lade_daten()
if __name__ == "__main__":
    main()

