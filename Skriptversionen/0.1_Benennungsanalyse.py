import tkinter as tk
from tkinter import filedialog

import pandas as pd
import json
import os
import xml.etree.ElementTree as ET

def lade_daten():
    """Fragt interaktiv nach Excel- und TEI-Dateien, lädt sie bei Zustimmung und gibt sie gesammelt zurück."""
    tk.Tk().withdraw()
    daten = {"excel": None, "xml": None}

    # 1. Excel-Tabelle laden?
    antwort_excel = input("Möchtest du die Tabelle mit den Benennungen laden? (j/n): ").strip().lower()
    if antwort_excel == "j":
        excel_pfad = filedialog.askopenfilename(
            title="Wähle die Excel-Datei mit den Benennungen",
            filetypes=[("Excel-Dateien", "*.xlsx")]
        )
        if excel_pfad:
            try:
                daten["excel"] = pd.read_excel(excel_pfad)
                print(f"✅ Excel-Datei geladen: {os.path.basename(excel_pfad)}")
                print("🔎 Vorschau:")
                print(daten["excel"].head())
            except Exception as e:
                print(f"❌ Fehler beim Laden der Excel-Datei: {e}")
        else:
            print("⚠️ Keine Excel-Datei ausgewählt.")

    # 2. TEI-XML-Datei laden?
    antwort_xml = input("Möchtest du die dazugehörige TEI-Datei laden? (j/n): ").strip().lower()
    if antwort_xml == "j":
        xml_pfad = filedialog.askopenfilename(
            title="Wähle die TEI-XML-Datei",
            filetypes=[("XML-Dateien", "*.xml")]
        )
        if xml_pfad:
            try:
                tree = ET.parse(xml_pfad)
                daten["xml"] = tree.getroot()
                print(f"✅ XML-Datei geladen: {os.path.basename(xml_pfad)}")
                print(f"🔎 XML-Root-Element: <{daten['xml'].tag}>")
            except Exception as e:
                print(f"❌ Fehler beim Laden der XML-Datei: {e}")
        else:
            print("⚠️ Keine XML-Datei ausgewählt.")

    return daten



def main():
    lade_daten()
if __name__ == "__main__":
    main()

