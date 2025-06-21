import tkinter as tk
from tkinter import filedialog

import pandas as pd
import re
import json
import os

import xml.etree.ElementTree as ET

from typing import Optional, Dict, Union
from xml.etree.ElementTree import Element
from openpyxl import load_workbook


DatenTyp = Dict[str, Optional[Union[pd.DataFrame, Element]]]

def initialisiere_projekt():
    """
    Fragt den Benutzer nach dem Buchnamen, legt JSON-Pfade an und lädt ggf. vorhandene Daten.
    Gibt ein Tupel zurück: (buchname, fehlende_benennungen, letzter_bearbeiteter_vers, pfade_dict)
    """

    buchname = input("Welches Buch bearbeiten wir heute? (z. B. Eneasroman): ").strip()

    # Pfade vorbereiten
    os.makedirs("data", exist_ok=True)
    benennungen_json_path = os.path.join("data", f"fehlende_benennungen_{buchname}.json")
    progress_json_path = os.path.join("data", f"progress_{buchname}.json")

    # Fehlende Benennungen laden oder initialisieren
    fehlende_benennungen = None
    if os.path.exists(benennungen_json_path):
        with open(benennungen_json_path, "r", encoding="utf-8") as f:
            fehlende_benennungen = json.load(f)
    if fehlende_benennungen is None:
        fehlende_benennungen = []

    # Fortschritt laden oder auf 0 setzen
    letzter_bearbeiteter_vers = 0
    if os.path.exists(progress_json_path):
        with open(progress_json_path, "r", encoding="utf-8") as f:
            letzter_bearbeiteter_vers = json.load(f).get("letzter_vers", 0)

    pfade = {
        "benennungen_json": benennungen_json_path,
        "progress_json": progress_json_path
    }

    return buchname, fehlende_benennungen, letzter_bearbeiteter_vers, pfade


def lade_daten() -> DatenTyp:
    """Fragt interaktiv nach Excel- und TEI-Dateien, lädt sie bei Zustimmung und gibt sie gesammelt zurück."""
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    daten: DatenTyp = {"excel": None, "xml": None}

    # 1. Excel-Tabelle laden oder neu anlegen
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
                daten["excel"] = pruefe_pflichtspalten(daten["excel"])
                print(f"✅ Excel-Datei geladen: {os.path.basename(excel_pfad)}")
            except Exception as e:
                print(f"❌ Fehler beim Laden der Excel-Datei: {e}")
        else:
            print("⚠️ Keine Excel-Datei ausgewählt.")

    elif antwort_excel == "n":
        neue_excel = input("Möchtest du stattdessen eine neue Excel-Datei anlegen? (j/n): ").strip().lower()
        if neue_excel == "j":
            speicherpfad = filedialog.asksaveasfilename(
                title="Speicherort für neue Excel-Datei wählen",
                defaultextension=".xlsx",
                initialdir=os.getcwd(),
                filetypes=[("Excel-Dateien", "*.xlsx")]
            )
            if speicherpfad:
                try:
                    vorlage_pfad = os.path.join(os.getcwd(), "vorlage_excel.xlsx")
                    wb = load_workbook(vorlage_pfad)
                    wb.save(speicherpfad)
                    daten["excel"] = pd.read_excel(speicherpfad)
                    daten["excel"] = pruefe_pflichtspalten(daten["excel"])
                    print(f"✅ Neue Excel-Datei angelegt: {os.path.basename(speicherpfad)}")
                except Exception as e:
                    print(f"❌ Fehler beim Erstellen der neuen Datei: {e}")
            else:
                print("⚠️ Kein Speicherort ausgewählt.")

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

def pruefe_pflichtspalten(df: pd.DataFrame) -> pd.DataFrame:
    """
    Prüft, ob alle Pflichtspalten im DataFrame vorhanden sind.
    Bei fehlenden Spalten wird nachgefragt, ob sie automatisch ergänzt werden sollen.
    Gibt den ggf. erweiterten DataFrame zurück.
    """
    pflichtspalten = [
        "benannte figur",
        "vers",
        "eigennennung",
        "nennende figur",
        "bezeichnung",
        "erzähler",
        "kollokationen"
    ]

    aktuelle_spalten_lower = [sp.lower() for sp in df.columns]
    fehlende_spalten = [sp for sp in pflichtspalten if sp not in aktuelle_spalten_lower]

    if not fehlende_spalten:
        print("✅ Alle Pflichtspalten sind vorhanden.")
        return df

    print("⚠️ Folgende Pflichtspalten fehlen:")
    for spalte in fehlende_spalten:
        print(f"   – {spalte}")

    for spalte in fehlende_spalten:
        antwort = input(f"Möchtest du die Spalte „{spalte}“ automatisch ergänzen? (j/n): ").strip().lower()
        if antwort == "j":
            df[spalte] = ""
            print(f"➕ Spalte „{spalte}“ ergänzt (leer).")
        else:
            print(f"⚠️ Spalte „{spalte}“ bleibt fehlend.")

    return df


def speichere_fortschritt(fehlende_benennungen, letzter_bearbeiteter_vers, pfade, vorheriger_vers=None, vorherige_benennungen=None):
    """
    Speichert Fortschritt und Benennungen nur, wenn sie sich geändert haben.
    `vorheriger_vers` und `vorherige_benennungen` dienen zum Vergleich.
    """

    # Fortschritt speichern, nur wenn sich etwas geändert hat
    if vorheriger_vers is None or letzter_bearbeiteter_vers != vorheriger_vers:
        with open(pfade["progress_json"], "w", encoding="utf-8") as f:
            json.dump({"letzter_vers": letzter_bearbeiteter_vers}, f, indent=4, ensure_ascii=False)
        print(f"✅ Fortschritt gespeichert (Vers: {letzter_bearbeiteter_vers})")

    # Benennungen speichern, nur wenn sie sich geändert haben
    if vorherige_benennungen is None or fehlende_benennungen != vorherige_benennungen:
        with open(pfade["benennungen_json"], "w", encoding="utf-8") as f:
            json.dump(fehlende_benennungen, f, indent=4, ensure_ascii=False)
        print(f"✅ Fehlende Benennungen gespeichert unter: {pfade['benennungen_json']}")


def lade_oder_erweitere_benennungen_dict():
    """
    Lädt oder erstellt ein zentrales Dictionary mit Figurenbenennungen aus Excel-Dateien.
    Rückgabe: dict mit Struktur {'Enthaltene Bücher': [...], 'Benennungen': {buch: [benennungen, ...]}}
    """

    os.makedirs("data", exist_ok=True)
    dict_path = os.path.join("data", "benennungen_dict.json")

    # Bestehendes Dict laden oder neues anlegen
    if os.path.exists(dict_path):
        with open(dict_path, "r", encoding="utf-8") as f:
            benennungen_dict = json.load(f)
        print(f"📚 Es wurde ein Dictionary gefunden.")
        buecher_liste = benennungen_dict.get("Enthaltene Bücher", [])
        if buecher_liste:
            print(f"👉 Enthaltene Bücher: {', '.join(buecher_liste)}")
        else:
            print("👉 Enthaltene Bücher: [leer]")
        erweitern = input("Möchtest du eine Datei ergänzen? (j/n): ").strip().lower()
    else:
        benennungen_dict = {"Enthaltene Bücher": [], "Benennungen": {}}
        print("❗ Es wurde noch kein Benennungs-Dictionary gefunden.")
        erweitern = "j"

    while erweitern == "j":
        print("📂 Bitte wähle eine Excel-Datei mit Benennungsdaten aus.")
        tk.Tk().withdraw()
        file_path = filedialog.askopenfilename(title="Excel-Datei auswählen", filetypes=[("Excel-Dateien", "*.xlsx")])
        if not file_path:
            print("⚠️ Keine Datei gewählt. Vorgang abgebrochen.")
            break

        buchname = input("Wie lautet der Name des Buchs (z. B. Eneasroman)? ").strip()

        try:
            df = pd.read_excel(file_path)
            relevante_spalten = ["Eigennennung", "Bezeichnung", "Erzähler"]
            benennungen = []

            for spalte in relevante_spalten:
                if spalte in df.columns:
                    benennungen.extend(df[spalte].dropna().tolist())

            # Doppelte entfernen und bereinigen
            benennungen = list(set(str(f).strip().lower() for f in benennungen if str(f).strip()))

        except Exception as e:
            print(f"❌ Fehler beim Einlesen der Datei: {e}")
            break

        benennungen_dict["Enthaltene Bücher"].append(buchname)
        benennungen_dict["Benennungen"][buchname] = benennungen
        print(f"✅ Buch '{buchname}' mit {len(benennungen)} Benennungen hinzugefügt.")

        erweitern = input("Möchtest du eine Datei ergänzen? (j/n): ").strip().lower()

    with open(dict_path, "w", encoding="utf-8") as f:
        json.dump(benennungen_dict, f, indent=4, ensure_ascii=False)
        print(f"💾 Aktuelles Dictionary unter: {dict_path}")

    return benennungen_dict


def main():

    # 🔹 Initialisierung: Buchwahl, Datenpfade, letzter Fortschritt
    buchname, fehlende_benennungen, letzter_bearbeiteter_vers, pfade = initialisiere_projekt()

    # 🔹 Lade globales Benennungsverzeichnis (aus Excel extrahiert oder ergänzt)
    benennungen_dict = lade_oder_erweitere_benennungen_dict()
    benennungsliste = benennungen_dict["Benennungen"].get(buchname, [])

    # 🔹 Merke Zustand vor Verarbeitung, damit keine unnötigen Speicherungen erfolgen
    vorheriger_vers = letzter_bearbeiteter_vers
    vorherige_benennungen = fehlende_benennungen.copy()

    # 🔹 Analyseprozess starten (Platzhalter)
    lade_daten()
    # → hier wird später die Prüfung auf fehlende Benennungen + leere Kollokationen eingefügt

    # 🔹 Fortschritt und Benennungen nur speichern, wenn sich etwas geändert hat
    speichere_fortschritt(
        fehlende_benennungen,
        letzter_bearbeiteter_vers,
        pfade,
        vorheriger_vers=vorheriger_vers,
        vorherige_benennungen=vorherige_benennungen
    )


if __name__ == "__main__":
    main()
