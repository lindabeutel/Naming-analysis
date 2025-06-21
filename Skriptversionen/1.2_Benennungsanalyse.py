import tkinter as tk
from tkinter import filedialog

import pandas as pd
import json
import os
import re
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
        while True:
            excel_pfad = filedialog.askopenfilename(
                title="Wähle die Excel-Datei mit den Benennungen",
                initialdir=os.getcwd(),
                filetypes=[("Excel-Dateien", "*.xlsx")]
            )
            if not excel_pfad:
                print("⚠️ Keine Excel-Datei ausgewählt.")
                break
            try:
                daten["excel"] = pd.read_excel(excel_pfad)
                daten["excel"].columns = [col.strip().lower() for col in daten["excel"].columns]
                daten["excel"] = pruefe_pflichtspalten(daten["excel"])
                print(f"✅ Excel-Datei geladen: {os.path.basename(excel_pfad)}")
                break  # Erfolgreich geladen, Schleife verlassen
            except PermissionError:
                print("🔒 Bitte schließe die Datei und lade sie danach erneut.")
            except Exception as e:
                print(f"❌ Fehler beim Laden der Excel-Datei: {e}")
                break

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

def pruefe_benennungen_und_kollokationen(daten, benennungsliste, fehlende_benennungen, letzter_bearbeiteter_vers,normalisierung_dict):
    """
    Durchläuft die Excel-Tabelle Vers für Vers und prüft auf fehlende Benennungen oder leere Kollokationen.
    Gibt ggf. erweiterte fehlende_benennungen zurück.
    """

    df = daten["excel"]
    root_tei = daten["xml"]
    gepruefte_benennungen = set()

    for index, row in df.iterrows():
        try:
            vers_nr = int(row["vers"])
        except Exception:
            continue

        if vers_nr < letzter_bearbeiteter_vers:
            continue

        # TEI-Vers holen
        vers_element = root_tei.find(f'.//tei:l[@n="{vers_nr}"]', namespaces={"tei": "http://www.tei-c.org/ns/1.0"})
        if vers_element is None:
            continue

        vers_text = ' '.join([
            seg.text for seg in vers_element.findall('.//tei:seg', namespaces={"tei": "http://www.tei-c.org/ns/1.0"}) if
            seg.text
        ])

        # 🔍 1. Kollokationen prüfen bei vorhandener Benennung
        vorhandene_benennung = str(row.get("eigennennung", "")).strip() or \
                               str(row.get("bezeichnung", "")).strip() or \
                               str(row.get("erzähler", "")).strip()

        if vorhandene_benennung and not str(row.get("kollokationen", "")).strip():
            pattern = r"\b" + re.escape(vorhandene_benennung.lower()) + r"\b"
            if re.search(pattern, vers_text.lower()):
                print("\n-------------------------------------------------------")
                print(f"ℹ️ Die Benennung \"{vorhandene_benennung}\" ist bereits eingetragen.")
                print(f"❗ Die Kollokationen-Spalte ist jedoch leer.")
                koll_text = kontext_kollokation(vers_nr, vorhandene_benennung, root_tei)

                # 📥 Speichere Kollokationen in DataFrame
                df.at[index, "kollokationen"] = koll_text

                # 🧠 Starte Normalisierung/Kategorisierung
                eintrag = {
                    "Vers": vers_nr,
                    "Benannte Figur": row.get("benannte figur", ""),
                    "Eigennennung": row.get("eigennennung", ""),
                    "Bezeichnung": row.get("bezeichnung", ""),
                    "Erzähler": row.get("erzähler", "")
                }
                normalisiere_und_kategorisiere_eintrag(eintrag, normalisierung_dict)

        # 🔍 2. Fehlende Benennungen erkennen (aus benennungsliste)
        for name in benennungsliste:
            pattern = r"\b" + re.escape(name.lower()) + r"\b"
            if not re.search(pattern, vers_text.lower()):
                continue

            key = (vers_nr, name.lower())
            if key in gepruefte_benennungen:
                continue
            gepruefte_benennungen.add(key)

            # Sammle alle Zeilen mit demselben Vers
            zeilen_selben_verses = df[df["vers"] == vers_nr]

            eintraege = []
            for _, z in zeilen_selben_verses.iterrows():
                eintraege.extend([
                    str(z.get("eigennennung", "")).lower(),
                    str(z.get("bezeichnung", "")).lower(),
                    str(z.get("erzähler", "")).lower()
                ])

            alle_tokens = []
            for text in eintraege:
                tokens = re.findall(r'\w+|[^\w\s]', text, re.UNICODE)
                alle_tokens.extend(tokens)

            normalisiertes_wort = normalisierung_dict.get(name.lower(), name.lower())
            if normalisiertes_wort in alle_tokens:
                continue  # ➤ bereits eingetragen

            # ❗ Neue Benennung erkannt
            print("\n-------------------------------------------------------")
            print(f"❗ Eine mögliche fehlende Benennung wurde im Text erkannt!")
            print(f"🔍 Gefundene Benennung: \"{name}\"")
            kontext_benennung(vers_nr, vers_text, name, root_tei)

            bestaetigt = input("Ist dies eine fehlende Benennung? (j/n): ").strip().lower()
            if bestaetigt != "j":
                continue

            erweitern = input("Möchtest du die Benennung erweitern (z. B. mit Adjektiv)? (j/n): ").strip().lower()
            if erweitern == "j":
                name = input("Bitte gib die erweiterte Benennung ein: ").strip()

            print("\nBitte wähle die richtige Kategorie für die Benennung:")
            print("[1] Eigennennung")
            print("[2] Bezeichnung")
            print("[3] Erzähler")
            print("[4] Überspringen")
            kategorie = input("👉 Deine Auswahl: ").strip()

            if kategorie == "4":
                continue

            benannte_figur = input("Gib die \"Benannte Figur\" ein: ").strip()
            nennende_figur = ""
            if kategorie == "2":
                nennende_figur = input("Gib die \"Nennende Figur\" ein: ").strip()

            eintrag = {
                "Vers": vers_nr,
                "Benannte Figur": benannte_figur,
                "Eigennennung": name if kategorie == "1" else "",
                "Bezeichnung": name if kategorie == "2" else "",
                "Erzähler": name if kategorie == "3" else "",
                "Nennende Figur": nennende_figur,
                "Status": "bestätigt"
            }

            koll_frage = input("Möchtest du die Kollokationen ebenfalls erfassen? (j/n): ").strip().lower()
            if koll_frage == "j":
                koll_text = kontext_kollokation(vers_nr, name, root_tei)
                eintrag["Kollokationen"] = koll_text

            # 🧠 Starte Normalisierung/Kategorisierung
            normalisiere_und_kategorisiere_eintrag(eintrag, normalisierung_dict)

            fehlende_benennungen.append(eintrag)

    return fehlende_benennungen


def kontext_benennung(vers_nr: int, vers_text: str, name: str, root_tei: Element) -> None:
    """
    Zeigt den aktuellen Vers mit Benennung sowie Vor- und Nachvers im Kontext.
    """
    tei_ns = {"tei": "http://www.tei-c.org/ns/1.0"}

    def hole_vers(nr):
        zeile = root_tei.find(f'.//tei:l[@n="{nr}"]', namespaces=tei_ns)
        if zeile is not None:
            return ' '.join([seg.text for seg in zeile.findall('.//tei:seg', namespaces=tei_ns) if seg.text])
        return ""

    print("\n📖 Kontext:")

    vorvers = hole_vers(vers_nr - 1)
    if vorvers:
        print(f"{vers_nr - 1}: {vorvers}")

    hervorgehoben = vers_text.replace(name, f"\033[1m\033[93m{name}\033[0m")
    print(f"{vers_nr}: {hervorgehoben}")

    nachvers = hole_vers(vers_nr + 1)
    if nachvers:
        print(f"{vers_nr + 1}: {nachvers}")


def kontext_kollokation(vers_nr: int, name: str, root_tei: Element) -> str:
    """
    Zeigt 13 Verse (±6) rund um den Zielvers mit Hervorhebung der Benennung.
    Gibt den gewählten Textabschnitt zurück.
    """
    tei_ns = {"tei": "http://www.tei-c.org/ns/1.0"}
    kontext = []
    vers_range = list(range(vers_nr - 6, vers_nr + 7))

    for nr in vers_range:
        zeile = root_tei.find(f'.//tei:l[@n="{nr}"]', namespaces=tei_ns)
        if zeile is not None:
            text = ' '.join([seg.text for seg in zeile.findall('.//tei:seg', namespaces=tei_ns) if seg.text])
            if nr == vers_nr:
                text = text.replace(name, f"\033[1m\033[93m{name}\033[0m")
            kontext.append((nr, text))

    print(f"\n📖 Kontextauszug zu Vers {vers_nr}:")
    for i, (nr, line) in enumerate(kontext, 1):
        print(f"{i}. {line}")

    auswahl = input("\n👉 Bitte gib die entsprechenden Nummern ein (z. B. '6-8' oder '7'): ").strip()

    try:
        if "-" in auswahl:
            start, end = map(int, auswahl.split("-"))
            zeilen = [kontext[i - 1][1] for i in range(start, end + 1)]
        else:
            zeilen = [kontext[int(auswahl) - 1][1]]
    except (ValueError, IndexError):
        print("⚠️ Ungültige Eingabe. Keine Kollokation gespeichert.")
        return ""

    return " / ".join(zeilen)


def simple_tokenize(text):
    return re.findall(r'\w+|[^\w\s]', text, re.UNICODE)

def speichere_normalisierungen_dict(normalisierung_dict, path="data/normalisierung.json"):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(normalisierung_dict, f, ensure_ascii=False, indent=2)

def lade_ignore_lemmas(path="ignore_lemmas.json"):
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return set(json.load(f))
    return set()

def speichere_ignore_lemmas(ignore_lemmas, path="ignore_lemmas.json"):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(sorted(ignore_lemmas), f, ensure_ascii=False, indent=2)

def lade_lemma_kategorien(path="lemma_kategorien.json"):
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def speichere_lemma_kategorien(lemma_kategorien, path="lemma_kategorien.json"):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(lemma_kategorien, f, ensure_ascii=False, indent=2)

def normalisiere_und_kategorisiere_eintrag(entry, normalisierung_dict, ignore_lemmas=None, lemma_kategorien=None):
    text = entry.get("Erzähler") or entry.get("Bezeichnung") or entry.get("Eigennennung")
    if not text:
        print("⚠ Kein Text zum Annotieren vorhanden – Eintrag wird übersprungen.\n")
        return None

    print("\n" + "=" * 60)
    print(f"▶ Vers: {entry.get('Vers')}")
    print(f"▶ Benannte Figur: {entry.get('Benannte Figur')}")
    typ = "Erzähler" if entry.get("Erzähler") else ("Bezeichnung" if entry.get("Bezeichnung") else "Eigennennung")
    print(f"▶ Typ: {typ}")
    print(f"\n▶ Originaltext: {text}")

    if ignore_lemmas is None:
        ignore_lemmas = lade_ignore_lemmas()
    if lemma_kategorien is None:
        lemma_kategorien = lade_lemma_kategorien()

    tokens = simple_tokenize(text.lower())

    fehlende = [t for t in tokens if t not in normalisierung_dict]
    if fehlende:
        print("\n▶ Lemmata bitte ergänzen (getrennt durch Komma):")
        user_input = input("> ").strip()
        neue_lemmata = [l.strip() for l in user_input.split(",") if l.strip()]
        if len(neue_lemmata) != len(fehlende):
            print("⚠ Anzahl der eingegebenen Lemmata stimmt nicht mit der Anzahl der unbekannten Tokens überein. Vorgang abgebrochen.\n")
            return None
        for token, lemma in zip(fehlende, neue_lemmata):
            normalisierung_dict[token] = lemma
        speichere_normalisierungen_dict(normalisierung_dict)

    lemmata = [normalisierung_dict.get(t, t) for t in tokens]

    print(f"\n▶ Lemma: {', '.join(lemmata)}\n")

    bezeichnungen = []
    epitheta = []
    history = []

    i = 0
    while i < len(lemmata):
        lemma = lemmata[i]
        if lemma in ignore_lemmas:
            i += 1
            continue

        vorgabe = f"[{lemma_kategorien.get(lemma, '')}]" if lemma in lemma_kategorien else ""
        print(f"{lemma:<12} → {vorgabe} ", end="")
        user_input = input().strip()

        if user_input == "<":
            if i == 0:
                print("↩️  Bereits am Anfang – Rücksprung nicht möglich.")
                continue
            i -= 1
            last_action = history.pop()
            if last_action["type"] == "a":
                bezeichnungen.pop()
            elif last_action["type"] == "e":
                epitheta.pop()
            elif last_action["type"] == "ignore":
                ignore_lemmas.discard(last_action["lemma"])
                speichere_ignore_lemmas(ignore_lemmas)
            elif last_action["type"] == "override":
                del lemma_kategorien[last_action["lemma"]]
                speichere_lemma_kategorien(lemma_kategorien)
            continue

        if user_input == "" and vorgabe:
            if vorgabe == "[a]":
                bezeichnungen.append(lemma)
                history.append({"type": "a", "lemma": lemma})
            elif vorgabe == "[e]":
                epitheta.append(lemma)
                history.append({"type": "e", "lemma": lemma})
            i += 1
            continue

        if user_input == "":
            ignore_lemmas.add(lemma)
            speichere_ignore_lemmas(ignore_lemmas)
            print(f"ℹ️ Lemma „{lemma}“ zur Ignorierliste hinzugefügt.")
            history.append({"type": "ignore", "lemma": lemma})
            i += 1
            continue

        if user_input in ("a", "e"):
            if user_input == "a":
                bezeichnungen.append(lemma)
            else:
                epitheta.append(lemma)
            lemma_kategorien[lemma] = user_input
            speichere_lemma_kategorien(lemma_kategorien)
            history.append({"type": user_input, "lemma": lemma})
            i += 1
            continue

        # Neue Eingabe mit Kategorie
        korrektur = user_input
        kat = ""
        while kat not in ("a", "e"):
            kat = input(f'Definiere die Kategorie für „{korrektur}“ [a/e]: ').strip().lower()

        if kat == "a":
            bezeichnungen.append(korrektur)
        else:
            epitheta.append(korrektur)

        lemma_kategorien[korrektur] = kat
        speichere_lemma_kategorien(lemma_kategorien)

        history.append({
            "type": "override",
            "lemma": korrektur
        })
        i += 1

    if not bezeichnungen and not epitheta:
        print("⚠ Kein Eintrag – bitte prüfe und bestätige erneut.")
        confirm = input("Eintrag wirklich überspringen? [j = ja / n = nein]: ").strip().lower()
        if confirm == "j":
            print("⏭ Eintrag wurde übersprungen.\n")
            return None
        else:
            return normalisiere_und_kategorisiere_eintrag(entry, normalisierung_dict, ignore_lemmas, lemma_kategorien)

    print("✅ Eintrag automatisch gespeichert.\n")
    return {
        **entry,
        "Bezeichnung 1": bezeichnungen[0] if len(bezeichnungen) > 0 else "",
        "Bezeichnung 2": bezeichnungen[1] if len(bezeichnungen) > 1 else "",
        "Bezeichnung 3": bezeichnungen[2] if len(bezeichnungen) > 2 else "",
        "Bezeichnung 4": bezeichnungen[3] if len(bezeichnungen) > 3 else "",
        "Epitheta 1": epitheta[0] if len(epitheta) > 0 else "",
        "Epitheta 2": epitheta[1] if len(epitheta) > 1 else "",
        "Epitheta 3": epitheta[2] if len(epitheta) > 2 else "",
        "Epitheta 4": epitheta[3] if len(epitheta) > 3 else "",
        "Epitheta 5": epitheta[4] if len(epitheta) > 4 else ""
    }

def starte_kategorisierung(buchname, daten, pfade):
    import os
    import json

    # 🔹 Normalisierungen laden
    normalisierung_dict = {}
    normalisierung_path = os.path.join("data", "normalisierung.json")
    if os.path.exists(normalisierung_path):
        with open(normalisierung_path, "r", encoding="utf-8") as f:
            normalisierung_dict = json.load(f)

    # 🔹 Bereits annotierte Einträge laden
    kategorisierte_eintraege = []
    kategorisierungsdatei = os.path.join("data", f"normalisierte_kategorisierung_{buchname}.json")
    if os.path.exists(kategorisierungsdatei):
        with open(kategorisierungsdatei, "r", encoding="utf-8") as f:
            kategorisierte_eintraege = json.load(f)

    # 🔹 Fortschritt laden
    letzter_vers = 0
    if os.path.exists(pfade["progress_json"]):
        with open(pfade["progress_json"], "r", encoding="utf-8") as f:
            letzter_vers = json.load(f).get("letzter_vers", 0)

    # 🔹 Hauptdurchlauf
    for _, row in daten["excel"].iterrows():
        try:
            vers_nr = int(row["vers"])
        except ValueError:
            continue
        if vers_nr < letzter_vers:
            continue

        entry = {
            "Vers": vers_nr,
            "Benannte Figur": row.get("benannte figur", ""),
            "Erzähler": row.get("erzähler", ""),
            "Bezeichnung": row.get("bezeichnung", ""),
            "Eigennennung": row.get("eigennennung", "")
        }

        annotiert = normalisiere_und_kategorisiere_eintrag(entry, normalisierung_dict)
        if annotiert:
            kategorisierte_eintraege.append(annotiert)
            letzter_vers = vers_nr

            with open(pfade["progress_json"], "w", encoding="utf-8") as f:
                json.dump({"letzter_vers": letzter_vers}, f, indent=2, ensure_ascii=False)

            with open(kategorisierungsdatei, "w", encoding="utf-8") as f:
                json.dump(kategorisierte_eintraege, f, indent=2, ensure_ascii=False)


def main():

    os.makedirs("data", exist_ok=True)

    # 🔹 Initialisierung: Buchwahl, Datenpfade, letzter Fortschritt
    buchname, fehlende_benennungen, letzter_bearbeiteter_vers, pfade = initialisiere_projekt()

    # 🔹 Lade globales Benennungsverzeichnis (aus Excel extrahiert oder ergänzt)
    benennungen_dict = lade_oder_erweitere_benennungen_dict()
    benennungsliste = benennungen_dict["Benennungen"].get(buchname, [])

    # 🔹 Merke Zustand vor Verarbeitung
    vorheriger_vers = letzter_bearbeiteter_vers
    vorherige_benennungen = fehlende_benennungen.copy()

    # 🔹 Excel- und TEI-Daten laden
    daten = lade_daten()

    # 🔍 Prüfung auf fehlende Benennungen und leere Kollokationen (optional)
    pruefen = input(
        "Möchtest du die Benennungen und leeren Kollokationen auf Basis der TEI-Datei prüfen? (j/n): ").strip().lower()
    normalisierung_dict = {}

    if pruefen == "j":
        normalisieren = input("Möchtest du die Normalisierung aktivieren? (j/n): ").strip().lower()
        if normalisieren == "j":
            normalisierung_path = os.path.join("data", "normalisierung.json")
            if os.path.exists(normalisierung_path):
                with open(normalisierung_path, "r", encoding="utf-8") as f:
                    normalisierung_dict = json.load(f)
            else:
                print("🆕 Noch keine Normalisierungsdatei vorhanden – sie wird beim ersten Eintrag erstellt.")

        fehlende_benennungen = pruefe_benennungen_und_kollokationen(
            daten,
            benennungsliste,
            fehlende_benennungen,
            letzter_bearbeiteter_vers,
            normalisierung_dict
        )

    # 🔹 Wenn TEI-Prüfung abgelehnt wurde – dann eigene Abfrage
    if pruefen == "n":
        antwort = input(
            "Möchtest du die vorhandenen Benennungen normalisieren und kategorisieren? (j/n): ").strip().lower()
        if antwort == "j":
            starte_kategorisierung(buchname, daten, pfade)

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
