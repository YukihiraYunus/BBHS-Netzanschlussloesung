def run():
    import pandas as pd
    import json

    # === 1. MAPPING ===
    original_mapping = {
        "Neuer Netzanschluss": ("vorgang", "vorgang", "optionen", "Ich möchte einen neuen Netzanschluss"),
        "Zeitlich befristeter Netzanschluss": ("vorgang", "vorgang", "optionen", "Zeitlich befristeter Netzanschluss"),
        "Änderung eines bestehenden Netzanschlusses: Umlegung": ("vorgang", "vorgang", "optionen", "Änderung eines bestehenden Netzanschlusses: Umlegung"),
        "Änderung eines bestehenden Netzanschlusses: Netzanschlussverstärkung": ("vorgang", "vorgang", "optionen", "Änderung eines bestehenden Netzanschlusses: Netzanschlussverstärkung"),
        "Änderung eines bestehenden Netzanschlusses: Stilllegung": ("vorgang", "vorgang", "optionen", "Änderung eines bestehenden Netzanschlusses: Stilllegung"),
        "Änderung eines bestehenden Netzanschlusses: Unterbrechung": ("vorgang", "vorgang", "optionen", "Änderung eines bestehenden Netzanschlusses: Unterbrechung"),
        "Änderung eines bestehenden Netzanschlusses: Trennung": ("vorgang", "vorgang", "optionen", "Änderung eines bestehenden Netzanschlusses: Trennung"),
        "Änderung eines bestehenden Netzanschlusses: Wiederinbetriebnahme": ("vorgang", "vorgang", "optionen", "Änderung eines bestehenden Netzanschlusses: Wiederinbetriebnahme"),
        "Änderung eines bestehenden Netzanschlusses: Zusammenlegung": ("vorgang", "vorgang", "optionen", "Änderung eines bestehenden Netzanschlusses: Zusammenlegung"),
        "Anschlusspunkt (Dropdown)": "anschlusspunkt",
        "Im Gebäude (Auswahl)": ("anschlusspunkt", "optionen", "Im Gebäude"),
        "Außerhalb des Gebäudes (Auswahl)": ("anschlusspunkt", "optionen", "Außerhalb des Gebäudes"),
        "Freifläche (Auswahl)": ("anschlusspunkt", "optionen", "Freifläche"),
        "Hausanschlusslänge (Textfeld)": ("vorgang", "hausanschlusslaenge"),
        "Hindernisse im Bereich der Anschlussstraße (Radiobutton)": ("vorgang", "hindernisseAnschlussstrasse"),
        "Beschreibung Hindernisse (Textfeld)": ("vorgang", "beschreibungHindernisseAnschlussstrasse"),
        "Hinweistextfeld: \"Termin hindernisfrei\"": ("vorgang", "informationstextTerminHindernisse"),
        "Termin Hindernisfrei (Textfeld)": ("vorgang", "terminHindernissfrei"),
        "Hauseinführung (Dropdown)": "hauseinfuehrung",
        "Einzel-/ Mehrspartenhauseinführung von Anschlussnehmer gestellt (Auswahl)": ("hauseinfuehrung", "optionen", "Einzel-/Mehrspartenhauseinführung von Anschlussnehmer gestellt"),
        "Einzeleinführung, durch VB gestellt (Auswahl)": ("hauseinfuehrung", "optionen", "Einzeleinführung, durch VB gestellt"),
        "Größe Hausanschlusssicherung (Textfeld)": "hausanschlusssicherung",
        "Gleichzeitig benötigte Gesamtleistung (Textfeld)": "gesamtleistungOhneZusatz",
        "Kumulierter Bedarf aller Gewerbeeinheiten (Textfeld)": "kumulierterBedarf",
        "Erwarteter Jahresverbrauch [in kWh]": "jahresverbrauch",
        "Gew.  Fertigstellungstermin": "fertigstellungstermin",
        "Straßenfrontlänge (in Meter)": "frontlaenge",
        "Grundstücksfläche (in qm)": "grundstuecksoberflaeche",
        "Anmeldung steuerbarer Verbrauchsgeräte (Radiobutton)": "anmeldungSteuerbareVerbrauchsgeraete",
        "Erfolgt durch die Umlegung/Erneuerung eine Erhöhung des Netzanschlusswertes (Bei Vorgang Umlegung)": ("vorgang", "erhoehungNetzanschlusswert"),
        "Beschreibung Umlegevorhaben (Bei Vorgang Umlegung)": ("vorgang", "beschreibungUmlegevorhaben"),
        "Gewünschtes Datum der Stillegung (Bei Vorgang Stillegung)": ("vorgang", "stilllegung"),
        "Weiterführende Informationen (bei Vorgang Stilllegung, Unterbrechnung, Trennung)": ("vorgang", "weiterfuehrendeInformationen"),
        "Gewünschtes Datum der Wiederinbetriebnahme (Bei Vorgang Wiederinbetriebnahme)": ("vorgang", "wiederinbetriebnahme"),
        "Elektrische Warmwasserbereitung (Radiobutton)": "warmwasserbereitung",
        "Art der Warmwasserbereitung (Dropdown)": "artDerWarmwasserbereitung",
        "Zentral (Auswahl)": ("artDerWarmwasserbereitung", "optionen", "zentral"),
        "Dezentral (Auswahl)": ("artDerWarmwasserbereitung", "optionen", "dezentral"),
        "Baustrom (Radiobutton)": "baustrom",
        "Eigenleistung (Radiobutton)": "eigenleistung",
        "Beschreibung der Eigenleistung (Textfeld)": "eigenleistungBeschreibung",
        "Dienstleistung (Textfeld)": "dienstleistung",
        "Leistungsaufstellung (Dokumentenuploadfeld)": ("dateien", "leistungsaufstellung"),
        "Projektschaubild/Zählerschema (Dokumentenuploadfeld)": ("dateien", "projektschaubildZaehlerschema"),
    }

    # Normalisiertes Mapping
    mapping = {key.strip().lower(): val for key, val in original_mapping.items()}

    # === 2. JSON LADEN ===
    with open("Strom/assets/Strom_aktualisiert.json", "r", encoding="utf-8") as f:
        strom_data = json.load(f)

    # === 3. EXCEL EINLESEN ===
    df = pd.read_excel("Strom/assets/ANLAGE Abfrageformular Teilantrag Strom 2.6.xlsx", sheet_name="Abfrage optionale Felder", header=None)
    relevant_rows = df.iloc[15:59, [2, 3, 4, 5]]
    relevant_rows.columns = ["label", "status", "pflicht", "hinweis"]


    # === 4. UPDATE-FUNKTION ===
    def update_json(data, key_path, status_value, required_value, hint_text):
        container = data

        if isinstance(key_path, str):
            if key_path in container:
                entry = container[key_path]
                if isinstance(entry, dict):
                    entry["aktiv"] = status_value
                    if "erforderlich" in entry:
                        entry["erforderlich"] = required_value
                    if "hinweistext" in entry:
                        entry["hinweistext"] = hint_text
                else:
                    print(f"⚠️  Kein dict bei: {key_path}")
            else:
                print(f"⚠️  Schlüssel nicht gefunden: {key_path}")
            return

        for key in key_path[:-1]:
            container = container.get(key, {})
            if not isinstance(container, (dict, list)):
                print(f"⚠️  Unerwartete Struktur bei: {key_path}")
                return

        last_key = key_path[-1]

        if isinstance(container, dict):
            if last_key in container:
                entry = container[last_key]
                if isinstance(entry, dict):
                    entry["aktiv"] = status_value
                    if "erforderlich" in entry:
                        entry["erforderlich"] = required_value
                    if "hinweistext" in entry:
                        entry["hinweistext"] = hint_text
                else:
                    print(f"⚠️  Kein dict bei: {key_path}")
            else:
                print(f"⚠️  Schlüssel nicht gefunden: {key_path}")

        elif isinstance(container, list):
            found = False
            for item in container:
                if item.get("label", "").strip().lower() == last_key.strip().lower():
                    item["aktiv"] = status_value
                    if "erforderlich" in item:
                        item["erforderlich"] = required_value
                    if "hinweistext" in item:
                        item["hinweistext"] = hint_text
                    found = True
            if not found:
                print(f"⚠️  Kein Listeneintrag mit label='{last_key}' gefunden für: {key_path}")




    # === 5. ANWENDUNG DES MAPPINGS ===
    strom_obj = strom_data["strom"]

    for i, row in relevant_rows.iterrows():
        label = str(row["label"]).strip().lower()
        status = str(row["status"]).strip().lower()
        pflicht = str(row["pflicht"]).strip().lower()
        hinweis_raw = str(row["hinweis"]).strip()

        # === Status-Auswertung ===
        if not status or "keine änderung" in status or "nicht änderbar" in status:
            print(f"ℹ️  Keine Änderung für Label: {row['label']}")
            continue
        status_value = False if "aus" in status else True

        # === Pflicht-Auswertung ===
        if "pflicht" in pflicht or "keine änderung" in pflicht:
            required_value = True
        else:
            required_value = False

        # === Hinweistext-Auswertung ===
        if "keine änderung" in hinweis_raw.lower():
            hint_text = None  # mache nichts (eigentlich: überspringen, aber update_json erwartet Wert)
        elif i + 15 == 33 and hinweis_raw.startswith("Beispiel im Testlink:"):
            hint_text = None
        elif hinweis_raw == "" or hinweis_raw.lower() == "nan":
            hint_text = None
        else:
            hint_text = hinweis_raw

        if label in mapping:
            update_json(strom_obj, mapping[label], status_value, required_value, hint_text)
        else:
            print(f"⚠️ Kein Mapping für Label gefunden: {row['label']}")


    # === 6. SPEICHERN ===
    with open("output/Strom/Strom_automatisiert.json", "w", encoding="utf-8") as f:
        json.dump(strom_data, f, indent=2, ensure_ascii=False)