def run():
    import pandas as pd
    import json

    # === 1. MAPPING ===
    original_mapping = {
        "Auswahl Vorgang (Dropdown)": ("vorgang", "vorgang"),
        "Neuer Netzanschluss": ("vorgang", "vorgang", "optionen", "Ich möchte einen neuen Netzanschluss"),
        "Zeitlich befristeter Netzanschluss": ("vorgang", "vorgang", "optionen", "Zeitlich befristeter Netzanschluss"),
        "Änderung eines bestehenden Netzanschlusses: Umlegung": ("vorgang", "vorgang", "optionen", "Änderung eines bestehenden Netzanschlusses: Umlegung"),
        "Änderung eines bestehenden Netzanschlusses: Netzanschlussverstärkung": ("vorgang", "vorgang", "optionen", "Änderung eines bestehenden Netzanschlusses: Netzanschlussverstärkung"),
        "Änderung eines bestehenden Netzanschlusses: Stilllegung": ("vorgang", "vorgang", "optionen", "Änderung eines bestehenden Netzanschlusses: Stilllegung"),
        "Änderung eines bestehenden Netzanschlusses: Unterbrechung": ("vorgang", "vorgang", "optionen", "Änderung eines bestehenden Netzanschlusses: Unterbrechung"),
        "Änderung eines bestehenden Netzanschlusses: Trennung": ("vorgang", "vorgang", "optionen", "Änderung eines bestehenden Netzanschlusses: Trennung"),
        "Änderung eines bestehenden Netzanschlusses: Wiederinbetriebnahme": ("vorgang", "vorgang", "optionen", "Änderung eines bestehenden Netzanschlusses: Wiederinbetriebnahme"),
        "Anzahl Gasgeräte (Textfeld)": ("gasgeraete"),
        "Vorzuhaltende Anschlussleistung am Übergabepunkt (Textfeld)": ("anschlussleistung"),
        "Erfolgt durch die Umlegung/Erneuerung eine Erhöhung des Netzanschlusswertes (Bei Vorgang Umlegung)": ("vorgang", "erhoehungNetzanschlusswert"),
        "Straßenfrontlänge (in Meter)": ("frontlaenge"),
        "Grundstücksfläche (in qm)": ("grundstuecksoberflaeche"),
        "gew. Fertigstellungstermin": ("fertigstellungstermin"),
        "Beschreibung Umlegevorhaben (Bei Vorgang Umlegung)": ("vorgang", "beschreibungUmlegevorhaben"),
        "Gewünschtes Datum der Stillegung (Bei Vorgang Stillegung)": ("vorgang", "stilllegung"),
        "Weiterführende Informationen (Bei Vorgang Stilllegung, Unterbrechung, Trennung)": ("vorgang", "weiterfuehrendeInformationen"),
        "Gewünschtes Datum der Wiederinbetriebnahme (Bei Vorgang Wiederinbetriebnahme)": ("vorgang", "wiederinbetriebnahme"),
        "Heizgas, Prozessgas, Kochgas (Checkbox)": ("verwendungszweck"),
        "Art der Versorgung (Dropdown)": ("versorgung"),
        "Vollversorgung": ("versorgung", "optionen", "Vollversorgung"),
        "Reserveversorgung": ("versorgung", "optionen", "Reserveversorgung")
    }

    mapping = {key.strip().lower(): val for key, val in original_mapping.items()}

    # === 2. JSON LADEN ===
    with open("Gas/assets/Gas.json", "r", encoding="utf-8") as f:
        gas_data = json.load(f)

    # === 3. EXCEL LADEN ===
    df = pd.read_excel("Gas/assets/ANLAGE Abfrageformular Teilantrag Gas 2.6.xlsx", sheet_name="Abfrage optionale Felder", header=None)
    relevant_rows = df.iloc[14:40, [2, 3, 4, 5]]
    relevant_rows.columns = ["label", "status", "pflicht", "hinweis"]

    # === 4. UPDATE-FUNKTION ===
    def update_json(data, key_path, status_value, required_value, hint_text):
        container = data
        if isinstance(key_path, str):
            entry = container.get(key_path)
            if isinstance(entry, dict):
                if status_value is not None:
                    entry["aktiv"] = status_value
                if "erforderlich" in entry:
                    entry["erforderlich"] = required_value
                if "hinweistext" in entry:
                    entry["hinweistext"] = hint_text
            return

        for key in key_path[:-1]:
            container = container.get(key, {})
        last_key = key_path[-1]
        if isinstance(container, dict) and last_key in container:
            entry = container[last_key]
            if isinstance(entry, dict):
                if status_value is not None:
                    entry["aktiv"] = status_value
                if "erforderlich" in entry:
                    entry["erforderlich"] = required_value
                if "hinweistext" in entry:
                    entry["hinweistext"] = hint_text
        elif isinstance(container, list):
            for item in container:
                if item.get("label", "").strip().lower() == last_key.strip().lower():
                    if status_value is not None:
                        item["aktiv"] = status_value
                    if "erforderlich" in item:
                        item["erforderlich"] = required_value
                    if "hinweistext" in item:
                        item["hinweistext"] = hint_text

    # === 5. ANWENDUNG DES MAPPINGS ===
    gas_obj = gas_data["gas"]
    for i, row in relevant_rows.iterrows():
        label = str(row["label"]).strip().lower()
        status = str(row["status"]).strip().lower()
        pflicht = str(row["pflicht"]).strip().lower()
        hinweis_raw = str(row["hinweis"]).strip()

        status_value = None if not status or "keine änder" in status else ("aus" not in status)
        required_value = "pflicht" in pflicht or "keine änder" in pflicht
        hint_text = None if "keine änder" in hinweis_raw.lower() else (
            None if hinweis_raw.strip() == "" or hinweis_raw.lower() == "nan" else hinweis_raw
        )

        if label in mapping:
            update_json(gas_obj, mapping[label], status_value, required_value, hint_text)
        else:
            print(f"⚠️ Kein Mapping für Label gefunden: {row['label']}")

    # === 6. SPEICHERN ===
    with open("output/Gas/Gas_automatisiert.json", "w", encoding="utf-8") as f:
        json.dump(gas_data, f, indent=2, ensure_ascii=False)
