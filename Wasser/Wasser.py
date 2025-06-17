def run():
    import pandas as pd
    import json

    # === 1. MAPPING ===
    mapping = {
        "Neuer Netzanschluss": ("vorgang", "vorgang", "optionen", "Ich möchte einen neuen Netzanschluss"),
        "Zeitlich befristeter Netzanschluss": ("vorgang", "vorgang", "optionen", "Zeitlich befristeter Netzanschluss"),
        "Änderung eines bestehenden Netzanschlusses: Umlegung": ("vorgang", "vorgang", "optionen", "Änderung eines bestehenden Hausanschlusses: Umlegung"),
        "Änderung eines bestehenden Netzanschlusses: Leistungsänderung": ("vorgang", "vorgang", "optionen", "Änderung eines bestehenden Hausanschlusses: Leistungsänderung"),
        "Änderung eines bestehenden Netzanschlusses: Stilllegung": ("vorgang", "vorgang", "optionen", "Änderung eines bestehenden Netzanschlusses: Stilllegung"),
        "Änderung eines bestehenden Netzanschlusses: Unterbrechung": ("vorgang", "vorgang", "optionen", "Änderung eines bestehenden Netzanschlusses: Unterbrechung"),
        "Änderung eines bestehenden Netzanschlusses: Trennung": ("vorgang", "vorgang", "optionen", "Änderung eines bestehenden Netzanschlusses: Trennung"),
        "Änderung eines bestehenden Netzanschlusses: Wiederinbetriebnahme": ("vorgang", "vorgang", "optionen", "Änderung eines bestehenden Netzanschlusses: Wiederinbetriebnahme"),
        "Spitzendurchfluss (ohne Löschwasser) in l/s": ("spitzendurchfluss"),
        "Voraussichtliche Bedarfsmenge in m³/h": ("bedarfsmenge"),
        "Straßenfrontlänge (in Meter)": ("frontlaenge"),
        "Grundstücksfläche ( in qm)": ("grundstuecksoberflaeche"),
        "Hausanschlusslänge (in Meter)": ("hausanschlusslaenge"),
        "gew. Fertigstellungstermin": ("fertigstellungstermin"),
        "Erfolgt durch die Umlegung/Erneuerung eine Erhöhung des Netzanschlusswertes (Bei Vorgang Umlegung)": ("vorgang", "erhoehungNetzanschlusswert"),
        "Beschreibung Umlegevorhaben (Bei Vorgang Umlegung)": ("vorgang", "beschreibungUmlegevorhaben"),
        "Gewünschtes Datum der Stillegung (Bei Vorgang Stillegung)": ("vorgang", "stilllegung"),
        "Weiterführende Informationen (Bei Vorgang Stilllegung, Unterbrechung, Trennung": ("vorgang", "weiterfuehrendeInformationen"),
        "Gewünschtes Datum der Wiederinbetriebnahme (Bei Vorgang Wiederinbetriebnahme)": ("vorgang", "wiederinbetriebnahme"),
        "Zählerstand (in kWh)": ("zaehlerstand"),
        "Zählernummer": ("zaehlernummer"),
        "Zusätzliche Einrichtungen mit besonderem Wasserbedarf vorhanden?": ("einrichtungenVorhanden"),
        "Art der Einrichtung": ("artDerEinrichtung"),
        "Bedarf der Einrichtung (in m³/s)": ("bedarfEinrichtung"),
        "Eigenwasserversorgung vorhanden oder geplant?": ("eigenwasserversorgung"),
        "Bauwassersäule (Checkbox)": ("bewaesserungssaeule"),
        "Weitergehende Anforderungen": ("wasser", "weitergehendeAnforderungen"),
        "Lageplan": ("wasser", "dateien", "plan"),
        "Grundrissskizze": ("wasser", "dateien", "skizze"),
        "Sonstige Dateien": ("wasser", "dateien", "zusatzdatei"),
        "Zusatzinformation": ("wasser", "zusatzabfragen", "Zusatzinformation")
    }

    # === 2. JSON laden ===
    with open("Wasser/assets/Wasser.json", "r", encoding="utf-8") as f:
        data = json.load(f)

    # === 3. Excel laden ===
    df = pd.read_excel("Wasser/assets/ANLAGE Abfrageformular Teilantrag Wasser 2.6.xlsx", sheet_name="Abfrage optionale Felder", header=None)
    df.columns = [f"col_{i}" for i in range(df.shape[1])]
    relevant_rows = df.iloc[15:42, [2, 3, 4, 5]]
    relevant_rows.columns = ["label", "status", "pflicht", "hinweis"]

    # === 4. Update Funktion ===
    def update_json(data, key_path, status_value=None, required_value=None, hint_text=None):
        container = data
        if isinstance(key_path, str):
            if key_path in container and isinstance(container[key_path], dict):
                if status_value is not None:
                    container[key_path]["aktiv"] = status_value
                if required_value is not None and "erforderlich" in container[key_path]:
                    container[key_path]["erforderlich"] = required_value
                # Hier: Hinweistext immer setzen, auch wenn es noch nicht existiert
                if hint_text is not None or hint_text is None:
                    container[key_path]["hinweistext"] = hint_text
            return

        for key in key_path[:-1]:
            container = container.get(key, {})
        last_key = key_path[-1]

        if isinstance(container, list):
            for item in container:
                if item.get("label", "").strip().lower() == last_key.strip().lower():
                    if status_value is not None:
                        item["aktiv"] = status_value
                    if required_value is not None and "erforderlich" in item:
                        item["erforderlich"] = required_value
                    if hint_text is not None or hint_text is None:
                        item["hinweistext"] = hint_text
                    return
        elif isinstance(container, dict):
            entry = container.get(last_key)
            if isinstance(entry, dict):
                if status_value is not None:
                    entry["aktiv"] = status_value
                if required_value is not None and "erforderlich" in entry:
                    entry["erforderlich"] = required_value
                if hint_text is not None or hint_text is None:
                    entry["hinweistext"] = hint_text


    # === 5. Anwenden des Mappings ===
    wasser_obj = data["wasser"]
    for _, row in relevant_rows.iterrows():
        label = str(row["label"]).strip()
        status = str(row["status"]).strip().lower()
        pflicht = str(row["pflicht"]).strip().lower()
        hinweis_raw = str(row["hinweis"]).strip()

        status_value = None if "keine änder" in status or status == "" else ("aus" not in status)
        required_value = "pflicht" in pflicht or "keine änder" in pflicht
        hint_text = None if "keine änder" in hinweis_raw.lower() else (
            None if hinweis_raw.strip() == "" or hinweis_raw.lower() == "nan" else hinweis_raw
        )

        if label in mapping:
            update_json(wasser_obj, mapping[label], status_value, required_value, hint_text)
        else:
            print(f"⚠️ Kein Mapping für Label gefunden: {label}")

    # === 6. Speichern ===
    with open("output/Wasser/wasser_automatisiert.json", "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
