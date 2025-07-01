def run():
    import pandas as pd
    import json

    # === 1. MAPPING ===   

    # === NOTIZEN ===
    # === "Bisherige Wärmeversorgung des Gebäudes": ("bisherigeWaermeversorgung"),
    # === Excl nicht bearbeitbar für dieses Feld
    #
    # === Anschlussleistung in kW (Bei Auswahl: Raumheizung; Gebrauchswarmwasser;Lüftung))
    # === Ist in Excl aber nicht in der Config
    #
    # === Excl musste geändert werden weil 3 mal das gleiche Key vorkommt
    # === Wenn ausgewählt = Textfeld für Anschlussleistung in kW => Wenn ausgewählt = Textfeld für Anschlussleistung in kW (Raumheizung)
    mapping = {
        "Neuer Netzanschluss": ("vorgang", "vorgang", "optionen", "Ich möchte einen neuen Hausanschluss"),
        "Änderung eines bestehenden Netzanschlusses: Umlegung": ("vorgang", "vorgang", "optionen", "Änderung eines bestehenden Hausanschlusses: Umlegung"),
        "Änderung eines bestehenden Netzanschlusses: Hausanschlussverstärkung": ("vorgang", "vorgang", "optionen", "Änderung eines bestehenden Hausanschlusses: Hausanschlussverstärkung"),
        "Änderung eines bestehenden Netzanschlusses: Reduzierung der Anschlussleistung": ("vorgang", "vorgang", "optionen", "Änderung eines bestehenden Hausanschlusses: Reduzierung der Anschlussleistung"),
        "Änderung eines bestehenden Netzanschlusses: Stilllegung": ("vorgang", "vorgang", "optionen", "Änderung eines bestehenden Hausanschlusses: Stilllegung"),
        "Änderung eines bestehenden Netzanschlusses: Unterbrechung": ("vorgang", "vorgang", "optionen", "Änderung eines bestehenden Hausanschlusses: Unterbrechung"),
        "Änderung eines bestehenden Netzanschlusses: Trennung": ("vorgang", "vorgang", "optionen", "Änderung eines bestehenden Hausanschlusses: Trennung"),
        "Änderung eines bestehenden Netzanschlusses: Wiederinbetriebnahme": ("vorgang", "vorgang", "optionen", "Änderung eines bestehenden Hausanschlusses: Wiederinbetriebnahme"),
        "Änderung eines bestehenden Netzanschlusses: Beseitigung": ("vorgang", "vorgang", "optionen", "Änderung eines bestehenden Hausanschlusses: Beseitigung"),
        "Vorzuhaltende Anschlussleistung am Übergabepunkt [kW] (Checkbox)": ("anschlussleistung"),
        "Raumheizung (Auswahl)": ("raumheizung"),
        "Wenn ausgewählt = Textfeld für Anschlussleistung in kW (Raumheizung)": ("raumheizungAnschlussleistung"),
        "Gebrauchswarmwasser (Auswahl)": ("gebrauchswarmwasser"),
        "Wenn ausgewählt = Textfeld für Anschlussleistung in kW (Gebrauchswarmwasser)": ("gebrauchswarmwasserAnschlussleistung"),
        "Lüftung (Auswahl)": ("lueftung"),
        "Wenn ausgewählt = Textfeld für Anschlussleistung in kW (Lüftung)": ("lueftungAnschlussleistung"),
        "Smart-Meter-Gateway (Radiobutton)": ("smartMeterGateway"),
        "Vertraglicher Anschlusswert (Eingabefeld)": ("vertraglicherAnschlusswert"),
        "Nachweis Berechnung (Uploadfeld)": ("nachweisBerechnung"),
        "Volumenstrom (Eingabefeld)": ("volumenstrom"),
        "Vorlauftemperatur (Eingabefeld)": ("vorlauftemperatur"),
        "Angabe der Rücklauftemperatur (Hinweisfeld)": ("ruecklauftemperatur"),
        "Zählernummer": ("zaehlernummer"),
        "Straßenfrontlänge ( in Meter)": ("frontlaenge"),
        "Grundstücksfläche ( in qm)": ("grundstuecksoberflaeche"),
        "Hausanschlusslänge (in Meter)": ("hausanschlusslaenge"),
        "gew. Lieferbeginn": ("gewuenschterLieferbeginn"),
        "gew. Fertigstellungstermin": ("fertigstellungstermin"),
        "Erfolgt durch die Umlegung/Erneuerung eine Erhöhung des Netzanschlusswertes (Bei Vorgang Umlegung)": ("vorgang", "erhoehungNetzanschlusswert"),
        "Beschreibung Umlegevorhaben (Bei Vorgang Umlegung)": ("vorgang", "beschreibungUmlegevorhaben"),
        "Gewünschtes Datum der Stillegung (Bei Vorgang Stillegung)": ("vorgang", "stilllegung"),
        "Weiterführende Informationen (Bei Vorgang Stilllegung, Unterbrechung, Trennung)": ("vorgang", "weiterfuehrendeInformationen"),
        "Gewünschtes Datum der Wiederinbetriebnahme (Bei Vorgang Wiederinbetriebnahme)": ("vorgang", "wiederinbetriebnahme"),

        "Eigenversorgung durch Gebäudeeigentümer (Auswahl)": ("eigenversorgung"),
        "Gewerbliche Wärmelieferung (Auswahl)": ("gewerblich"),
        "Auswahl Contracting/Fernwärme bei gewerbliche Wärmelieferung": ("contractingFernwaerme"),

        "Zusatzabfragen": ("zusatzabfragen", "Zusatzabfragen"),
        "Lageplan des Grundstücks": ("dateien", "plan"),
        "Grundrissskizze des Gebäudes": ("dateien", "skizze"),
        "Skizze des Hausanschlussraums": ("dateien", "zusatzdatei"),
        "Preisblatt": ("preisblatt"),
    }

    # === 2. JSON laden ===
    with open("Fernwärme/assets/Fernwärme.json", "r", encoding="utf-8") as f:
        data = json.load(f)

    # === 3. Excel laden ===
    df = pd.read_excel("Fernwärme/assets/ANLAGE Abfrageformular Teilantrag Fernwärme 2.6.xlsx", sheet_name="Abfrage optionale Felder", header=None)
    df.columns = [f"col_{i}" for i in range(df.shape[1])]
    relevant_rows = df.iloc[15:54, [2, 3, 4, 5]]
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
    fern_obj = data["fernwaerme"]
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
            update_json(fern_obj, mapping[label], status_value, required_value, hint_text)
        else:
            print(f"⚠️ Kein Mapping für Label gefunden: {label}")

    # === 6. Speichern ===
    with open("output/Fernwärme/fernwärme_automatisiert.json", "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
