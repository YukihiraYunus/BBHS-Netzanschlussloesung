def run():
    import pandas as pd
    import json

    # === 1. Mapping ===
    # === Excl musste geändert werden (2x Nein)
    mapping = {
        "Vorgang (Dropdownfeld)": ("vorgang", "vorgang"),
        "Anmeldung": ("vorgang", "vorgang", "optionen", "Anmeldung"),
        "Anlagen- und Anschlussveränderung": ("vorgang", "vorgang", "optionen", "Anlagen- und Anschlussveränderung"),
        "Stilllegung": ("vorgang", "vorgang", "optionen", "Stilllegung"),
        "Datum der geplanten technischen Inbetriebsetzung (Textfeld)": ("vorgang", "datumTechnischeInbetriebnahme"),
        "Datum der technischen Außerbetriebsetzung (Textfeld)": ("vorgang", "datumTechnischeAuserbetriebnahme"),
        
        # Angaben zum Gerät
        "Hersteller der Klimaanlage": ("geraete", "hersteller"),
        "Typ der Klimaanlage": ("geraete", "typ"),
        "Elektrische Leistung der Klimaanlage in kW": ("geraete", "leistung"),
        "Maximaler Anlaufstrom der Klimaanlage in Ampere (A)": ("geraete", "maxAnlaufstrom"),
        "Maximale Netzbezugsleistung der Klimaanlage in kW": ("geraete", "maxNetzbezugsleistung"),
        "Anzahl baugleicher Klimaanlagen": ("geraete", "anzahl"),
        "Gesamtleistung der Klimaanlage(n) in kW": ("geraete", "gesamtleistung"),

        # Kundenanlage
        "Mess- und Betriebskonzept (Dropdownfeld)": ("kundenanlage", "messkonzept"),
        "Ist bereits ein zu nutzender Zähler vorhanden? (Radiobutton)": ("kundenanlage", "nutzenderZaehler"),
        "Soll die Messung über einen separaten Zähler erfolgen? (Radiobutton)": ("kundenanlage", "separaterZaehler"),
        "Angabe der Zählernummer des vorhandenen Zählers (Textfeld)": ("kundenanlage", "angabeZaehlernummer"),
        "Wählen Sie aus, ob das Gerät über ein Lastmanagement gespeichert wird und wenn ja, wie die Steuerung erfolgt. (Dropdownfeld)": ("kundenanlage", "lastmanagement"),
        "Nein": ("kundenanlage", "lastmanagement", "optionen", "Nein"),
        "Ja (dynamisch)": ("kundenanlage", "lastmanagement", "optionen", "Ja (dynamisch)"),
        "Ja (statisch)": ("kundenanlage", "lastmanagement", "optionen", "Ja (statisch)"),
        "Wird die Klimaanlage für betriebsnotwendige Zwecke oder ein einer kritischen Infrastruktur eingesetzt? (Radiobutton)": ("kundenanlage", "betriebsnotwendigeZwecke"),

        # Angaben zur Steuerung
        "Handelt es sich um eine Bestandsanlage vor dem 01.01.2024? (Dropdown)": ("steuerung", "handeltBestandsanlage"),
        "Ja (Anlage ohne Steuerung §14a alte Fassung)": ("steuerung", "handeltBestandsanlage", "optionen", "Ja (Anlage ohne Steuerung § 14a alte Fassung)"),
        "Ja (Anlage mit Steuerung§14a alte Fassung)": ("steuerung", "handeltBestandsanlage", "optionen", "Ja (Anlage mit Steuerung §14a  alte Fassung)"),
        "Nein (nicht vor dem 01.01.2024)": ("steuerung", "handeltBestandsanlage", "optionen", "Nein"),
        "Ist trotz des Bestandsschutzes ein Wechsel in die freiwillige Steuerbarkeit gewünscht? (Radiobutton)": ("steuerung", "wechselFreiwilligeSteuerbarkeit"),
        "Wählen Sie aus, wie die Steuerung umgesetzt wird. (Dropdown)": ("steuerung", "steuerungsart"),
        "Direkt": ("steuerung", "steuerungsart", "optionen", "Direkt"),
        "Energiemanagementsystem": ("steuerung", "steuerungsart", "optionen", "Energiemanagementsystem"),
        "Wählen Sie aus, welches Modul der Netzentgeldreduzierung nach §14a EnWG Sie anwenden möchten.": ("steuerung", "netzentgeltreduzierung"),
        "Modul 1: Pauschale Reduzierung der Netzentgelte": ("steuerung", "netzentgeltreduzierung", "optionen", "Modul 1: Pauschale Reduzierung der Netzentgelte"),
        "Modul 2: Prozentuale Reduzierung des Arbeitspreises": ("steuerung", "netzentgeltreduzierung", "optionen", "Modul 2: Prozentuale Reduzierung des Arbeitspreises"),
        "Wählen Sie aus, wer für die Herstellung der Steuerbarkeit beauftragt wird. (Dropdown)": ("steuerung", "hersteller"),
        "Grundzuständigen Messstellenbetreiber": ("steuerung", "hersteller", "optionen", "Grundzuständigen Messstellenbetreiber"),
        "Wettbewerblichen Messstellenbetreiber": ("steuerung", "hersteller", "optionen", "Wettbewerblicher Messstellenbetreiber"),
        "Netzbetreiber": ("steuerung", "hersteller", "optionen", "Netzbetreiber"),
        "Angabe des wettbewerlichen Messstellenbetreibers": ("steuerung", "messstellenbetreiber"),

        # Dokumentenupload
        "Lageplan (Uploadfeld)": ("dateien", "lageplan"),
        "Datenblatt (Uploadfeld)": ("dateien", "datenblatt"),
        "Zusatzfeld Upload (Uploadfeld)": ("dateien", "zusatzdatei"),

        # AGB
        "AGB (Checkbox)": ("agb",)
    }

    # === 2. JSON laden ===
    with open("Verbrauchsgeräte/assets/Raumkühlung.json", "r", encoding="utf-8") as f:
        data = json.load(f)

    # === 3. Excel laden ===
    df = pd.read_excel("Verbrauchsgeräte/assets/ANLAGE Abfrageformular Teilantrag Verbrauchsgeräte.xlsx", sheet_name="Abfrage optionale Felder", header=None)
    df.columns = [f"col_{i}" for i in range(df.shape[1])]
    relevant_rows = df.iloc[244:305, [2, 4, 5, 6]]
    relevant_rows.columns = ["label", "status", "pflicht", "hinweis"]

    # === 4. Update Funktion ===
    def update_json(data, key_path, status_value=None, required_value=None, hint_text=None):
        container = data

        # Fall: einfacher Schlüssel (keine Schachtelung)
        if isinstance(key_path, str):
            if key_path in container and isinstance(container[key_path], dict):
                if status_value is not None and "aktiv" in container[key_path]:
                    container[key_path]["aktiv"] = status_value
                if required_value is not None and "erforderlich" in container[key_path]:
                    container[key_path]["erforderlich"] = required_value
                if hint_text is not None and "hinweistext" in container[key_path]:
                    container[key_path]["hinweistext"] = hint_text
            else:
                print(f"⚠️ Schlüssel '{key_path}' nicht im JSON gefunden oder keine dict-Struktur.")
                return

        # Sonst: durch das Pfad-Tuple durchiterieren
        for key in key_path[:-1]:
            if key not in container:
                print(f"⚠️ Pfadbestandteil '{key}' fehlt im Pfad {key_path}")
                return
            container = container[key]
            if not isinstance(container, (dict, list)):
                print(f"⚠️ Strukturfehler bei '{key}' im Pfad {key_path} (nicht dict oder list)")
                return

        last_key = key_path[-1]

        # Fall: letzte Ebene ist Liste von dicts → Suche nach label
        if isinstance(container, list):
            found = False
            for item in container:
                if item.get("label", "").strip().lower() == last_key.strip().lower():
                    if status_value is not None and "aktiv" in item:
                        item["aktiv"] = status_value
                    if required_value is not None and "erforderlich" in item:
                        item["erforderlich"] = required_value
                    if "hinweistext" in item:
                        item["hinweistext"] = hint_text
                    found = True
                    break
            if not found:
                print(f"⚠️ Kein Listeneintrag mit label='{last_key}' gefunden im Pfad {key_path}")

        # Fall: letzte Ebene ist dict → direkter Zugriff
        elif isinstance(container, dict):
            entry = container.get(last_key)
            if isinstance(entry, dict):
                if status_value is not None and "aktiv" in entry:
                    entry["aktiv"] = status_value
                if required_value is not None and "erforderlich" in entry:
                    entry["erforderlich"] = required_value
                if "hinweistext" in entry:
                    entry["hinweistext"] = hint_text
            else:
                print(f"⚠️ '{last_key}' in {key_path} nicht gefunden oder keine dict-Struktur.")


    # === 5. Anwenden des Mappings ===
    raum_obj = data["raumkuehlung"]
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
            update_json(raum_obj, mapping[label], status_value, required_value, hint_text)
        else:
            print(f"⚠️ Kein Mapping für Label gefunden: {label}")

    # === 6. Speichern ===
    with open("output/Verbrauchsgeräte/raumkühlung_automatisiert.json", "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
