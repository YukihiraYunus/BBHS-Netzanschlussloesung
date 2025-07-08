def run():
    import pandas as pd
    import json

    # === 1. Mapping ===
    # === Excl musste geändert werden (2x Nein)
    mapping = {
        # Vorgang
        "Vorgang (Dropdown)": ("vorgang", "vorgang"),
        "Anmeldung": ("vorgang", "vorgang", "optionen", "Anmeldung"),
        "Anlagen- und Anschlussveränderung": ("vorgang", "vorgang", "optionen", "Anlagen- und Anschlussveränderung"),
        "Stilllegung": ("vorgang", "vorgang", "optionen", "Stilllegung"),
        "Datum der (geplanten) technischen Inbetriebsetzung": ("vorgang", "datumTechnischeInbetriebnahme"),
        "Datum der technischen Außerbetriebsetzung": ("vorgang", "datumTechnischeAuserbetriebnahme"),

        # Angaben zum Gerät
        "Hersteller der Wärmepumpe": ("angebenGeraet", "herstellerWaermepumpe"),
        "Typ der Wärmepumpe": ("angebenGeraet", "typWaermepumpe"),
        "Art der Wärmepumpe (Dropdownfeld)": ("angebenGeraet", "artwaermepumpe"),
        "Luft-Luft": ("angebenGeraet", "artwaermepumpe", "optionen", "Luft-Luft"),
        "Luft-Wasser": ("angebenGeraet", "artwaermepumpe", "optionen", "Luft-Wasser"),
        "Wasser-Wasser": ("angebenGeraet", "artwaermepumpe", "optionen", "Wasser-Wasser"),
        "Sole-Wasser": ("angebenGeraet", "artwaermepumpe", "optionen", "Sole-Wasser"),
        "Sonstiges": ("angebenGeraet", "artwaermepumpe", "optionen", "Sonstiges"),
        "Elektrische Leistung (ohne Zusatzheizung) in kW": ("angebenGeraet", "elektrischeLeistungWarmwasser"),
        "Elektrische Leistung der Zusatzheizung in kW": ("angebenGeraet", "elektrischeLeistungRaumheizung"),
        "Maximaler Anlaufstrom der Wärmepumpe in Ampere (A)": ("angebenGeraet", "maximalerAnlaufstrom"),
        "Maximale Netzbezugsleistung der Wärmepumpe in kW": ("angebenGeraet", "maximaleNetzbezugsleistungWaermepumpe"),
        "Anzahl baugleicher Wärmepumpen": ("angebenGeraet", "anzahlBaugleicherWaermepumpen"),
        "Gesamtleistung der Wärmepumpe(n) in kW": ("angebenGeraet", "gesamtleistungWaermepumpen"),

        # Kundenanlage / Messkonzept
        "Mess- und Betriebskonzept (Dropdownfeld)": ("messkonzept", "messkonzept"),
        "Messkonzept 1": ("messkonzept", "messkonzept", "optionen", "Messkonzept 1"),
        "Messkonzept 2": ("messkonzept", "messkonzept", "optionen", "Messkonzept 2"),
        "Messkonzept 3": ("messkonzept", "messkonzept", "optionen", "Messkonzept 3"),
        "Abweichendes Messkonzept": ("messkonzept", "messkonzept", "optionen", "Abweichendes Messkonzept"),
        "Abweichendes Mess- und Betriebskonzept (Uploadfeld)": ("messkonzept", "messBetriebskonzept"),
        "Ist bereits ein zu nutzender Zähler vorhanden?": ("messkonzept", "nutzenderZaehler"),
        "Soll die Messung über einen separaten Zähler erfolgen?": ("messkonzept", "separaterZaehler"),
        "Angabe der Zählernummer des vorhandenen Zählers": ("messkonzept", "zaehlernummer"),

        # Lastmanagement
        "Lastenmanagement (Dropdownfeld)": ("messkonzept", "lastmanagement"),
        "Nein": ("messkonzept", "lastmanagement", "optionen", "Nein"),
        "Ja (dynamisch)": ("messkonzept", "lastmanagement", "optionen", "Ja (dynamisch)"),
        "Ja (statisch)": ("messkonzept", "lastmanagement", "optionen", "Ja (statisch)"),

        # Zusatzinformationen
        "Wird die Wärmepumpe für betriebsnotwendige Zwecke oder einer kritischen Infrastruktur eingesetzt? (Ja/Nein)": ("waermepumpebetriebsnotwendigeZwecke",),
        "Wird eine Leistungserhöhung der Hausanschlusssicherung benötigt? (Ja/Nein)": ("leistungserhoehungHausanschlusssicherung",),
        "Benötigte Hausanschlusssicherung in Ampere (A)": ("benoetigteHausanschlusssicherung",),

        # Zusatzabfragen
        "[Zusatzabfrage 1]": ("zusatzabfragen",),
        "[Zusatzabfrage 2]": ("zusatzabfragen",),
        "[Zusatzabfrage x]": ("zusatzabfragen",),

        # Steuerung
        "Handelt es sich um eine Bestandsanlage vor dem 01.01.2024? (Dropdown)": ("angabenSteuerung", "individuelleVereinbarung"),
        "Ja (Anlage ohne Steuerung §14a alte Fassung)": ("angabenSteuerung", "individuelleVereinbarung", "optionen", "Ja (Anlage ohne Steuerung § 14a alte Fassung)"),
        "Ja (Anlage mit Steuerung§14a alte Fassung)": ("angabenSteuerung", "individuelleVereinbarung", "optionen", "Ja (Anlage mit Steuerung §14a  alte Fassung)"),
        "Nein (nicht vor dem 01.01.2024)": ("angabenSteuerung", "individuelleVereinbarung", "optionen", "Nein"),
        "Ist trotz des Bestandsschutzes ein Wechsel in die freiwillige Steuerbarkeit gewünscht? (Radiobutton)": ("angabenSteuerung", "wechselFreiwilligeSteuerbarkeit"),
        "Wählen Sie aus, wie die Steuerung umgesetzt wird. (Dropdown)": ("angabenSteuerung", "steuerungsart"),
        "Direkt": ("angabenSteuerung", "steuerungsart", "optionen", "Direkt"),
        "Energiemanagementsystem": ("angabenSteuerung", "steuerungsart", "optionen", "Energiemanagementsystem"),
        "Wählen Sie aus, welches Modul der Netzentgeldreduzierung nach §14a EnWG Sie anwenden möchten.": ("angabenSteuerung", "netzentgeltreduzierung"),
        "Modul 1: Pauschale Reduzierung der Netzentgelte": ("angabenSteuerung", "netzentgeltreduzierung", "optionen", "Modul 1: Pauschale Reduzierung der Netzentgelte"),
        "Modul 2: Prozentuale Reduzierung des Arbeitspreises": ("angabenSteuerung", "netzentgeltreduzierung", "optionen", "Modul 2: Prozentuale Reduzierung des Arbeitspreises"),
        "Wählen Sie aus, wer für die Herstellung der Steuerbarkeit beauftragt wird. (Dropdown)": ("angabenSteuerung", "herstellungSteuerbarkeit"),
        "Grundzuständigen Messstellenbetreiber": ("angabenSteuerung", "herstellungSteuerbarkeit", "optionen", "Grundzuständigen Messstellenbetreiber"),
        "Wettbewerblichen Messstellenbetreiber": ("angabenSteuerung", "herstellungSteuerbarkeit", "optionen", "Wettbewerblicher Messstellenbetreiber"),
        "Netzbetreiber": ("angabenSteuerung", "herstellungSteuerbarkeit", "optionen", "Netzbetreiber"),
        "Angabe des wettbewerlichen Messstellenbetreibers": ("angabenSteuerung", "messstellenbetreiber"),

        # Dokumente
        "Lageplan (Uploadfeld)": ("dateien", "lageplan"),
        "Datenblatt (Uploadfeld)": ("dateien", "datenblatt"),
        "Zusatzfeld Upload (Uploadfeld)": ("dateien", "zusatzdatei"),

        # AGB
        "AGB (Checkbox)": ("agb",),
    }


    # === 2. JSON laden ===
    with open("Verbrauchsgeräte/assets/Wärmepumpe.json", "r", encoding="utf-8") as f:
        data = json.load(f)

    # === 3. Excel laden ===
    df = pd.read_excel("Verbrauchsgeräte/assets/ANLAGE Abfrageformular Teilantrag Verbrauchsgeräte.xlsx", sheet_name="Abfrage optionale Felder", header=None)
    df.columns = [f"col_{i}" for i in range(df.shape[1])]
    relevant_rows = df.iloc[90:159, [2, 4, 5, 6]]
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
    pump_obj = data["waermepumpe"]
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
            update_json(pump_obj, mapping[label], status_value, required_value, hint_text)
        else:
            print(f"⚠️ Kein Mapping für Label gefunden: {label}")

    # === 6. Speichern ===
    with open("output/Verbrauchsgeräte/wärmepumpe_automatisiert.json", "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
