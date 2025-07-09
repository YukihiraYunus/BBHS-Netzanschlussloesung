def run():
    import pandas as pd
    import json

    # === 1. Mapping ===
    # === Excl musste geändert werden (2x Nein)
    mapping = {
        # Vorgang
        "Auswahl Vorgang (Dropdown)": ("vorgang", "vorgang"),
        "Anmeldung": ("vorgang", "vorgang", "optionen", "Neuanschluss"),
        "Anlagen- und Anschlussveränderung": ("vorgang", "vorgang", "optionen", "Anlagenerweiterung oder -änderung"),
        "Stilllegung": ("vorgang", "vorgang", "optionen", "Stilllegung"),
        "Datum der geplanten technischen Inbetriebsetzung (Bei Vorgang Neuanschluss und Anlagen- und Anschlussveränderung)": ("inbetriebsetzungstermin",),
        "Datum der geplanten technischen Betriebsbereitschaft nach EEG (Bei Vorgang Neuanschluss)": ("betriebsbereitschaft",),
        "Datum der technischen Außerbetriebsetzung (Bei Vorgang Stilllegung)": ("ausserbetriebsetzung",),
        "SEE-Nummer (Bei Vorgang Stilllegung)": ("seeNummer",),

        # Energieträger / Speicher
        "Anlagenart / Energieträger": ("energietraeger",),
        "Solarenergie": ("energietraeger", "optionen", "Solarenergie"),
        "Wasserkraft": ("energietraeger", "optionen", "Wasserkraft"),
        "Windkraft": ("energietraeger", "optionen", "Windkraft"),
        "Geothermie": ("energietraeger", "optionen", "Geothermie"),
        "Biomasse": ("energietraeger", "optionen", "Biomasse"),
        "Speicher": ("energietraeger", "optionen", "Speicher"),
        "Soll gleichzeitg ein Speicher angemeldet werden bzw. ist in der Anlage ein Speicher integriert?": ("hatSpeichereinheit",),

        # Steuerbare Lasten
        "Befindet sich hinter dem Netzanschlusspunkt eine steuerbare Last nach §14a EnWG, die über 4,2 kW aus dem Netz bezieht.": ("anlage", "steuerbareLast"),
        "Keine": ("anlage", "steuerbareLast", "optionen", "Keine"),
        "Ladepunkte": ("anlage", "steuerbareLast", "optionen", "Ladepunkte"),
        "Wärmepumpe": ("anlage", "steuerbareLast", "optionen", "Wärmepumpe"),
        "Klimaanlage": ("anlage", "steuerbareLast", "optionen", "Klimaanlage"),
        "Stromspeicher": ("anlage", "steuerbareLast", "optionen", "Stromspeicher"),
        "Gesamt-Minderleistung der Anlage am Netzanschluss in kW": ("anlage", "gesamtMinderleistung"),

        # Steuerung §14a
        "Wählen Sie aus, ob der Netzbetreiber über eine bereits verbaute Sterubox vor Ort zugreifen kann. (Steuerung nach § 14a EnWG)": ("anlage", "steuerung14a"),
        "Nutzung vorhandener technischer Einrichtung möglich": ("anlage", "steuerung14a", "optionen", "Nutzung vorhandener technischer Einrichtung möglich"),
        "Auftrag zum Einbau technischer Einrichtugen durch den grundzuständigen Messstellenbetreiber": ("anlage", "steuerung14a", "optionen", "Einbau durch grundzuständigen Messstellenbetreiber"),
        "Auftrag zum Einbau neuer technischer Einrichtungen durch den wettbewerblichen Messstellenbetreiber": ("anlage", "steuerung14a", "optionen", "Einbau durch wettbewerblichen Messstellenbetreiber"),

        # Messkonzept
        "Wählen Sie bitte das Mess- und Betriebskonzept der Anlage aus  (Dropdownfeld)": ("messkonzept", "messkonzept"),
        "Messkonzept 1": ("messkonzept", "messkonzept", "optionen", "Messkonzept 1"),
        "Messkonzept 2": ("messkonzept", "messkonzept", "optionen", "Messkonzept 2"),
        "Messkonzept 3": ("messkonzept", "messkonzept", "optionen", "Messkonzept 3"),
        "Messkonzept x": ("messkonzept", "messkonzept", "optionen", "Messkonzept x"),
        "Abweichendes Messkonzept": ("messkonzept", "messkonzept", "optionen", "Abweichendes Messkonzept"),
        "Abweichendes Mess- und Betriebskonzept (Upload)": ("messkonzept", "messBetriebskonzeptUpload"),
        "Laden Sie bitte den Schaltplan der Anlage hoch. (Upload)": ("messkonzept", "schaltplan"),

        # Hausanschluss
        "Wird eine Leistungserhöhung der Hausanschlusssicherung benötigt?": ("anschluss", "leistungserhoehung"),
        "Benötigte Hausanschlussicherung in Ampere (A)": ("anschluss", "anschlusssicherung"),

        # Zusatzabfragen Kundenanlage
        "[Zusatzabfrage 1]": ("zusatzabfragenKundenanlage",),
        "[Zusatzabfrage 2]": ("zusatzabfragenKundenanlage",),
        "[Zusatzabfrage x]": ("zusatzabfragenKundenanlage",),

        # Technische Betriebseinrichtung – PV
        "Leistung je PV-Modul gleichen Typs in kWp": ("technik", "pvModule", "leistungProModul"),
        "Anzahl der PV-Module gleichen Typs": ("technik", "pvModule", "anzahl"),
        "Ort der Anbringung": ("technik", "pvModule", "ort"),
        "Dachfläche": ("technik", "pvModule", "ort", "optionen", "Dachfläche"),
        "Freifläche": ("technik", "pvModule", "ort", "optionen", "Freifläche"),
        "Gesamtleistung aller PV-Module [in kWp]": ("technik", "pvModule", "gesamtleistung"),

        # Steuerung §9
        "Wählen Sie aus, ob der Netzbetreiber über eine bereits verbaute Steuerbox vor Ort zugreifen kann. (Steuerung nach § 9 EEG) (Dropdownfeld)": ("technik", "steuerung9"),
        "Auftrag zum Einbau neuer technischer Einrichtungen durch den grundzuständigen Messstellenbetreiber": ("technik", "steuerung9", "optionen", "Einbau durch grundzuständigen Messstellenbetreiber"),
        "Auftrag zum Einbau neuer technischer Einrichtungen durch den wettbewerblichen Messstellenbetreiber": ("technik", "steuerung9", "optionen", "Einbau durch wettbewerblichen Messstellenbetreiber"),

        # Wechselrichter Erzeugung
        "Wechselrichterhersteller": ("technik", "wechselrichter", "hersteller"),
        "Wechselrichtertyp": ("technik", "wechselrichter", "typ"),
        "Maximale Scheinleistung des Wechselrichters [in kVA]": ("technik", "wechselrichter", "scheinleistung"),
        "Maximale Wirkleistung des Wechselrichters [in kW]": ("technik", "wechselrichter", "wirkleistung"),
        "Wählen Sie aus, wie das System mit den Phasen der Anlage des Anschlussnutzers gekoppelt wird. (Mehrfachanmeldung möglich)": ("technik", "wechselrichter", "phasen"),
        "1-phasig": ("technik", "wechselrichter", "phasen", "optionen", "1-phasig"),
        "2-phasig": ("technik", "wechselrichter", "phasen", "optionen", "2-phasig"),
        "3-phasig": ("technik", "wechselrichter", "phasen", "optionen", "3-phasig"),
        "Drehstrom": ("technik", "wechselrichter", "phasen", "optionen", "Drehstrom"),
        "Einheitenzertifikat des Wechselrichters (Upload)": ("technik", "wechselrichter", "einheitenzertifikat"),
        "ZEREZ-Registernummer": ("technik", "wechselrichter", "zerez"),
        "Anzahl der Wechselrichter dieses Typs": ("technik", "wechselrichter", "anzahl"),
        "Summe der maximalen Scheinleistung [in kVA]": ("technik", "wechselrichter", "summeScheinleistung"),
        "Summe der maximalen Wirkleistung [in kW]": ("technik", "wechselrichter", "summeWirkleistung"),

        # PAV / Leistungsflussüberwachung
        "Wird eine Leistungsflussüberwachung eingesetzt?": ("technik", "pav", "eingesetzt"),
        "Hersteller der Leistungsflussüberwachung": ("technik", "pav", "hersteller"),
        "Typ der Leistungsflussüberwachung": ("technik", "pav", "typ"),
        "Zertifikat der Leistungsflussüberwachung (Upload)": ("technik", "pav", "zertifikat"),
        "ZEREZ-Registernummer": ("technik", "pav", "zerez"),
        "Gewünschte Einspeisewirkleistung (PAV, E-Wert) für den Netzanschlusspunkt (in kW)": ("technik", "pav", "einspeisewirkleistung"),

        # Vergütung
        "Geben Sie die gewünschte Veräußerungsform der Einspeisung an.": ("verguetung", "veraeusserungsform"),
        "Unentgeltliche Abnahme": ("verguetung", "veraeusserungsform", "optionen", "Unentgeltliche Abnahme"),
        "Geförderte Direktvermarktung (Marktprämier) nach § 20 EEG": ("verguetung", "veraeusserungsform", "optionen", "Geförderte Direktvermarktung"),
        "Einspeisevergütung nach § 21 Abs. 1 EEG": ("verguetung", "veraeusserungsform", "optionen", "Einspeisevergütung"),
        "Mieterstromzuschlag nach § 21 Abs. 3 EEG": ("verguetung", "veraeusserungsform", "optionen", "Mieterstromzuschlag"),
        "Sonstige Direktvermarktung nach § 21a EEG": ("verguetung", "veraeusserungsform", "optionen", "Sonstige Direktvermarktung"),
        "Ohne gesetzliche Vergütung": ("verguetung", "veraeusserungsform", "optionen", "Ohne gesetzliche Vergütung"),
        "Vergütungsrelevante Nachweise zur Veräußerungsform (Upload)": ("verguetung", "nachweiseVeraeusserung"),
        "Nachweise über die Fernsteuerbarkeit der Erzeugungsanlage (Upload)": ("verguetung", "fernsteuerbarkeitNachweis"),
        "Ist ein Mieterstromzuschlag § 21 Abs. 3 EEG erwünscht?": ("verguetung", "mieterstromZuschlag"),
        "Relevante Nachweise für den Mieterstromzuschlag (Uploadfeld)": ("verguetung", "nachweiseMieterstrom"),
        "Kontoinhaber": ("verguetung", "kontoinhaber"),
        "IBAN": ("verguetung", "iban"),
        "BIC": ("verguetung", "bic"),
        "Verwendungszweck": ("verguetung", "verwendungszweck"),
        "Sind Sie ein Unternehmen in Schwierigkeiten oder bestehen gegen Ihr Unternehmen offene Rückforderungsanschprüche? (Dropdpown)": ("verguetung", "rueckforderungStatus"),
        "Ich bin ein Unternehmen in Schwierigkeiten": ("verguetung", "rueckforderungStatus", "optionen", "Ich bin ein Unternehmen in Schwierigkeiten"),
        "Gegen mein Unternehmen bestehen offene Rückforderungsansprüche": ("verguetung", "rueckforderungStatus", "optionen", "Gegen mein Unternehmen bestehen offene Rückforderungsansprüche"),
        "Beides nicht zutreffend": ("verguetung", "rueckforderungStatus", "optionen", "Beides nicht zutreffend"),

        # Speicherbetrieb
        "Betriebsweise des Speichers (Dropdown)": ("speicher", "betriebsweise"),
        "Kein Bezug und keine Einspeisung": ("speicher", "betriebsweise", "optionen", "Kein Bezug und keine Einspeisung"),
        "Bezug und keine Einspeisung": ("speicher", "betriebsweise", "optionen", "Bezug und keine Einspeisung"),
        "Kein Bezug aber Einspeisung": ("speicher", "betriebsweise", "optionen", "Kein Bezug aber Einspeisung"),
        "Bezug und Einspeisung": ("speicher", "betriebsweise", "optionen", "Bezug und Einspeisung"),

        # Dokumente
        "Lageplan (Uploadfeld)": ("dateien", "lageplan"),
        "Grundrissskizze (Uploadfeld)": ("dateien", "grundriss"),
        "Sonstige Dateien (Uploadfeld)": ("dateien", "sonstige"),
        "Zusatzfeld Upload (Uploadfeld)": ("dateien", "zusatz"),
        "AGB (Checkbox)": ("agb",),
    }




    # === 2. JSON laden ===
    with open("EEG/assets/EEG.json", "r", encoding="utf-8") as f:
        data = json.load(f)

    # === 3. Excel laden ===
    df = pd.read_excel("EEG/assets/ANLAGE Abfrageformular Teilantrag Erzeugungsanlagen.xlsm", sheet_name="Abfrage optionale Felder", header=None)
    df.columns = [f"col_{i}" for i in range(df.shape[1])]
    relevant_rows = df.iloc[15:89, [2, 4, 5, 6]]
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
    eeg_obj = data["eeg"]
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
            update_json(eeg_obj, mapping[label], status_value, required_value, hint_text)
        else:
            print(f"⚠️ Kein Mapping für Label gefunden: {label}")

    # === 6. Speichern ===
    with open("output/EEG/EEG.json", "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
