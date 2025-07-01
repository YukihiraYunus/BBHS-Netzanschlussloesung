import json

def set_hinweistext_to_null(obj):
    if isinstance(obj, dict):
        for key, value in obj.items():
            if key == "hinweistext":
                obj[key] = None
            else:
                set_hinweistext_to_null(value)
    elif isinstance(obj, list):
        for item in obj:
            set_hinweistext_to_null(item)

# Lade die JSON-Datei
with open("Hinweistexte zu null/EEG.json", "r", encoding="utf-8") as f:
    data = json.load(f)

# Bearbeite die Daten
set_hinweistext_to_null(data)

# Speichere die geänderte JSON-Datei
with open("Hinweistexte zu null/EEG_bereinigt.json", "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

print("Alle Hinweistexte wurden auf null gesetzt und die Datei wurde als 'EEG_bereinigt.json' gespeichert.")
