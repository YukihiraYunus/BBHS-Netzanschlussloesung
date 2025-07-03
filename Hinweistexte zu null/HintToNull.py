import json

def clean_placeholder_hinweistexte(obj):
    if isinstance(obj, dict):
        for key, value in obj.items():
            if key == "hinweistext" and isinstance(value, str) and "Hinweistext" in value:
                obj[key] = None
            else:
                clean_placeholder_hinweistexte(value)
    elif isinstance(obj, list):
        for item in obj:
            clean_placeholder_hinweistexte(item)

# Lade die JSON-Datei
with open("Hinweistexte zu null/OutletGebäude.json", "r", encoding="utf-8") as f:
    data = json.load(f)

# Bearbeite die Daten
clean_placeholder_hinweistexte(data)

# Speichere die bereinigte JSON-Datei
with open("Hinweistexte zu null/JSON_bereinigt.json", "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

print("Nur Platzhalter-Hinweistexte wurden auf null gesetzt (z. B. 'Hinweistext xyz').")
