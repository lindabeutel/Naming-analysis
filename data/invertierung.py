import json
from collections import defaultdict

# Pfade ggf. anpassen
source_path = "lemma_normalization.json"
target_path = "lemma_variants.json"

# Quelle laden
with open(source_path, "r", encoding="utf-8") as f:
    normalization = json.load(f)

# Invertierung vornehmen
inverted = defaultdict(list)
for form, lemma in normalization.items():
    inverted[lemma].append(form)

# Sortierung innerhalb der Listen (optional, aber hilfreich)
for lemma in inverted:
    inverted[lemma] = sorted(set(inverted[lemma]))

# Speichern
with open(target_path, "w", encoding="utf-8") as f:
    json.dump(dict(inverted), f, ensure_ascii=False, indent=2)

print(f"✅ Lemma-Varianten gespeichert in: {target_path}")
