# GISIB Werkboek Generator — README (NL)

Deze code genereert een Excel-werkboek vanuit OTL/asset-invulsheets (**Bomen**, **Beplantingen**, **Verhardingen**; **Water** kan later worden toegevoegd).
Het werkboek bevat o.a. bladen voor toelichting, invul, attribuuttabel, attributen-hiërarchie, domeinwaarden en variabelen.

Wanneer er een nieuwe asset wordt toevoegd, kijk hoe de gelaagdheid van typering is opgebouwd. Bomen is bijvoorbeeld anders opgebouwd nar beplantingen en verhardingen. Zie de manier van mapping in de mappings.py

## 📁 Benodigde invoer
Je hebt de volgende Excel-bestanden (invulsheets/bronnen) nodig:
- **OTL-catalogus** met minimaal de bladen:
    - `Objecttypen` (Taxonomie per objecttype)
    - `OTL Objecttypen Eigenschappen` (attribuutselectie per typering)
    - `OTL Enumeratietype` (enumeraties/domeinwaarden)
- **Bomen** (invulsheet/bron met OTL-kolommen)
- **Beplantingen** (idem)
- **Verhardingen** (idem)
- *(optioneel, later)* **Water**

> Let op: bladnamen zijn hoofdletter- en spatiegevoelig. Gebruik exact dezelfde namen.

## 🧱 Module-overzicht
- `assets.py` — definieert `Assets` (bijv. `BOOM`, `GROENOBJECT`, `VERHARDINGSOBJECT`, `TERREINDEEL`).
- `attributes.py` — maakt attribuuttabellen per asset:
  - `_attributes_per_typering(filepath, mapping, sheet_name)` (privé, 1 bestand)
  - `create_attributes_per_typering(filepaths, mappings, sheet_name, objecttypen_otl)` (meerdere bestanden combineren)
- `domain_values.py` — bouwt domeinwaarden (enumeraties) per asset naast elkaar:
  - `_domain_values(filepath, asset, sheet_name)` (privé, 1 bestand)
  - `create_domain_values(filepaths, assets_by_key, sheet_name, include)`
- `type_taxonomy.py` — maakt de hiërarchische **objecttype-tabel** (voorheen `df_sorted`) met tolerante kolomnaam-detectie en null-normalisatie.
- `output_sheet.py` — Excel-hulpfuncties en `build_workbook_minimal(...)` voor de outputbladen.
- `output_sheetnames.py` — `SHEETS_OUT` (Enum) met vaste bladnamen.
- `toelichting_invulsheet.py` — `build_intro_text(...)` voor de Toelichting.
- `mappings.py` — kolom-mappings per asset (welke kolommen worden gelezen/hernomen).
- `utils.py` — algemene hulpfuncties.
- `main.py` — voorbeeld/entrypoint dat de volledige pipeline aanroept.

## 🔧 Installatie
1. Gebruik **Python 3.10+**.
2. (Aanbevolen) Maak een virtual environment.
3. Installeer afhankelijkheden, bijv.:
   ```bash
   pip install pandas openpyxl numpy python-dotenv
   ```
   of via `requirements.txt` (als aanwezig):
   ```bash
   pip install -r requirements.txt
   ```

## ⚙️ Configuratie
Je kunt paden hardcoden of via `.env` aanleveren.

**Voorbeeld `.env`** (optioneel):
```env
FP_BOMEN=/pad/naar/bomen.xlsx
FP_BEPLANTING=/pad/naar/beplantingen.xlsx
FP_VERHARDING=/pad/naar/verhardingen.xlsx
```

**Voorbeeld in code:**
```python
filepaths = {
    "bomen": "/data/bomen.xlsx",
    "groen": "/data/beplantingen.xlsx",
    "grijs": "/data/verhardingen.xlsx",
    # "water": "/data/water.xlsx",   # later
}

from assets import Assets
assets_by_key = {
    "grijs": Assets.VERHARDINGSOBJECT.value,
    "groen": Assets.GROENOBJECT.value,
    "bomen": Assets.BOOM.value,
    # "water": Assets.WATER.value,   # later
}
```

## ▶️ Gebruik (typische flow)
```python
from assets import Assets
from attributes import create_attributes_per_typering
from domain_values import create_domain_values # bouwt objecttype_tabel
from output_sheet import build_workbook_minimal
from mappings import mapping_attrs, mapping_attrs_bomen  # pas aan naar jouw mappings

# 1) Paden
filepaths = {
    "bomen": "/data/bomen.xlsx",
    "groen": "/data/beplantingen.xlsx",
    "grijs": "/data/verhardingen.xlsx",
}

# 2) Attribuuttabel (samengevoegd over assets)
mappings = {
    "grijs": mapping_attrs,
    "groen": mapping_attrs,
    "bomen": mapping_attrs_bomen,
    # "water": mapping_attrs_water,  # later
}
attribuuttabel = create_attributes_per_typering(
    filepaths=filepaths, mappings=mappings, objecttypen_otl=None
)

# 3) Objecttype-hiërarchie (voor Attributen-blad)
objecttype_tabel = create_type_hierarchy_table(attribuuttabel)

# 4) Domeinwaarden
assets_by_key = {
    "grijs": Assets.VERHARDINGSOBJECT.value,
    "groen": Assets.GROENOBJECT.value,
    "bomen": Assets.BOOM.value,
}
domein_waarden = create_domain_values(
    filepaths=filepaths, assets_by_key=assets_by_key, include=["grijs","groen","bomen"]
)

# 5) Kolomvolgorde (optioneel)
columns = list(domein_waarden.notnull().sum().sort_values(ascending=True).index)

# 6) Workbook schrijven
output_path = "/data/werkboek.xlsx"
from toelichting_invulsheet import build_intro_text  # alleen nodig als je custom tekst wilt maken
build_workbook_minimal(
    objecttype_tabel=objecttype_tabel,
    attribuuttabel=attribuuttabel,
    domein_waarden=domein_waarden,
    columns=columns,
    output_path=output_path,
    fp_bomen=filepaths["bomen"],
    fp_beplanting=filepaths["groen"],
    fp_verharding=filepaths["grijs"],
)
print(f"Gereed: {output_path}")
```

## ➕ Water toevoegen (later)
1. Voeg `WATER` toe aan `Assets` in `assets.py`.
2. Voeg in `mappings.py` een mapping toe (bijv. `mapping_attrs_water`).
3. Neem `"water"` op in `filepaths`, `mappings` en `assets_by_key`.
4. Lever het Water-invulbestand (met juiste bladnamen) aan.

## ✅ Outputbladen
- **Toelichting** — gegenereerde tekst met paden/instructies.
- **Invulsheet** — lege tabel met kolommen:
  `CAD-ID`, `GISIB-ID`, `Bewerkingscode`, `Objecttype`, `Type`, `Type gedetailleerd`, `Type extra gedetailleerd`.
- **Tabel** — attribuuttabel (met index).
- **Attributen** — `objecttype_tabel`, kolomgewijs aflopend gesorteerd (NaN onderaan).
- **Domeinwaarden** — per asset naast elkaar geplaatste enumeraties; gevraagde kolommen of, als die ontbreken, alle beschikbare kolommen.
- **Variabelen** — kolom **Bewerkingscode** met waarden:
  `Nieuw`, `Verwijderen`, `Aanpassen`, `Instant laten`.

## 🛠️ Troubleshooting
- Controleer bladnamen exact (`OTL Objecttypen Eigenschappen`, `OTL Enumeratietype`).
- Lege **Domeinwaarden**? Dan misten waarschijnlijk kolommen — er is een fallback om **alle** kolommen te schrijven; check je `columns` parameter.
- Pas `mappings.py` aan als bronkolommen afwijken.
- Zet `invul_data_rows` (parameter van `build_workbook_minimal`) naar wens.

## 📜 Licentie
Interne/werkproject—kies zelf een licentie indien nodig.
