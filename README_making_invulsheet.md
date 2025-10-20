# 🧭 README – Na het draaien van het script / After Running the Script
### Invulsheet afronden in Excel / Completing the Invulsheet in Excel

Deze handleiding beschrijft de **stappen in Excel ná het draaien van het script** om het *Invulsheet* volledig gebruiksklaar te maken.  
This guide describes the **manual Excel steps after the generation script** to finalize the Invulsheet.

> Formules staan **in NL én EN varianten onder elkaar**. Where applicable, **structured references** (table specifiers) are shown for both Dutch and English Excel.

---

## 📁 Bladen / Sheets
| NL Blad | EN Sheet | Doel / Purpose |
|---|---|---|
| **Invulsheet** | **Invulsheet** | Hoofdblad voor invoer / Main data entry |
| **Type Taxonomie** | **Type Taxonomie** | Object-/typeringstabel / Object/type taxonomy |
| **Attributen** | **Attributen** | Attributenhiërarchie / Attribute hierarchy |
| **variabelen** | **variabelen** | Hulplijsten / Helper lists |
| **domeinwaarden** | **domeinwaarden** | Domeinwaarden (enumeraties) / Domain values |
| *(optioneel)* **Toelichting invulsheet** | *(optional)* **Toelichting invulsheet** | Uitleg / Guidance |

**Tabellen / Tables**
- Op / On **Attributen** → **`Attribuuttabel`**
- Op / On **Type Taxonomie** → **`Taxonomie_tabel`**

**Structured reference specifiers (localized):**
- NL: `[#Alles]`, `[#Gegevens]`, `[#Kopteksten]`, `[#Totalen]`, `[#Deze rij]`
- EN: `[#All]`, `[#Data]`, `[#Headers]`, `[#Totals]`, `[#This Row]`

Bron: Microsoft Support (gestructureerde verwijzingen / structured references).

---

## ⚙️ Stap 0 — Voorbereiding / Step 0 — Preparation
- `Attributes.xlsx` → hiërarchietabel / attribute hierarchy  
- `taxonomie.xlsx` → object/type taxonomie / taxonomy  
- `domeinwaardes.xlsx` → domeinwaarden / domain values

---

## 🧩 Stap 1 — `Attributen` vullen / Fill `Attributen`
1) Kopieer data → **Attributen!A1** → **Invoegen → Tabel** (kopteksten aan) → hernoem tabel: **`Attribuuttabel`**.  
EN: Copy data → **Attributen!A1** → **Insert → Table** (headers on) → rename table: **`Attribuuttabel`**.

**Alternatief / Alternative (named range):**  
- NL & EN: `Attribuuttabel = Attributen!$A$1:$Z$500`

---

## 🧭 Stap 2 — `Type Taxonomie` vullen / Fill `Type Taxonomie`
1) Plak data → **Type Taxonomie!A1**.  
2) **Ctrl+G → Speciaal → Constanten** / **Ctrl+G → Special → Constants**.  
3) **Ctrl+F → Vervangen → “ZZ” → leeg** / **Ctrl+F → Replace → “ZZ” → empty**.  
4) **Invoegen → Tabel** / **Insert → Table**, hernoem naar **`Taxonomie_tabel`**.  
5) **Formules → Namen maken op basis van selectie (Bovenste rij)** / **Formulas → Create from Selection (Top row)**.

---

## 🧱 Stap 3 — `domeinwaarden` vullen / Fill `domeinwaarden`
1) Plak data → **domeinwaarden!A1** → **Ctrl+G → Speciaal/ Special → Constanten/Constants**.  
2) **Formules/ Formulas → Namen maken / Create from Selection (Bovenste rij / Top row)**.


## 🔢 Stap 4 — `variabelen` vullen / Fill `variabelen`
| Cel / Cell | Waarde / Value |
|---|---|
| A1 | Bewerkingscode |
| A2 | Nieuw |
| A3 | Verwijderen |
| A4 | Aanpassen |
| A5 | Instant laten |

**Named range / Benoemd bereik:**  
- `Bewerkingscode = variabelen!$A$2:$A$5`

---

## 🧾 Stap 5 — `Invulsheet` opzetten / Setup `Invulsheet`
Kolommen / Columns: `CAD-ID`, `GISIB-ID`, `Bewerkingscode`, `Objecttype`, `Type`, `Type gedetailleerd`, `Type extra gedetailleerd` (≥100 rijen / rows).

### 5.1 Keuzelijst “Bewerkingscode” / Dropdown “Bewerkingscode”
**NL (Gegevensvalidatie → Lijst → Bron):**
```
=Bewerkingscode
```
**EN (Data Validation → List → Source):**
```
=Bewerkingscode
```
*(of gebruik / or use direct cell range)*

### 5.2 Regel: precies één ID / Rule: exactly one ID (CAD or GISIB)
**NL (A2 & B2 → Gegevensvalidatie → Aangepast → Formule):**
```
=OF(EN($A2<>"";$B2="");EN($A2="";$B2<>""))
```
**EN (A2 & B2 → Data Validation → Custom → Formula):**
```
=OR(AND($A2<>"",$B2=""),AND($A2="",$B2<>""))
```

### 5.3 Hiërarchische keuzelijsten / Hierarchical dropdowns
**Objecttype (kolom D / column D)**  
NL:
```
=Objecttype
```
EN:
```
=Objecttype
```

**Type (kolom E / column E) — afhankelijk van D / dependent on D**  
NL:
```
=INDIRECT(SUBSTITUEREN(SUBSTITUEREN($D2;" ";"_");"-";"_"))
```
EN:
```
=INDIRECT(SUBSTITUTE(SUBSTITUTE($D2," ","_"),"-","_"))
```

**Type gedetailleerd (kolom F) — afhankelijk van E / dependent on E**  
NL:
```
=INDIRECT(SUBSTITUEREN(SUBSTITUEREN($E2;" ";"_");"-";"_"))
```
EN:
```
=INDIRECT(SUBSTITUTE(SUBSTITUTE($E2," ","_"),"-","_"))
```

**Type extra gedetailleerd (kolom G) — afhankelijk van F / dependent on F**  
NL:
```
=INDIRECT(SUBSTITUEREN(SUBSTITUEREN($F2;" ";"_");"-";"_"))
```
EN:
```
=INDIRECT(SUBSTITUTE(SUBSTITUTE($F2," ","_"),"-","_"))
```

### 5.4 Verborgen kolom “sleutel” / Hidden key column
**NL (I2):**
```
=TEKST.COMBINEREN(",";ONWAAR;$D2:$G2)
```
**EN (I2):**
```
=TEXTJOIN(",",FALSE,$D2:$G2)
```

Naar beneden doortrekken tot en met einde tabel (cell 101)

### 5.5 Optionele hulpkolom / Optional helper column (J2)
**NL:**
```
=ALS($D2<>"";ALS($D2="Terreindeel";"Groenobject";$D2);"")
```
**EN:**
```
=IF($D2<>"",IF($D2="Terreindeel","Groenobject",$D2),"")
```
Naar beneden doortrekken tot en met einde tabel (cell 101)
---

## 🧱 Stap 6 — Benoemde bereiken en tabellen / Named ranges & tables
Controleer / Verify:  
- `Attribuuttabel` op / on **Attributen**  
- `Taxonomie_tabel` op / on **Type Taxonomie**

Maak / Create (both localized specifiers shown):
```
Nederlands:
Header = Attribuuttabel[#Kopteksten]

Engels:
Header = Attribuuttabel[#Headers]

Attributen= Attribuuttabel[OTLProperty_prefLabel]

Attribuuttabel_n= Attributen!$A$2:$Z$500   ; exact data-bereik / exact data area
```

---

## 🎨 Stap 7 — Voorwaardelijke opmaak / Conditional Formatting (matrix)
Toepassen op matrix met kolomkoppen (attributen) en rijen (sleutel).  
Apply to the attribute matrix area (columns = attributes, rows = key).

**7.1 Groen / Green (applicable)**  
NL:
```
=ALS(INDEX(Attribuuttabel_n;VERGELIJKEN(K$1;Attributen;0);VERGELIJKEN($I2;Header;0))=1;WAAR;ONWAAR)
```
EN:
```
=IF(INDEX(Attribuuttabel_n,MATCH(K$1,Attributen,0),MATCH($I2,Header,0))=1,TRUE,FALSE)
```

**7.2 Grijs / Gray (not applicable)**  
NL:
```
=ALS(INDEX(Attribuuttabel_n;VERGELIJKEN(K$1;Attributen;0);VERGELIJKEN($I2;Header;0))=1;ONWAAR;WAAR)
```
EN:
```
=IF(INDEX(Attribuuttabel_n,MATCH(K$1,Attributen,0),MATCH($I2,Header,0))=1,FALSE,TRUE)
```

**7.3 Zwarte rand / Black border (enumerations)**  
NL:
```
=NIET(ALS(ISVERWIJZING(INDIRECT($J2 & "_" & SUBSTITUEREN(SUBSTITUEREN(K$1;" ";"_");"-";"_")));AANTALARG(INDIRECT($J2 & "_" & SUBSTITUEREN(SUBSTITUEREN(K$1;" ";"_");"-";"_")))=0;WAAR))
```
EN:
```
=NOT(IF(ISREF(INDIRECT($J2 & "_" & SUBSTITUTE(SUBSTITUTE(K$1," ","_"),"-","_"))),COUNTA(INDIRECT($J2 & "_" & SUBSTITUTE(SUBSTITUTE(K$1," ","_"),"-","_")))=0,TRUE))
```

---

## 🧩 Stap 8 — Lege fallback / Blank named range
**NL & EN:**
```
_Blank = variabelen!$B$1
```

---

## 🧮 Stap 9 — Datavalidatie matrix / Data validation (matrix)
**NL (bijv. K2):**
```
=ALS(ISVERWIJZING(INDIRECT($J2 & "_" & SUBSTITUEREN(SUBSTITUEREN(K$1;" ";"_");"-";"_")));INDIRECT($J2 & "_" & SUBSTITUEREN(SUBSTITUEREN(K$1;" ";"_");"-";"_"));_Blank)
```
**EN (e.g., K2):**
```
=IF(ISREF(INDIRECT($J2 & "_" & SUBSTITUTE(SUBSTITUTE(K$1," ","_"),"-","_"))),INDIRECT($J2 & "_" & SUBSTITUTE(SUBSTITUTE(K$1," ","_"),"-","_")),_Blank)
```

---

## 🗓️ Stap 10 — Datumopmaak / Date formatting
**NL:** Getalnotatie → Aangepast → `dd-mm-jjjj`  
**EN:** Number Format → Custom → `dd-mm-yyyy`

---

## ✅ Stap 11 — Checks
- Tabellen bestaan / Tables exist: `Taxonomie_tabel`, `Attribuuttabel`  
- Named ranges: `Header` (NL `#Kopteksten`, EN `#Headers`), `Attributen`, `Attribuuttabel_n`, `Bewerkingscode`, `_Blank`  
- Dropdowns hiërarchisch OK / Hierarchical dropdowns OK  
- Conditional formatting OK  
- Validatie één ID / One-ID validation OK

---

### 🧠 Functienaam-vertalingen / Function name pairs
- TEXTJOIN ⇄ **TEKST.COMBINEREN**  
- SUBSTITUTE ⇄ **SUBSTITUEREN**  
- ISREF ⇄ **ISVERWIJZING**  
- IF/AND/OR ⇄ **ALS/EN/OF**  
- MATCH ⇄ **VERGELIJKEN**  
- COUNTA ⇄ **AANTALARG**  
- INDEX ⇄ **INDEX**

### 🔖 Structured reference specifiers (EN ⇄ NL)
- `[#All]` ⇄ `[#Alles]`  
- `[#Data]` ⇄ `[#Gegevens]`  
- `[#Headers]` ⇄ `[#Kopteksten]`  
- `[#Totals]` ⇄ `[#Totalen]`  
- `[#This Row]` ⇄ `[#Deze rij]`

---

🎉 **Klaar / Done!** Het Invulsheet is volledig ingericht. / The Invulsheet is fully configured.
