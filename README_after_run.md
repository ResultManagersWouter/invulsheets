# Vervolgstappen na het genereren van het werkboek (NL)

Dit document beschrijft **wat je in Excel doet ná het draaien van het script** om het invulsheet volledig gebruiksklaar te maken. De stappen zijn gebaseerd op jouw werkwijze en formules. Waar van toepassing staan **twee formulevarianten**: met **puntkomma (;)** voor Nederlandstalige Excel en met **komma (,)** voor Engelstalige Excel.

> **Bladen** die je in het eindresultaat gebruikt: `Invulsheet`, `Tabel`, `Attributen`, `Variabelen`, `Domeinwaarden` (en optioneel `Toelichting`).

---

## 0) Voorbereiding
- Zorg dat je de volgende bestanden hebt (gegenereerd of uit je pipeline):
  - `Attributes.xlsx` → hiërarchietabel (objecttype/typering; “taxonomie”).
  - `taxonomie.xlsx` → hetzelfde of gerelateerde tabel (indien apart gebruikt).
  - `domeinwaardes.xlsx` → per asset de enumeraties (OTL-domeinwaarden).
- Open het door het script gegenereerde Excel-werkboek **óf** begin in een lege Excel als je handmatig opbouwt.

> Het script schrijft normaliter de bladen al aan met tabellen. Deze handleiding toont ook hoe je dat **handmatig** kunt doen of bijwerken.

---

## 1) Bladen aanmaken (als je een lege Excel gebruikt)
Maak de volgende sheets aan (bladnamen exact overnemen):
1. `Invulsheet`
2. `Tabel`
3. `Attributen`
4. `Variabelen`
5. `Domeinwaarden`

---

## 2) Attributen-blad vullen (hiërarchietabel)
1. Open `Attributes.xlsx` en **kopieer de volledige tabel**.
2. Ga naar je doelbestand en plak de tabel op het blad **`Attributen`** (vanaf cel `A1`).
3. Selecteer de **hele geplakte tabel** (kop + data).
4. Tab **Insert** → **Table** (met “My table has headers” / “Tabel bevat kopteksten” aangevinkt).
5. Ga naar **Formulas** → **Name Manager**:
   - Zoek `Table1` (of de nieuwe tabelnaam).
   - Klik **Edit** en wijzig de naam in **`Attribuuttabel`**.

> Deze tabelnaam wordt later gebruikt in formules en validaties.

---

## 3) Tabel-blad vullen (taxonomie)
1. Open `taxonomie.xlsx` en **plak de inhoud** op het blad **`Tabel`** (vanaf `A1`).
2. Selecteer het volledige bereik met data.
3. Druk **Ctrl+G** → **Special** → **Constants** (laat alle vier opties aangevinkt) → **OK**.
4. Druk **Ctrl+F** → **Replace** om alle **`ZZ`** te vervangen door **lege waarden** (zoek `ZZ`, vervang door niets) → **Replace All** → **Close**.
5. **Formulas** → **Create from Selection** → vink **alleen** **Top row** aan → **OK**.

> Hiermee maak je voor elke kolomkop een Named Range (handig voor datavalidatie).

---

## 4) Domeinwaarden-blad vullen
1. Open `domeinwaardes.xlsx` en **kopieer** de volledige tabel naar **`Domeinwaarden`** (`A1`).
2. Selecteer het volledige gebied.
3. **Ctrl+G** → **Special** → **Constants** (alle opties aan) → **OK**.
4. **Formulas** → **Create from Selection** → vink **Top row** aan → **OK**.

> Dit maakt Named Ranges voor elke kolomnaam (bijv. `Boom_OTLPropertyX`). Deze gebruik je met `INDIRECT(...)` in validaties.

---

## 5) Variabelen-blad vullen
In `Variabelen` zet je de bewerkingscodes:

- **A1**: `Bewerkingscode`
- **A2**: `Nieuw`
- **A3**: `Verwijderen`
- **A4**: `Aanpassen`
- **A5**: `Instant laten`

Maak er desgewenst een **Table** van en/of gebruik **Formulas → Create from Selection** (Top row) om een Named Range `Bewerkingscode` aan te maken.

---

## 6) Invulsheet opzetten
Kolommen (minstens) in deze volgorde aanmaken op `Invulsheet`:

1. `CAD-ID`
2. `GISIB-ID`
3. `Bewerkingscode`
4. `Objecttype`
5. `Type`
6. `Type gedetailleerd`
7. `Type extra gedetailleerd`

> Zorg dat er **100 datarijen** zijn (of naar wens).

### 6.1 Datavalidatie “Bewerkingscode”
- Ga naar **C2** → **Data** → **Data Validation** → **List** → **Source**: `=Bewerkingscode`
- Kopieer validatie naar beneden t/m rij 100.

### 6.2 Regel: precies één ID invullen (CAD of GISIB)
- Selecteer **A2** → **Data Validation** → **Allow: Custom** → **Formula**:

**NL Excel (;)**
```
=OF(EN($A2<>"";$B2="");EN($A2="";$B2<>""))
```

**EN Excel (,)**
```
=OR(AND($A2<>"",$B2=""),AND($A2="",$B2<>""))
```

- **Error Alert**: Titel bijv. “Maximaal 1 ID”, tekst: “Vul maximaal 1 van de 2 ID’s (CAD of GISIB).”
- Kopieer de validatie van **A2** naar **B2**, en dan naar beneden t/m rij 100.

### 6.3 Keuzelijsten voor typering (afhankelijk van hiërarchie)
- **Objecttype (kolom D)**: Data Validation → **List** → **Source**: `=Objecttype`
- Kopieer naar beneden t/m rij 100.

- **Type (kolom E)**: Data Validation → **List** → **Source** met `INDIRECT` op kolom D:

**NL Excel (;)**
```
=INDIRECT(SUBSTITUTE(SUBSTITUTE($D2;" ";"_");"-";"_"))
```

**EN Excel (,)**
```
=INDIRECT(SUBSTITUTE(SUBSTITUTE($D2," ","_"),"-","_"))
```

- **Type gedetailleerd (kolom F)**: verwijst naar kolom **E**

**NL**
```
=INDIRECT(SUBSTITUTE(SUBSTITUTE($E2;" ";"_");"-";"_"))
```

**EN**
```
=INDIRECT(SUBSTITUTE(SUBSTITUTE($E2," ","_"),"-","_"))
```

- **Type extra gedetailleerd (kolom G)**: verwijst naar kolom **F**

**NL**
```
=INDIRECT(SUBSTITUTE(SUBSTITUTE($F2;" ";"_");"-";"_"))
```

**EN**
```
=INDIRECT(SUBSTITUTE(SUBSTITUTE($F2," ","_"),"-","_"))
```

Kopieer deze validaties naar beneden t/m rij 100.

### 6.4 “Sleutel” kolom (concatenatie)
- Maak **kolom I** met kop **`sleutel`**.
- In **I2**:

**NL**
```
=TEKST.SAMENVOEGEN(",";ONWAAR;$D2:$G2)
```
of (nieuwe Excel):
```
=TEXTJOIN(",";ONWAAR;$D2:$G2)
```

**EN**
```
=TEXTJOIN(",",FALSE,$D2:$G2)
```

- Kopieer omlaag t/m rij 100.
- Verberg kolom **I** (en later kolom **H** als je die gebruikt voor hulpwaarden).

### 6.5 Optionele hulp-kolom J (asset-afleiding)
- **J** kop: (bijv. leeg of `asset_derived`)
- **J2**:

**NL**
```
=ALS($D2<>"";ALS($D2="Terreindeel";"Groenobject";$D2);"")
```

**EN**
```
=IF($D2<>"",IF($D2="Terreindeel","Groenobject",$D2),"")
```

- Kopieer naar **J100** en **verberg kolom J**.

---

## 7) Controleer naamgeving van de tabel op blad “Tabel”
- **Formulas → Name Manager**: controleer dat de geplakte tabel **`Attribuuttabel`** heet.
- In **K1** (bijv. op `Tabel`) kun je voor controle plaatsen:
```
=TRANSPOSE(Attribuuttabel[OTLProperty_prefLabel])
```

Maak drie Named Ranges via **Name Manager → New**:
1. **Header** = `Attribuuttabel[#Headers]`
2. **Attributen** = `Attribuuttabel[OTLProperty_prefLabel]`
3. **Attribuuttabel_n** = (losse **Named Range** die exact het **data-bereik** van de tabel dekt; nodig voor INDEX/MATCH in CF)

---

## 8) Conditional Formatting (Attributenmatrix)
Stel CF in op het matrixbereik waar **rijen** naar `sleutel` verwijzen en **kolommen** de attributen zijn.

### 8.1 Groen als attribuut *van toepassing is*
Formule (selecteer het matrixbereik eerst):

**NL**
```
=ALS(INDEX(Attribuuttabel_n;VERGELIJKEN(K$1;Attributen;0);VERGELIJKEN($I2;Header;0))=1;WAAR;ONWAAR)
```

**EN**
```
=IF(INDEX(Attribuuttabel_n,MATCH(K$1,Attributen,0),MATCH($I2,Header,0))=1,TRUE,FALSE)
```

Kies **Green fill**.

### 8.2 Grijs als attribuut *niet van toepassing is*
**NL**
```
=ALS(INDEX(Attribuuttabel_n;VERGELIJKEN(K$1;Attributen;0);VERGELIJKEN($I2;Header;0))=1;ONWAAR;WAAR)
```

**EN**
```
=IF(INDEX(Attribuuttabel_n,MATCH(K$1,Attributen,0),MATCH($I2,Header,0))=1,FALSE,TRUE)
```

Kies **Gray fill**.

### 8.3 Zwarte rand bij keuzelijsten (enumeraties)
Formule (bijv. te plaatsen vanaf **K2** over de matrix; gebruik **Manage Rules** om “Applies to” te verbreden):

**NL**
```
=NIET(ALS(ISVERW(INDIRECT($J2 & "_" & SUBSTITUEREN(SUBSTITUEREN(K$1;" ";"_");"-";"_")));AANTALARG(INDIRECT($J2 & "_" & SUBSTITUEREN(SUBSTITUEREN(K$1;" ";"_");"-";"_")))=0;WAAR))
```

**EN**
```
=NOT(IF(ISREF(INDIRECT($J2 & "_" & SUBSTITUTE(SUBSTITUTE(K$1," ","_"),"-","_"))),COUNTA(INDIRECT($J2 & "_" & SUBSTITUTE(SUBSTITUTE(K$1," ","_"),"-","_")))=0,TRUE))
```

**Formatting**: alleen **zwarte rand**.

> Gebruik **Home → Conditional Formatting → Manage Rules** om het bereik (Applies to) over **alle kolommen** en t/m rij **100** door te trekken.

---

## 9) Named Range `_Blank`
Voor de data-validatie in de attributenmatrix is een lege fallback handig.

1. Ga naar blad **`Variabelen`**.
2. Kies een **lege cel** (bijv. **B1**).
3. **Formulas → Name Manager → New**:
   - Name: `_Blank`
   - Refers to: `=Variabelen!$B$1`

---

## 10) Datavalidatie in de attributenmatrix (op basis van domeinwaarden)
In de **eerste cel** van de attributenmatrix (bijv. K2) stel je in:

**Data Validation → List → Source**

**NL**
```
=ALS(ISVERW(INDIRECT($J2 & "_" & SUBSTITUEREN(SUBSTITUEREN(K$1;" ";"_");"-";"_")));INDIRECT($J2 & "_" & SUBSTITUEREN(SUBSTITUEREN(K$1;" ";"_");"-";"_"));_Blank)
```

**EN**
```
=IF(ISREF(INDIRECT($J2 & "_" & SUBSTITUTE(SUBSTITUTE(K$1," ","_"),"-","_"))),INDIRECT($J2 & "_" & SUBSTITUTE(SUBSTITUTE(K$1," ","_"),"-","_")),_Blank)
```

Trek **horizontaal** door en dan **naar beneden** t/m rij **101**.

---

## 11) Datumvelden opmaken
Selecteer de datumkolommen (bijv. `Begin garantieperiode`, `Eind garantieperiode`, `Objecteindtijd`, `Opleverdatum`) en stel **Number Format** in op **dd/mm/yyyy** (of **dd-mm-jjjj**).

---

## 12) Laatste checks
- `Attribuuttabel` bestaat in **Name Manager**.
- `Bewerkingscode` Named Range bestaat.
- `Header`, `Attributen`, `Attribuuttabel_n`, `_Blank` bestaan.
- Validaties werken (dropdowns verschijnen, hiërarchische afhankelijkheid is oké).
- CF kleurt groen/grijs correct; zwarte randen bij attributen met enumeraties.

---

**Klaar!** Je invulsheet is nu volledig ingericht en klaar voor gebruik.
