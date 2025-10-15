import re
def extract_otl_version(fp: str) -> str | None:
    m = re.search(r'otl[-\s:]?(\d+\.\d+\.\d+)', fp, flags=re.IGNORECASE)
    return m.group(1) if m else None

def build_intro_text(fp_bomen: str, fp_beplanting: str, fp_verharding: str) -> str:
    ver_bomen = extract_otl_version(fp_bomen) or "?"
    ver_bepl  = extract_otl_version(fp_beplanting) or "?"
    ver_verh  = extract_otl_version(fp_verharding) or "?"
    tekst = f"""Toelichting invulsheet

Op basis van de OTL-exports van verschillende assets is dit invulsheet opgesteld. Het betreft de volgende assettypen:

- Beplantingen (OTL {ver_bepl}) - groenobjecten en terreindelen

- Verhardingen (OTL {ver_verh})

- Bomen (OTL {ver_bomen})

Per object kan worden aangegeven om wat voor wijziging het gaat:
Nieuw, Instanthouden, Aanpassen of Verwijderen.

Vul daarbij in om welk object het gaat:

Bij een object uit de CAD-tekening: vul het CAD-ID in.

Bij een bestaand GISIB-object: vul het GISIB-ID in.
Dit ID is te vinden via:
https://data.amsterdam.nl/data/geozoek/?center=52.3735159%2C4.8892885&lagen=oor-groenobjecten%7Coor-verhardingen%7Coor-terreindelen&legenda=true&modus=kaart&term=Objecten+Openbare+Ruimte&zoom=11

Daarna geef je, indien van toepassing, het objecttype, type, type gedetailleerd en extra gedetailleerd op.
Als een waarde niet van toepassing is, laat je het veld leeg.

De aannemer vult vervolgens alle groene cellen in.
Deze groene cellen zijn de gegevens die volgens de OTL moeten worden aangeleverd.
De grijze cellen hoeven niet te worden ingevuld.

Als een groen veld een zwarte rand heeft, betreft het een keuzelijst:
klik één van de beschikbare waarden aan.
De overige velden zijn vrije invulvelden.

Datums worden altijd ingevuld in het formaat dag-maand-jaar.
"""
    return tekst