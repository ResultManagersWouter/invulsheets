import os


from dotenv import load_dotenv
from assets import Assets
from mappings import maps
from domain_values import create_domain_values
from attributes import create_attributes_per_typering
from type_taxonomy import create_overview_taxonomy,create_type_table
from utils import print_sheet_names
from global_vars import vandaag, ROWS_OUPUT,OUTPUT_PATH

from output_sheet import build_workbook_minimal

load_dotenv()

objecttypen_otl = [o.value for o in Assets]
# dictonary based on filepaths.
filepaths = {
    "bomen": os.environ.get("FP_BOMEN"),
    "groen": os.environ.get("FP_BEPLANTING"),
    "grijs": os.environ.get("FP_VERHARDING"),
    # "water" :fp_water# df_water
}

assets_by_key = {
    "grijs": Assets.VERHARDINGSOBJECT.value,
    "groen": Assets.GROENOBJECT.value,
    "bomen": Assets.BOOM.value,
    # "water": Assets.WATER.value,  # wanneer je WATER toevoegt aan de Enum
}

# Press the green button in the gutter to run the script.
if __name__ == "__main__":
    print_sheet_names(filepaths=filepaths)

    objecttypen_overview = create_overview_taxonomy(
        filepaths=filepaths, sheet_name="Objecttypen", objecttypen_otl=objecttypen_otl
    )
    objecttype_tabel = create_type_table(objecten_overview=objecttypen_overview)

    domein_waarden = create_domain_values(
        filepaths=filepaths,
        assets_by_key=assets_by_key,
        include=["grijs", "groen", "bomen"],
        sheet_name="OTL Enumeratietype",  # voeg "water" toe zodra beschikbaar
    )

    attributes_per_typering = create_attributes_per_typering(
        filepaths=filepaths,
        mappings=maps,
        sheet_name="OTL Objecttypen Eigenschappen",
        objecttypen_otl=objecttypen_otl,
    )
    cols = [
        "Objecttype",
        "Type",
        "Type gedetailleerd",
        "Type extra gedetailleerd",
    ]

    # build_workbook_minimal(
    # objecttype_tabel = objecttype_tabel,
    # attribuuttabel = attributes_per_typering,
    # domein_waarden = domein_waarden,
    # columns  =cols,
    # output_path = OUTPUT_PATH",
    # fp_bomen = filepaths["bomen"],
    # fp_beplanting  = filepaths["groen"],
    # fp_verharding =  filepaths["grijs"],
    # invul_data_rows= ROWS_OUPUT)