import pandas as pd
def _attributes_per_typering(filepath: str, mapping: dict, sheet_name: str = "OTL Objecttypen Eigenschappen") -> pd.DataFrame:
    """Read and filter attributes for a single Excel file."""
    xls = pd.ExcelFile(filepath)
    assert sheet_name in xls.sheet_names, f"{sheet_name} not found in {filepath}"

    columns = list(mapping.keys())
    df = (
        pd.read_excel(filepath, sheet_name=sheet_name, usecols=columns)
        .rename(columns=mapping)
        .loc[lambda d: d.Demarcatie == "Aannemer"]
        .loc[lambda d: d.Objecttype.notnull()]
        .loc[lambda d: d.OTLProperty_prefLabel != "Heeft document"]
        .drop(columns=["Demarcatie"])
    )
    return df


def create_attributes_per_typering(filepaths: dict, mappings: dict, sheet_name: str, objecttypen_otl=None) -> pd.DataFrame:
    """
    Combines attributes per typering from multiple Excel files.

    Parameters
    ----------
    filepaths : dict
        Dictionary with keys like 'grijs', 'groen', 'bomen', etc.
    mappings : dict
        Dictionary of mappings per key in filepaths (e.g. {'bomen': mapping_attrs_bomen, 'groen': mapping_attrs, ...}).
    sheet_name : str, optional
        Excel sheet to read, by default "OTL Objecttypen Eigenschappen".
    objecttypen_otl : list, optional
        If provided, filters the resulting DataFrame to only include those Objecttypes.
    """
    dfs = [
        _attributes_per_typering(filepath=fp, mapping=mappings[key], sheet_name=sheet_name)
        for key, fp in filepaths.items()
        if key in mappings  # only process if a mapping is defined
    ]

    combined = pd.concat(dfs, ignore_index=True)
    if objecttypen_otl is not None:
        combined = combined.loc[lambda d: d.Objecttype.isin(objecttypen_otl)]

    cols = [
        "Objecttype",
        "Type",
        "Type gedetailleerd",
        "Type extra gedetailleerd",
    ]

    combined = combined.assign(
        sleutel=lambda d: d[cols].agg(lambda s: ",".join(s.astype(str)), axis=1)
    )
    # remove literal 'nan' and collapse extra commas, trim edges
    combined["sleutel"] = (
        combined["sleutel"]
        .str.replace("nan", "", regex=False)
    )
    tabel = pd.crosstab(combined.OTLProperty_prefLabel, combined.sleutel,
                        dropna=False)

    return tabel
