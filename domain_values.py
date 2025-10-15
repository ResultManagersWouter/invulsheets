from enum import Enum
import pandas as pd

def _domain_values(filepath: str, asset: str, sheet_name: str) -> pd.DataFrame:
    """
    Private helper: lees en pivoteer domeinwaardes voor één Excel-bestand.
    """
    otl_data = pd.read_excel(filepath, sheet_name=sheet_name)

    tmp = (
        otl_data.copy()
        .assign(
            OtlEnumerationValueName=lambda d: d["OtlEnumerationValueName"].apply(
                lambda x: x.strip() if isinstance(x, str) else x
            )
        )
        .dropna(subset=["OtlEnumerationValueName"])
        .query("OtlEnumerationValueName != ''")
        .drop_duplicates(subset=["OtlPropertyName", "OtlEnumerationValueName"])
    )

    # rij-index per attribuut
    tmp["row"] = tmp.groupby("OtlPropertyName").cumcount()

    # wide pivot: attributen als kolommen
    wide = (
        tmp.pivot(index="row", columns="OtlPropertyName", values="OtlEnumerationValueName")
        .sort_index(axis=1)
        .reset_index(drop=True)
    )

    # behoud alleen kolommen met >1 niet-null waarde
    keep_cols = (
        wide.notnull().sum().sort_values(ascending=True).reset_index()
        .loc[lambda d: d.iloc[:, 1] > 1, "OtlPropertyName"]
        .tolist()
    )
    wide = wide.loc[:, keep_cols]

    # prefix met asset en normaliseer kolomnamen
    wide.columns = [f'{asset}_{c.replace(" ", "_").replace("-", "_")}' for c in wide.columns]

    # optioneel: kolommen met meeste missings naar rechts
    columns_order = list(wide.isnull().sum().sort_values(ascending=False).index)
    wide = wide.loc[:, columns_order]

    return wide


def create_domain_values(
    filepaths: dict,
    assets_by_key: dict,
    sheet_name: str,
    include: list | None = None,
) -> pd.DataFrame:
    """
    Combineer domeinwaardes van meerdere Excel-bestanden (axis=1) en orden kolommen.

    Parameters
    ----------
    filepaths : dict
        Bijv. {'grijs': '/pad/grijs.xlsx', 'groen': '...', 'bomen': '...', 'water': '...'}
    assets_by_key : dict
        Bijv. {'grijs': Assets.VERHARDINGSOBJECT.value, 'groen': Assets.GROENOBJECT.value, ...}
    sheet_name : str
        Excel-werkblad met enumeraties.
    include : list | None
        Beperk tot een subset keys (bijv. ['grijs','groen','bomen','water']).
    """
    keys = include if include is not None else list(filepaths.keys())

    parts = [
        _domain_values(filepath=filepaths[k], asset=assets_by_key[k], sheet_name=sheet_name)
        for k in keys
        if k in filepaths and k in assets_by_key
    ]

    if not parts:
        return pd.DataFrame()

    combined = pd.concat(parts, axis=1)

    # finale kolomordering: oplopend op aantal niet-null waarden
    cols = combined.notnull().sum().sort_values(ascending=True).index.tolist()
    combined = combined.loc[:, cols]

    return combined
