import pandas as pd
from collections import defaultdict
import math
def _create_overview_taxonomy(filepath:str, sheet_name:str,objecttypen_otl : list)->pd.DataFrame:
    assert sheet_name in pd.ExcelFile(filepath).sheet_names
    df = (
        pd.read_excel(filepath,
                      sheet_name =sheet_name)
    )
    print(set(objecttypen_otl) & set(df.ExternalObjecttype_prefLabel.tolist()))
    assert len(set(objecttypen_otl) & set(df.ExternalObjecttype_prefLabel.tolist())) >= 1
    df = (df
        .loc[:,["ExternalObjecttype_prefLabel","ExternalValue_1_prefLabel","ExternalValue_2_prefLabel","ExternalValue_3_prefLabel"]]
        .loc[lambda df: df.ExternalObjecttype_prefLabel.isin(objecttypen_otl)]
        .dropna(how="all")
        .rename(columns={
            "ExternalObjecttype_prefLabel":"Objecttype",
            "ExternalValue_1_prefLabel":"Type",
            "ExternalValue_2_prefLabel":"Type gedetailleerd",
            "ExternalValue_3_prefLabel":"Type extra gedetailleerd",
            }
               )
        .fillna("ZZ")
    )
    return df

def create_overview_taxonomy(filepaths : list[str],sheet_name:str,objecttypen_otl:list):
    objecten_overview = pd.concat([
        _create_overview_taxonomy(filepath=fp, sheet_name=sheet_name, objecttypen_otl=objecttypen_otl)
        for fp in filepaths.values()
    ])
    return objecten_overview




# ---- private helpers ---------
def _first_present(df: pd.DataFrame, names: list[str]) -> str | None:
    """Return the first column name from `names` that exists in df, else None."""
    for n in names:
        if n in df.columns:
            return n
    return None

def _seen_contains(seen: list, value) -> bool:
    """like `value in seen`, but treats NaN as equal to NaN (so counted once)."""
    for v in seen:
        if (isinstance(v, float) and math.isnan(v)) and (isinstance(value, float) and math.isnan(value)):
            return True
        if v == value:
            return True
    return False

def _rectangularize(flat: dict) -> pd.DataFrame:
    """Pad lists in dict to same length and return a DataFrame."""
    max_len = max((len(v) for v in flat.values()), default=0)
    rect = {k: (list(v) + [None] * (max_len - len(v))) for k, v in flat.items()}
    return pd.DataFrame(rect)

# --------- public function ---------
def create_type_table(objecten_overview: pd.DataFrame) -> pd.DataFrame:
    """
    Recreates the original pipeline:
      - normalizes stringy nulls ("NULL", "None", "NaN") to "ZZ" (same as your code)
      - resolves level columns with typo-tolerance
      - builds parent->child lists (including a single null/NaN child)
      - includes a master list for the top level
      - rectangularizes and orders columns by null counts (descending)
    Returns the sorted DataFrame (same shape/contents as your df_sorted).
    """
    # --- Clean input (preserve original behavior: map stringy nulls to "ZZ") ---
    df = objecten_overview.copy()
    df = df.replace({r'^\s*(NULL|None|NaN)\s*$': "ZZ"}, regex=True)

    # --- Resolve level columns ---
    col_obj  = _first_present(df, ["Objecttype"])
    col_typ  = _first_present(df, ["Type"])
    col_det  = _first_present(df, ["Type gedetailleerd", "Type gedetailleeerd"])  # typo-tolerant
    col_xdet = _first_present(df, ["Type extra gedetailleerd"])                   # optional

    levels = [c for c in [col_obj, col_typ, col_det, col_xdet] if c is not None]
    if not levels:
        raise ValueError("No expected columns found.")

    # --- Build mappings (include one null/NaN child) ---
    flat: dict = defaultdict(list)
    for _, row in df.iterrows():
        for i in range(len(levels) - 1):
            parent_col = levels[i]
            child_col  = levels[i + 1]
            parent_val = row[parent_col]
            child_val  = row[child_col]

            if parent_val is not None:
                flat.setdefault(parent_val, [])
                if not _seen_contains(flat[parent_val], child_val):
                    flat[parent_val].append(child_val)

    # Master list for the first column (matches original: excludes None, not NaN)
    top_col = levels[0]
    flat[top_col] = [x for x in pd.unique(df[top_col]) if x is not None]

    # --- Rectangularize and sort columns by null count (descending) ---
    df_out = _rectangularize(flat)
    null_counts = df_out.isna().sum().sort_values(ascending=False).index
    df_sorted = df_out.replace({None: pd.NA}).loc[:, list(null_counts)]

    return df_sorted
