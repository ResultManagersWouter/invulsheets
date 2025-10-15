import pandas as pd
def print_sheet_names(filepaths: list[str]):
    for fp in filepaths.values():
        print(fp, "\n", pd.ExcelFile(fp).sheet_names)