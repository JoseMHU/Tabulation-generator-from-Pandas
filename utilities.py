import numpy as np
import pandas as pd
from settings.config import JsonFile


category_order = dict(JsonFile(path="settings/config_values.json").file)

for var in category_order.keys():
    category_order[var].insert(0, "Total")
    category_order[var].insert(0, "Var")


def remove_duplicates(x: list) -> list:
    return np.unique(x).tolist()


def interpret_instruction(instrution: str, question: str) -> tuple:
    result = instrution.split(",")
    if result[0] == "${multiply}":
        result.append("*")
        result[0] = question
        
    return tuple(result)


def calculate_percentages(structure: dict, df: pd.DataFrame, db: pd.DataFrame,
                          table_type: int, factor: any, split_df_by=None) -> pd.DataFrame:
    
    def calculate_factor(_filter, _factor, _db, axis=0):
        if type(_filter) is not type(None):
            if type(_factor) is tuple:
                if _factor[2] == "*":
                    multiplication = _db.loc[_filter, _factor[0]] * _db.loc[_filter, _factor[1]]
                    result = multiplication.sum(axis=axis)
                else:
                    raise ValueError("Unrecognized operation in SUM of Var column")
            else:
                result = _db.loc[_filter, _factor].sum(axis=axis)
            return result
        else:
            if type(_factor) is tuple:
                if _factor[2] == "*":
                    multiplication = _db[_factor[0]] * _db[_factor[1]]
                    result = multiplication.sum(axis=axis)
                else:
                    raise ValueError("Unrecognized operation in SUM of Var column")
            elif _factor == "df":
                result = _db.sum(axis=axis)
            else:
                result = _db[_factor].sum(axis=axis)
            return result

    if table_type == 1:
        columns_to_convert = [x for x in structure if x != "Var"]
        df_selected = df[columns_to_convert]
        
        if split_df_by:
            _filter = (db[split_df_by[0]] == split_df_by[1])
            divider = calculate_factor(_filter, factor, db)
        else:
            divider = calculate_factor(None, factor, db)
        df[columns_to_convert] = df_selected.div(divider, axis=0)
        
    elif table_type == 2:
        columns_to_convert = [x for x in structure if x != "Var"]
        df_selected = df[columns_to_convert]
        divider = calculate_factor(None, "df", df_selected)
        df[columns_to_convert] = df_selected.div(divider, axis=1)
        
    else:
        columns_to_convert = [x for x in structure if x not in ("Var", "Total")]
        df_selected = df[columns_to_convert]
        divider = calculate_factor(None, "df", df_selected, 1)
        df[columns_to_convert] = df_selected.div(divider, axis=0)
        df["Total"] = df["Total"].div(df["Total"], axis=0)
    
    return df
