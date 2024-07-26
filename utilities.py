import numpy as np
import pandas as pd
import scipy.stats as stats
from settings.config import JsonFile


category_order = dict(JsonFile(path="settings/config_values.json").file)

for var in category_order.keys():
    category_order[var].insert(0, "Var")


def remove_duplicates(x: list) -> list:
    return np.unique(x).tolist()


def interpret_instruction(instruction: str) -> tuple:
    result = instruction.split(",")
    if result[1] == "${multiply}":
        result.append("*")

    return tuple(result)


def calculate_percentages(structure: dict | None, df: pd.DataFrame, table_type: int) -> pd.DataFrame:
    if table_type == 1:
        columns_to_convert = [x for x in structure if x != "Var"]
        df_selected = df[columns_to_convert]
        divider = df["Total"].sum()
        df[columns_to_convert] = df_selected.div(divider, axis=0)
        df["Total"] = df["Total"] / divider

    elif table_type == 2:
        columns_to_convert = [x for x in structure if x != "Var"]
        df_selected = df[columns_to_convert]
        divider = df_selected.sum()
        df[columns_to_convert] = df_selected.div(divider, axis=1)
        df["Total"] = df["Total"] / df["Total"].sum(axis=0)

    elif table_type == 3:
        columns_to_convert = [x for x in structure if x not in ("Var", "Total")]
        df_selected = df[columns_to_convert]
        divider = df_selected.sum(axis=1)
        df[columns_to_convert] = df_selected.div(divider, axis=0)
        df["Total"] = df["Total"].div(df["Total"], axis=0)

    else:
        df.drop([x for x in df.columns.tolist() if x not in ("Var", "Total")], axis=1, inplace=True)
        df["%"] = df["Total"] / df["Total"].sum(axis=0)

    return df


def chi2_test(df: pd.DataFrame, alpha=0.05) -> tuple:
    df.drop(columns=["Total", "Var"], inplace=True)
    df.drop(df.index[-1], inplace=True)

    observed = np.array(df)

    try:
        _, p_val, _, _ = stats.chi2_contingency(observed)
    except ValueError:
        return (False, "The internally computed table of expected frequencies has a zero element.",
                "Null")

    if p_val < alpha:
        return (True,
                "* The Chi-squeare statistic is significant at the .05 level.",
                f"p-value: {np.round(p_val, 3)}")
    else:
        return (False,
                "* The Chi-squeare statistic is not significant at the .05 level.",
                f"p-value: {np.round(p_val, 3)}")

    
def insert_total_row(df: pd.DataFrame) -> pd.DataFrame:
    _sum = df.drop(columns="Var").sum()
    
    row_total = pd.DataFrame(_sum).T
    row_total["Var"] = "Total"
    
    return pd.concat([df, row_total], ignore_index=True)
