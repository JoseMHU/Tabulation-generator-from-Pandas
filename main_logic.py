import pandas as pd
from excel_class import ExcelFile
from utilities import (category_order, remove_duplicates,
                       interpret_instruction, calculate_percentages, insert_total_row)


def start_code(db_name: str, table_type: int, db: pd.DataFrame, db_vars: pd.DataFrame, footer_text=None,
               split_df_by=None) -> None:
    """

    Parameters
    ----------
    db_name: str
        Name of the Excel file to create.
    db: pd.DataFrame
        DataFrame from which the output tables will be created on the pages of the Excel workbook.
    db_vars: pd.DataFrame
        Dataframe containing the list of variables to be used to generate the tabulations, the variables used for data
        cross-referencing, and the method for processing the results to be displayed.
    footer_text: str | None
        Table footer for the tabulations.
    split_df_by : tuple
        Value by which the data source is divided into groups.
        It is a tuple where the 1st element is the name of the column to be
        filtered and the 2nd is the value to be filtered by.
    table_type : int
        Specifies the format in which the data is presented within the tables.
        table_type = 0 -> Absolute value
        table_type = 1 -> % total
        table_type = 2 -> % columns
        table_type = 3 -> % rows
        table_type = 4 -> % Simple absolute value and % total
        table_type = 5 -> Long description
    Returns
    -------
    None

    """
    metadata = pd.DataFrame()
    if table_type == 4:
        file_name = f"{db_name}_simple_table.xlsx"
    elif table_type != 0:
        file_name = f"{db_name}_percentage_table.xlsx"
    else:
        file_name = f"{db_name}_absolute_value_table.xlsx"
    
    secondary_vars_list = [db_vars.loc[index,"VAR"] for index in db_vars.index if db_vars.loc[index, "DISAGGREGATE"]]

    if not secondary_vars_list and table_type != 4:
        raise TypeError("A disaggregation of data was indicated (table_type != 4) but there is no variable indicated to"
                        "disaggregate. (All variables in DISAGGREGATE of db_vars = False)")
    elif table_type == 4:
        secondary_vars_list = ["pass"]
      
    if table_type <= 4:
      excel_file = ExcelFile(file_name=file_name, tables_type=table_type)
    else:
      excel_file = None
    
    def structure_generator(field: str) -> dict:
        field_set = remove_duplicates([x for x in db[field].tolist() if str(x) != "nan"])
        field_set.insert(0, "Var")
        try:
            if set(field_set).issubset(category_order[field]):
                category = category_order[field]
            else:
                category = field_set
        except KeyError:
            category = field_set
            
        result = {}
        for i in category:
            result[i] = []
        
        return result
    
    for page in remove_duplicates(db_vars["PAGE"]):
        data = []
        for index in db_vars.index:
            if db_vars.loc[index, "PAGE"] == page:
                question = db_vars.loc[index, "VAR"]
                
                if "${" in db_vars.loc[index, "SUM"]:
                    factor = interpret_instruction(db_vars.loc[index, "SUM"])
                else:
                    factor = db_vars.loc[index, "SUM"]

                for secondary_var in secondary_vars_list:
                    if table_type == 4:
                        secondary_var = question
                    df_dict = {}
                    structure = structure_generator(secondary_var)
                    columns_to_convert = [x for x in structure if x not in ("Var", "Total")]

                    list_var = [x for x in structure_generator(question).keys() if x not in ("Var", "Total")]

                    for var in list_var:
                        structure["Var"].append(var)

                        for s_var in columns_to_convert:

                            if split_df_by:
                                _filter = (db[secondary_var] == s_var) & (db[question] == var) & (
                                            db[split_df_by[0]] == split_df_by[1])
                            else:
                                _filter = (db[secondary_var] == s_var) & (db[question] == var)

                            if type(factor) is tuple:
                                if factor[3] == "*":
                                    multiplication = db.loc[_filter, factor[0]] * db.loc[_filter, factor[2]]
                                    value = multiplication.sum()
                                else:
                                    raise ValueError("Unrecognized operation in SUM of Var column")
                            else:
                                value = db.loc[_filter, factor].sum()
                            structure[s_var].append(value)

                    df = pd.DataFrame(structure)
                    df_without_var = df.drop(columns=["Var"])
                    df["Total"] = df_without_var.sum(axis=1)
                    cols = df.columns.tolist()
                    cols.insert(1, cols.pop(cols.index("Total")))
                    df = df[cols]

                    if table_type != 0:
                        df = calculate_percentages(structure, df, table_type)
                    elif table_type == 4:
                        df = calculate_percentages(None, df, table_type)

                    if table_type != 3:
                        df_dict[secondary_var] = insert_total_row(df)
                    else:
                        df_dict[secondary_var] = df

                    data.append((question, df_dict,  f"This table was calculated by a variable operation {factor[0]} * "
                                                     f"{factor[2]}" if type(factor) is tuple else None))
                    if table_type == 4:
                        break
                      
                    elif table_type == 5:
                        df.insert(0, 'Ask', question)
                        if len(metadata) > 0:
                            metadata = pd.concat([metadata, df], ignore_index=True)
                        else:
                            metadata = df
                        break
                      
        if table_type <= 4:            
          for question, df_dict, footer_text_factor in data:
              if footer_text and footer_text_factor:
                  footer_text_table = str(footer_text) + " // " + str(footer_text_factor)
              elif footer_text_factor:
                  footer_text_table = str(footer_text_factor)
              elif footer_text:
                  footer_text_table = str(footer_text)
              else:
                  footer_text_table = None
              excel_file.add_sheet(sheet_name=f"Section - {page}", data=df_dict, main_var_name=question,
                                   footer_text=footer_text_table)
        if table_type <= 4:    
          excel_file.restart()
    if table_type <= 4:
      excel_file.save()
    else:
      metadata.to_excel("results//METADATA.xlsx", index=False, header=True)
