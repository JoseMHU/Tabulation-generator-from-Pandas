import pandas as pd
from excel_class import ExcelFile
from utilities import (category_order, remove_duplicates,
                       interpret_instruction, calculate_percentages)


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
        table_tupe = 3 -> % rows
    Returns
    -------
    None

    """
    if table_type != 0:
        file_name = f"{db_name}_percentage_table.xlsx"
    else:
        file_name = f"{db_name}_absolute_value_table.xlsx"
    
    secondary_vars = [db_vars.loc[index, "VAR"] for index in db_vars.index if db_vars.loc[index, "DISAGGREGATE"]]
    excel_file = ExcelFile(file_name=file_name, tables_type=table_type)
    
    def structure_generator(field: str) -> dict:
        field_set = remove_duplicates(db[field].tolist())
        field_set.insert(0, "Total")
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
                    factor = interpret_instruction(db_vars.loc[index, "SUM"], question)
                else:
                    factor = db_vars.loc[index, "SUM"]
                
                for secondary_var in secondary_vars:
                    df_dict = {}
                    structure = structure_generator(secondary_var)
                    columns_to_convert = [x for x in structure if x not in ("Var", "Total")]
                    
                    for var in remove_duplicates(db[question].dropna()):
                        structure["Var"].append(var)
                        
                        if split_df_by:
                            _filter = (db[question] == var) & (db[split_df_by[0]] == split_df_by[1])
                        else:
                            _filter = (db[question] == var)
                            
                        if type(factor) is tuple:
                            if factor[2] == "*":
                                multiplication = db.loc[_filter, factor[0]] * db.loc[_filter, factor[1]]
                                value = multiplication.sum()
                            else:
                                raise ValueError("Unrecognized operation in SUM of Var column")
                        else:
                            value = db.loc[_filter, factor].sum()
                        structure["Total"].append(value)
                        
                        for s_var in columns_to_convert:
                            
                            if split_df_by:
                                _filter = (db[secondary_var] == s_var) & (db[question] == var) & (
                                        db[split_df_by[0]] == split_df_by[1])
                            else:
                                _filter = (db[secondary_var] == s_var) & (db[question] == var)
                                
                            if type(factor) is tuple:
                                if factor[2] == "*":
                                    multiplication = db.loc[_filter, factor[0]] * db.loc[_filter, factor[1]]
                                    value = multiplication.sum()
                                else:
                                    raise ValueError("Unrecognized operation in SUM of Var column")
                            else:
                                value = db.loc[_filter, factor].sum()
                            structure[s_var].append(value)
                    df = pd.DataFrame(structure)
                    if table_type != 0:
                        df = calculate_percentages(structure, df, db, table_type, factor, split_df_by)
                    df_dict[secondary_var] = df
                    data.append((question, df_dict))
                    
        for question, df_dict in data:
            excel_file.add_sheet(sheet_name=f"Section - {page}", data=df_dict, main_var_name=question,
                                 footer_text=footer_text)
        excel_file.restart()
    excel_file.save()
