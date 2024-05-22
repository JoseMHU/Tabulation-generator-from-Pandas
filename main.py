from main_logic import start_code
import pandas as pd

if __name__ == "__main__":
    """

    Parameters
    ----------
    split_df_by : tuple
        Value by which the data source is divided into groups. 
        It is a tuple where the 1st element is the name of the column to be
        filtered and the 2nd is the value to be filtered by.
    table_type : int
        Specifies the format in which the data is presented within the tables.
        table_type = 0 -> Absolute value
        table_type = 1 -> % Total
        table_type = 2 -> % Columns
        table_type = 3 -> % Rows
        table_type = 4 -> % Simple absolute value and % total
    db : pd.DataFrame
        Main data source.
    db_vars : pd.DataFrame
        Base settings and specifications.
    """
    for i in range(0, 5):
        start_code(db_name=f"TEST_{i}", table_type=i,
                   db=pd.read_excel("example_db//example.xlsx",
                                    sheet_name="DB"),
                   db_vars=pd.read_excel("example_db//example.xlsx",
                                         sheet_name="VAR"),
                   split_df_by=("DB", "TEST1"), footer_text="Footer")
