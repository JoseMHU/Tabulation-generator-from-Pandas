from main_logic import start_code
import pandas as pd
    
if __name__ == "__main__":
    """
    table_type: int
        Specifies the format in which the data is presented within the tables.
        table_type = 0 -> Absolute value
        table_type = 1 -> %total
        table_type = 2 -> %columns
        table_tupe = 3 -> %rows
    """
    for db_name in ("TEST1", "TEST2"):
        start_code(db_name=db_name,
                   table_type=2,
                   db=pd.read_excel("example_db//example.xlsx",
                                    sheet_name="DB"),
                   db_vars=pd.read_excel("example_db//example.xlsx",
                                         sheet_name="VAR"),
                   split_df_by=("DB", db_name),
                   footer_text="Footer")
        