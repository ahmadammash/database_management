# Import the required libraries
import pandas as pd

# Import the required functions
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# Establish a connection to the SQL Server database
project_root, cnxn, cursor, engine, current_file = setup_cnxn()


def create_combined_supp_unanon():

    # Query all the supplementary unanonymised tables
    all_supp_unanon_query = """
    SELECT TABLE_SCHEMA, TABLE_NAME 
    FROM INFORMATION_SCHEMA.TABLES
    WHERE TABLE_NAME LIKE '%Supplementary_Unanon%'
    """
    all_supp_unanon = pd.read_sql(
        all_supp_unanon_query,
        engine
    )

    # Query each supplementary table
    supp_unanon_dfs_list = []
    for idx, row in all_supp_unanon.iterrows():
        schema = row['TABLE_SCHEMA']
        table = row['TABLE_NAME']
        df_query = f"""
        SELECT * FROM [{schema}].[{table}]
        """
        df = pd.read_sql(
            df_query,
            cnxn
        )
        supp_unanon_dfs_list.append(df)

    # Concatenate the supplementary dfs
    supp_unanon_combined = pd.concat(supp_unanon_dfs_list, ignore_index=True)

    # upload the combined dataframe to the imp_documents schema
    supp_unanon_combined.to_sql(
        "supp_unanon_combined",
        engine,
        if_exists = "replace",
        schema="imp_documents",
        index=False
    )
