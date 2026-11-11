# Import the required libraries
import pandas as pd
from sqlalchemy.types import BIGINT
#from code_files.utilities.config.mappings import set_data_types

# Import the required functions
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# Establish a connection to the SQL Server database
project_root, cnxn, cursor, engine, current_file = setup_cnxn()


def create_combined_supp_anon():

    # Query all the supplementary anonymised tables
    all_supp_anon_query = """
    SELECT TABLE_SCHEMA, TABLE_NAME 
    FROM INFORMATION_SCHEMA.TABLES
    WHERE TABLE_NAME LIKE '%Supplementary_Anon%'
    """
    all_supp_anon = pd.read_sql(
        all_supp_anon_query,
        engine
    )

    # Query each supplementary table
    supp_anon_dfs_list = []
    for idx, row in all_supp_anon.iterrows():
        schema = row['TABLE_SCHEMA']
        table = row['TABLE_NAME']
        df_query = f"""
        SELECT * FROM [{schema}].[{table}]
        """
        df = pd.read_sql(
            df_query,
            cnxn
        )
        supp_anon_dfs_list.append(df)

    # Concatenate the supplementary dfs
    supp_anon_combined = pd.concat(supp_anon_dfs_list, ignore_index=True)

    # create the data types dictionary
    # data_types=set_data_types(supp_anon_combined, "anonymised supplementary")
    data_types = {
        'NHS_Number': BIGINT
    }

    # upload the combined dataframe to the imp_documents schema
    supp_anon_combined.to_sql(
        "supp_anon_combined",
        engine,
        if_exists = "replace",
        schema="imp_documents",
        index=False,
        dtype=data_types
    )
