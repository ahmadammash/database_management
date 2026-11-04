# Import the required libraries
import pandas as pd

# Import the required functions
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# Establish a connection to the database
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

# query all the supplementary unanonymised table and schema names
query_suppTables = """
SELECT TABLE_NAME, TABLE_SCHEMA
FROM INFORMATION_SCHEMA.TABLES
WHERE TABLE_NAME LIKE '%Supplementary_Unanon%'
"""

# run the query
all_supp_unanon_tables = pd.read_sql(
    query_suppTables, cnxn
)

# query the master list (anonymised)
query_masterList = """
SELECT * FROM [imp_documents].[master_list_long]
"""

# run the query
master_list = pd.read_sql(
    query_masterList, cnxn
)

# check which supplementary tables do not have an "Indication" column - if any
for idx, row in all_supp_unanon_tables.iterrows():
    schema = row['TABLE_SCHEMA']
    tbl = row['TABLE_NAME']
    query = f"""
    SELECT * FROM [{schema}].[{tbl}]
    """
    df_supp = pd.read_sql(
        query, cnxn
    )
    if 'Indication' not in df_supp.columns:
        print(tbl)
