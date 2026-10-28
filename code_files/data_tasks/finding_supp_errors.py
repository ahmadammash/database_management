# ================================================================================================================
# IMPORTS AND SYSTEMS SETUP
# ================================================================================================================

# Import the required libraries
import warnings
from tkinter import *

import pandas as pd

# Suppress warnings
warnings.simplefilter("ignore", UserWarning)
warnings.simplefilter("ignore", FutureWarning)

# Import functions from the project directory
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# ================================================================================================================
# DATABASE CONNECTION & INITIALISATION
# ================================================================================================================

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()


query_anon_supp_tbls = """SELECT TABLE_NAME, TABLE_SCHEMA FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME LIKE '%Supplementary_Anon%'"""
all_anon_supp_tbls = pd.read_sql(query_anon_supp_tbls, cnxn)
query_unanon_supp_tbls = """SELECT TABLE_NAME, TABLE_SCHEMA FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME LIKE '%Supplementary_UnAnon%'"""
all_unanon_supp_tbls = pd.read_sql(query_unanon_supp_tbls, cnxn)
# for idx, row in all_anon_supp_tbls.iterrows():
#     schema = row["TABLE_SCHEMA"]
#     table = row["TABLE_NAME"]
#     query_tbl = f"""
#     SELECT * FROM [{schema}].[{table}]
#     """
#     df = pd.read_sql(query_tbl, cnxn)
#     if "index" in df.columns:
#         print(f"{schema}.{table}")

# all_anon_supp_prefix = (
#     all_anon_supp_tbls["TABLE_NAME"]
#     .str.split("_", n=2)
#     .str[:2]
#     .str.join("_")
# )

# all_unanon_supp_prefix = (
#     all_unanon_supp_tbls["TABLE_NAME"]
#     .str.split("_", n=2)
#     .str[:2]
#     .str.join("_")
# )
# counter = 0
# for tbl in list(all_unanon_supp_prefix):
#     if tbl not in list(all_anon_supp_prefix):
#         print(tbl)
#         counter+=1

missing_tbls = pd.DataFrame()
all_columns_list = []
# get the prefixes from the anonymised tables
anon_supp_tbls = list(all_anon_supp_tbls['TABLE_NAME'])
prefixes_anon = ["_".join(table.split("_")[:2]) for table in anon_supp_tbls]

for idx, row in all_unanon_supp_tbls.iterrows():
    # get the schema, table, and prefix from the unaonymised tables
    schema = row["TABLE_SCHEMA"]
    table = row["TABLE_NAME"]
    prefix = "_".join(table.split("_")[:2])

    # get the missing anonymised supplementary tables
    # if prefix not in prefixes_anon:
    #     print(f"{schema}.{table}")

    # query the table
    query = f"""
    SELECT * FROM [{schema}].[{table}] 
    """
    df = pd.read_sql(query, cnxn)

    # Drop the DonorID column when it exists alongside the NHS Number column
    # if 'NHS_Number' in df.columns and 'DonorID' in df.columns:
    #     df = df.drop(columns=['DonorID'])

    # get all the column names from the dataframes
    all_columns_list.extend(df.columns.tolist())

    if 'Unnamed: 6' in df.columns:
        print(table)

all_columns_list = set(all_columns_list)

    


