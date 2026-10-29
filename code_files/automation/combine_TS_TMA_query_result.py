# Import the required libraries
import sys

import pandas as pd

import os
import sys

from pathlib import Path

# Import functions from the project directory
from code_files.utilities.io.setup_db_cnxn import setup_cnxn
from code_files.utilities.io.database import create_gcloud_engine

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()
env_path = project_root / "code_files" / "utilities" / "io" / ".env"
engine = create_gcloud_engine(env_path)

# get the code files directory
code_files = project_root / "code_files"

# Set the mssql variable
if "AmmashA.IA" in str(current_dir):
    mssql = True
else:
    mssql = False



# add the root path to the system
if str(code_files) not in sys.path:
    sys.path.insert(0, str(code_files))

from code_files.data_processing.utilities_processing.functions import clean_ids
from code_files.utilities.io.database import create_gcloud_engine


# ----------------------------------------------------------------------------------------------------------------------
# In[]:
# query the sample id to tma number mapping
link_file_query = """
SELECT * FROM misc_docs.link_file
"""
link_file_df = pd.read_sql(link_file_query, engine)


nhsb_query = """
SELECT * FROM nhsb.nhsb_tri_samples
WHERE "Sample_ID" IN (
'TS-17 - 0900',
'TS-17 - 0918',
'TS-17 - 0937',
'TS-16 - 0019',
'TS-16 - 0534',
'TS-17 - 1100',
'TS-17 - 1085'
)
"""
nhsb_df = pd.read_sql(nhsb_query, engine)

chu_nice_query = """
SELECT * FROM chu_nice.chu_nice_consolidated
"""
chu_nice_df = pd.read_sql(chu_nice_query, engine)

hannover_query = """
SELECT * FROM hannover.combined_hannover
"""
hannover_combined = pd.read_sql(hannover_query, engine)
harmonised_hannover_combined = (
    hannover_combined.groupby('Sample_ID', as_index = False).first()
)

# concatenate all the dataframes from the schema named iom_aviano
all_iom_aviano_tbls_query = """
SELECT table_name, table_schema
FROM information_schema.tables
where table_schema = 'iom_aviano'
"""
iom_dfs_list = []
all_iom_aviano_tbls = pd.read_sql(all_iom_aviano_tbls_query, engine)
for idx, row in all_iom_aviano_tbls.iterrows():
    tbl = row['table_name']
    schema = row['table_schema']
    # query each df
    df_query = f"""
    SELECT * FROM "{schema}"."{tbl}"
    """
    df = pd.read_sql(df_query, engine)
    iom_dfs_list.append(df)
iom_combined = pd.concat(iom_dfs_list, ignore_index=True)

harmonised_iom_combined = (
    iom_combined.groupby('Sample_ID', as_index = False).first()
)

# concatenate all dataframes from the schema named hgugm
all_hgugm_tbls_query = """
SELECT table_name, table_schema
FROM information_schema.tables
WHERE table_schema = 'hgugm'
"""
hgugm_dfs_list = []
all_hgugm_tbls = pd.read_sql(all_hgugm_tbls_query, engine)
for idx, row in all_hgugm_tbls.iterrows():
    tbl = row['table_name']
    schema = row['table_schema']
    # query each df
    df_query = f"""
    SELECT * FROM "{schema}"."{tbl}"
    """
    df = pd.read_sql(df_query, engine)
    hgugm_dfs_list.append(df)
hgugm_combined = pd.concat(hgugm_dfs_list, ignore_index=True)   
harmonised_hgugm_combined = (
    hgugm_combined.groupby('Sample_ID', as_index = False).first()
)
tristar_inventory = pd.read_sql(
    """
    SELECT * FROM inventory.tristar_inventory
    """, engine
)

# In[]:
sample_ids = [
    s.strip().replace(' ','') for s in """
TS-18 - 0380
TS-18 - 0383
TS-18 - 1965
TS-18 - 1970
TS-18 - 1993
TS-18 - 0281





    """.splitlines() if s.strip()
]
original_ids = [
    s.strip().replace(' ', '') for s in """
LH12-0623  5C3
LH12-2100 A
LB16 229 TL A
LB16 950 TL B
LB16 998 TL B
LB17 260 TLC bis



""".splitlines() if s.strip()


]
