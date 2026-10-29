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

# define a function to combine and harmonise data from a specific center
def combine_TS_center_data(schema_name):
    all_tables_query = f"""
    SELECT table_name, table_schema
    FROM information_schema.tables
    WHERE table_schema = '{schema_name}'
    """
    all_tables = pd.read_sql(all_tables_query, engine)
    dfs_list = []
    for idx, row in all_tables.iterrows():
        tbl = row['table_name']
        schema = row['table_schema']
        df_query = f"""
        SELECT * FROM "{schema}"."{tbl}"
        """
        df = pd.read_sql(df_query, engine)

        # Make sure the Sample_ID column exists in the dataframe
        if 'Sample_ID' not in df.columns:
            print(f"The dataset {tbl} does not have a Sample_ID column.")
        
        dfs_list.append(df)
    df_combined = pd.concat(dfs_list, ignore_index=True)
    harmonised_df_combined = (
        df_combined.groupby('Sample_ID', as_index = False).first()
    )
    return harmonised_df_combined


hgugm_combined = combine_TS_center_data("hgugm")
iom_combined = combine_TS_center_data("iom_aviano")



