# Import the required libraries
import re
import warnings

import numpy as np
import pandas as pd
from sqlalchemy import text

# Disable UserWarnings and FutureWarnings
warnings.filterwarnings("ignore", category=UserWarning)
warnings.filterwarnings("ignore", category=FutureWarning)

# Import the required functions from the files
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

# ----------------------------------------------------------------------------------------------------

# query all the NHSB samples
fetch_nhsb_tables = """
SELECT table_schema, table_name
FROM information_schema.tables
WHERE table_schema NOT IN ('information_schema', 'pg_catalog', 'dbo', 'guest')
                 AND table_schema NOT LIKE 'db%' AND table_schema NOT LIKE '%sys%'
                 AND table_schema IN ('hgugm', 'hannover', 'nhsb')
"""

# define a list to contain all the dfs from the query output
dfs = []

# run the query to fetch all the user-defined tables
with engine.connect() as cnxn:
    all_tables = cnxn.execute(text(fetch_nhsb_tables))

    # convert the result into a dataframe
    all_tables = pd.DataFrame(all_tables.fetchall(), columns=all_tables.keys())


# Query each table in the all_tables dataframe
for idx, row in all_tables.iterrows():
    schema = row["table_schema"]
    table = row["table_name"]

    try:
        single_query = f'SELECT "Sample_ID" FROM "{schema}"."{table}"'

        with engine.connect() as cnxn:
            df = pd.read_sql(single_query, cnxn)
            df["source_table"] = table
            dfs.append(df)
    except Exception as e:
        print(f"""
              Query failed for the table "{schema}"."{table}"
              """)
        print(e)


# Concatenate the list of dataframes to one dataframe
compiled_dfs = pd.concat(dfs)

# prelim datasets
prelim_datasets = ["hgugm_2014_pancreas", "nhsb_tri_samples"]

# add the dataset_type column
compiled_dfs["dataset_type"] = np.where(
    compiled_dfs["source_table"].isin(prelim_datasets), "Prelim", "Follow up"
)

# Sort by the dataset_type column
compiled_dfs = compiled_dfs.sort_values(by="dataset_type")

# drop any duplicate sample ids
compiled_dfs = compiled_dfs.drop_duplicates("Sample_ID", keep="first")

# group by the dataset type column
compiled_dfs_grouped = compiled_dfs.groupby("dataset_type").count()

# clean the sample ids
compiled_dfs["Sample_ID"] = compiled_dfs["Sample_ID"].apply(
    lambda x: re.sub(r"[^A-Za-z0-9]", "", str(x))
)

# %%
