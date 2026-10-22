# Import the required libraries
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


# query all the IOM samples
fetch_iom_tables = """
SELECT table_schema, table_name
FROM information_schema.tables
WHERE table_schema NOT IN ('information_schema', 'pg_catalog', 'dbo', 'guest')
                 AND table_schema NOT LIKE 'db%' AND table_schema NOT LIKE '%sys%'
                 AND table_schema = 'iom_aviano'
"""

# define a list to contain all the dfs from the query output
dfs = []

# Run the query to fetch all user_defined tables
with engine.connect() as cnxn:
    all_tables = cnxn.execute(text(fetch_iom_tables))

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
prelim_datasets = [
    "ATCG_1072",
    "ATCG_1076",
    "ATCG_1109",
    "CRC_KR",
    "Multi_TMA_120",
    "Multi_TMA_450",
    "TA_1081",
    "TMO_Bladder",
    "TMO_Breast",
    "TMO_Colorectal",
    "TMO_Glioma",
]

# follow up datasets
followup_datasets = [
    "CRC_Data",
    "NHL_DLBCL",
    "NHL_FL",
    "NHL_TMA2_Italy",
    "NHL_TMA2_UK",
    "TMO_Cutaneous",
    "TMO_Papillary_Thyroid",
    "TMO_Prostate",
]

# add the dataset_type column to compiled_dfs
conditions = [
    compiled_dfs["source_table"].isin(prelim_datasets),
    compiled_dfs["source_table"].isin(followup_datasets),
]
options = ["Prelim", "Follow up"]
compiled_dfs["dataset_type"] = np.select(conditions, options, "N/A")

# Drop the duplicate sample IDs
compiled_dfs = compiled_dfs.drop_duplicates("Sample_ID")

# group by the dataset_type column
compiled_dfs_grouped = compiled_dfs.groupby("dataset_type").count()
