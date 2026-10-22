# import the required libraries
import json

import pandas as pd
from sqlalchemy import text

# Import required functions
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# setup a connection to the database
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

# load the cohort name to schema mapping json file
cohort_schema_mapping_path = project_root / "json_files" / "cohort_schema_mapping.json"
with open(cohort_schema_mapping_path, "r") as f:
    cohort_schema_mapping = json.load(f)

# load the cohort names mapping json file
cohort_mapping_path = project_root / "json_files" / "cohort_names_mapping.json"
with open(cohort_mapping_path, "r") as f:
    cohort_mapping = json.load(f)


query = """
SELECT table_schema, table_name
FROM information_schema.tables
WHERE table_schema not in (
'information_schema',
'inventory',
'hgugm',
'chu_nice',
'iom_aviano',
'pg_catalog',
'hannover',
'misc_docs',
'nhsb',
'tma_data',
'tma'
)
"""
all_tables = pd.read_sql(query, engine)

for idx, row in all_tables.iterrows():
    schema = row["table_schema"]
    table = row["table_name"]
    with engine.connect() as cnxn:
        cnxn.execute(text(f'DROP TABLE IF EXISTS "{schema}"."{table}"'))
        cnxn.commit()
