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

# query the sample id to tma number mapping
tma_mapping_query = """
SELECT * FROM tma.sample_id_tma_num_mapping
"""
tma_mapping = pd.read_sql(tma_mapping_query, engine)