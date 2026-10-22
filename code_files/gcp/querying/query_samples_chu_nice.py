# Import the required libraries
import warnings

import pandas as pd

# Disable UserWarnings and FutureWarnings
warnings.filterwarnings("ignore", category=UserWarning)
warnings.filterwarnings("ignore", category=FutureWarning)

# Import the required functions from the files
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

# ----------------------------------------------------------------------------------------------------

# query the chu-nice table
query_chu_nice = """
    SELECT * FROM "chu_nice"."chu_nice_consolidated"
    """
with engine.connect() as cnxn:
    chu_nice_data = pd.read_sql(query_chu_nice, cnxn)

# drop any duplicates across the three ID columns
chu_nice_data = chu_nice_data.drop_duplicates(
    subset=["Sample_ID", "Biobank_ID", "Pathology_ID"]
)
