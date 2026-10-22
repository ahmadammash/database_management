# Import the required libraries
import sys

import pandas as pd

# Import functions from the project directory
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

# Set the mssql variable
if "AmmashA.IA" in str(current_dir):
    mssql = True
else:
    mssql = False

# get the code files directory
code_files = project_root / "code_files"

# add the root path to the system
if str(code_files) not in sys.path:
    sys.path.insert(0, str(code_files))

from code_files.data_processing.utilities_processing.functions import clean_ids
from code_files.utilities.io.database import create_gcloud_engine

regex_pattern = r"[^A-Za-z0-9]"

try:
    link_file = pd.read_excel(
        "C:/Users/medaamm/Box/Data file (AA)/Pending follow up data (TriStar projects)/Link file/CTA Total Inventory V5.4 (22.03.31) Sent.xlsx",
        sheet_name="Archived Inventory Report",
        header=1,
    )
except (FileNotFoundError, OSError):
    link_file = pd.read_excel(
        "C:/Users/AmmashA.IA/Box/Data file (AA)/Pending follow up data (TriStar projects)/Link file/CTA Total Inventory V5.4 (22.03.31) Sent.xlsx",
        sheet_name="Archived Inventory Report",
        header=1,
    )

link_file = link_file.iloc[:, :17]
column_map = {
    "SAMPLE ID#": "Sample_ID",
    "Original ID": "Original_ID",
    "Reference Incoming": "Reference_Incoming",
    "TRK-Incoming": "TRK_Incoming",
    "Bl": "Block_Number",
    "Type": "Histological_Diagnosis_Original",
    "Sub-Type": "Histological_Subtype",
    "Archived Inventory Report": "Archived_Inventory_Report",
}
link_file = link_file.rename(columns=column_map)

link_file = clean_ids(link_file, ["Sample_ID", "Original_ID"], regex_pattern)

env_path = project_root / "code_files" / "utilities" / "io" / ".env"
engine = create_gcloud_engine(env_path)

cols_to_trim = ["Sample_ID", "Original_ID"]
for col in cols_to_trim:
    link_file[col] = link_file[col].str.replace(r"\s+", "", regex=True)

link_file.to_sql(
    "link_file", engine, schema="misc_docs", if_exists="replace", index=False
)
