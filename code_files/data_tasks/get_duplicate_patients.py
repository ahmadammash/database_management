import sys
from pathlib import Path

import pandas as pd

# Add the project to the root directory
current_dir = Path(__file__).resolve()
project_root = next(
    (p for p in Path(".").resolve().parents if (p / "pyproject.toml").exists()),
    Path(".").resolve(),
)
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

# import the required functions
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# setup a connection to the database
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

# Query the patient list
query_patientList = """
SELECT * FROM [imp_documents].[patientList]
"""
patientList = pd.read_sql(query_patientList, cnxn)
unique_NHS_Numbers = set(
    list(patientList.loc[patientList.NHS_Number.notna()]["NHS_Number"])
)
patientNumber = 0

NHS_patientNumber_mapping = {k: i for i, k in enumerate(unique_NHS_Numbers, start=1)}

# Add the Patient Number column to the patient list using the mapping dictionary created
for idx, row in patientList.iterrows():
    patientList.loc[idx, "Patient_Number"] = NHS_patientNumber_mapping.get(
        row["NHS_Number"]
    )


# Get the patients with duplicate patient numbers
duplicate_records = patientList[
    patientList.duplicated(subset=["Patient_Number"], keep=False)
]
