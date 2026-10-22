# Import the required libraries
import sys
from pathlib import Path

import pandas as pd

# Add the project root to the current directory
current_dir = Path(__file__).resolve()
project_root = next(
    (p for p in Path(".").resolve().parents if (p / "pyproject.toml").exists()),
    Path(".").resolve(),
)
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))
json_files_dir = project_root / "json_files"

# Import the required functions
from code_files.utilities.io.database import create_gcloud_engine

# Establish a connection to GCP
if "medaamm" in str(current_dir) or "ahmad" in str(current_dir):
    # create the connection engine to GCP
    env_path = project_root / "code_files" / "utilities" / "io" / ".env"
    engine = create_gcloud_engine(path=env_path)

    # display a message if the engine failed to be created
    if engine is None:
        print("Cannot start upload without a valid database connection.")

# Query the tristar inventory
query_inventory = """
SELECT * FROM inventory.tristar_inventory
"""

# Query the link file
query_link_file = """
SELECT * FROM misc_docs.link_file
"""

# Query the CHU-NICE consolidated file
query_chu_nice = """
SELECT * FROM chu_nice.chu_nice_consolidated
"""

with engine.connect() as connection:
    tristar_inventory_df = pd.read_sql(query_inventory, connection)
    link_file_df = pd.read_sql(query_link_file, connection)
    chu_nice = pd.read_sql(query_chu_nice, connection)

# Add the Sample_IDs which are in the link file but not in the tristar inventory
new_sample_ids = link_file_df[
    ~link_file_df["Sample_ID"].isin(tristar_inventory_df["Sample_ID"])
]
samples_data_to_add = new_sample_ids[
    [
        "Sample_ID",
        "Original_ID",
        "Origin",
        "Gender",
        "Age",
        "T",
        "N",
        "M",
        "Grade",
        "Stage",
        "Block_Number",
        "Histological_Diagnosis_Original",
        "Histological_Subtype",
    ]
]
# Add the data to the existing tristar inventory
tristar_inventory_df = pd.concat(
    [tristar_inventory_df, samples_data_to_add], ignore_index=True
)

# Import the 4526 file
file_4526 = pd.read_excel(
    "C:/Users/medaamm/Box/Data file (AA)/Miscellaneous Requests/4526 NSCLC post treatment with follow up data.xlsx",
    sheet_name="CHU-Nice NSCLC post treatment",
    header=13,
)

# rename the columns in the dataframe
column_name_map = {
    "FFPE block ID": "Sample_ID",
    "Neoadjuvant type of treatment": "Neoadjuvant_Treatment_Type",
    "Beginning of treatment": "Neoadjuvant_Treatment_StartDate",
    "End of treatment date": "Neoadjuvant_Treatment_EndDate",
}
file_4526 = file_4526.rename(columns=column_name_map)

file_4526["Sample_ID"] = file_4526["Sample_ID"].replace(r"\s+", "", regex=True)

chu_nice = chu_nice.merge(
    file_4526[
        [
            "Sample_ID",
            "Neoadjuvant_Treatment_Type",
            "Neoadjuvant_Treatment_StartDate",
            "Neoadjuvant_Treatment_EndDate",
        ]
    ],
    how="left",
    on="Sample_ID",
)

# upload the chu_nice dataframe to GCP
with engine.connect() as connection:
    chu_nice.to_sql(
        "chu_nice_consolidated", con=connection, schema="chu_nice", if_exists="replace"
    )
