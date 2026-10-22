# Import the Required Libraries
import pandas as pd

# Import the required functions
from code_files.utilities.config.mappings import set_data_types_tristar_pathway
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

# Define the patient list configuration file
config_file = project_root / "json_files" / "patientList_config.json"

# define the path to the TriStar pathway file
tristar_pathway_path = (
    "C:/Users/AmmashA.IA/OneDrive - NHS/TrisStar metadata/Plans/TriStar pathway.xlsx"
)

# read in the TriStar pathway dataset
tristar_pathway = pd.read_excel(tristar_pathway_path, sheet_name="Shipping details")

# convert the columns to title case and rename columns
tristar_pathway.columns = [
    col.upper() if "ID" in col or "MTA" in col else col.title()
    for col in tristar_pathway.columns
]
tristar_pathway.columns = tristar_pathway.columns.str.replace(" ", "_")
column_name_map = {"COHORT_ID": "Cohort_ID"}
tristar_pathway = tristar_pathway.rename(columns=column_name_map)

# process the date columns
date_columns = ["Shipment_Date", "Box_Updated"]
for date in date_columns:
    tristar_pathway[date] = tristar_pathway[date].replace("Yes", pd.NA)

# Upload to the SQL database
tristar_pathway.to_sql(
    name="tristar_pathway",
    con=engine,
    if_exists="replace",
    schema="imp_documents",
    index=False,
    dtype=set_data_types_tristar_pathway(tristar_pathway),
)
