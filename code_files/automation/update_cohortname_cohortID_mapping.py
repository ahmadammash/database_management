# Import the required libraries
import json

import pandas as pd

from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# Import the required functions
from code_files.utilities.processing.cleaners import clean_logic

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

# define the path to the cohort mapping json file
cohort_names_mapping_path = project_root / "json_files" / "cohort_names_mapping.json"

# define the path to the cohort name to cohort ID mapping json file
cohortname_cohortID_mapping_path = (
    project_root / "json_files" / "cohortname_cohortID_mapping.json"
)

# load the cohort names mapping json file
with open(cohort_names_mapping_path, "r") as f:
    cohort_names_mapping = json.load(f)


# Query the tristar pathway from the database
tristar_pathway = pd.read_sql("SELECT * FROM imp_documents.tristar_pathway", cnxn)

# Select just the cohort name and cohort ID columns.
tristar_pathway = tristar_pathway[["Cohort", "Cohort_ID"]]


# Define a function to get the cohort name to cohort ID mappings
def update_cohortname_cohortID_mappings():
    """
    This function uses the tristar pathway dataset to update the cohort name to cohort ID mapping json file.
    The function collects all the cohort names from the tristar pathway document, cleans them using the
    clean_logic function, and maps them to the corresponding cohort ID from the tristar pathway dataset
    before dumping the changes to the JSON file/
    """
    # create the empty dictionary for the mapping
    cohortname_cohortID_dict = {}
    # Loop through the tristar pathway cohort names and IDs columns
    for cohort, cohort_id in zip(
        tristar_pathway["Cohort"], tristar_pathway["Cohort_ID"]
    ):
        # Clean the cohort name
        cleaned_cohort = clean_logic(cohort)
        # Set the key to the cleaned cohort name and the value to the corresponding ID
        cohortname_cohortID_dict[cleaned_cohort] = cohort_id
    # Dump the changes in the JSON file
    with open(cohortname_cohortID_mapping_path, "w") as f:
        json.dump(cohortname_cohortID_dict, f, indent=4)
