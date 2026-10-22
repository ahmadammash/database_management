# Import the required libraries
import json

# Import the required functions
from code_files.querying.fetch_information import fetch_schema_names
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

# Define a function to update the onedrive to cohort schema mapping
def update_onedrive_cohort_schema_mapping(
    cohort_names_mapping, onedrive_cohort_schema_mapping
):
    """
    This function updates the onedrive cohort to schema names mapping. First, the function retrieves the schema names
    directly from the database. Second, the function loops through the cohort names mapping dictionary and the schema
    names before mapping each onedrive cohort name to a schema name based on the words in the cohort name. The function
    then dumps the changes in the json file.
    """
    # Get the schema names
    schema_names = fetch_schema_names()

    # Loop through the cohort names mapping dictionary
    for cohort in cohort_names_mapping.values():
        # Loop through the schema names
        for schema in schema_names:
            # if the schema is part of the cohort name, map the cohort to the schema
            if schema in cohort.lower():
                onedrive_cohort_schema_mapping[cohort] = schema
            # otherwise, look for specific words to help mapping the cohort name to a schema name
            elif "normal" in cohort.lower():
                onedrive_cohort_schema_mapping[cohort] = "normaltissue"
            elif "5k" in cohort.lower():
                onedrive_cohort_schema_mapping[cohort] = "5K_IO"
            elif "h&n" in cohort.lower() or "hn" in cohort.lower():
                onedrive_cohort_schema_mapping[cohort] = "h_and_n"
            elif "glioblastoma" in cohort.lower():
                onedrive_cohort_schema_mapping[cohort] = "giloblastoma"
            elif "seminal_vesicle" in cohort.lower():
                onedrive_cohort_schema_mapping[cohort] = "seminal_vesicle"
            elif "egfr" in cohort.lower():
                onedrive_cohort_schema_mapping[cohort] = "egfr"
            elif "biopsies_and_resections" in cohort.lower():
                onedrive_cohort_schema_mapping[cohort] = "biopsies_resections"
            elif "cholangiocaricinoma" in cohort.lower():
                onedrive_cohort_schema_mapping[cohort] = "cholangiocarcinoma"

    # fetch the schema names from the database
    schema_names = fetch_schema_names()["schema_name"].to_list()
    # Define the path to the json fike
    onedrive_cohort_schema_mapping_path = (
        project_root / "json_files" / "onedrive_cohort_schema_mapping.json"
    )
    # Dump the changes in the json file
    with open(onedrive_cohort_schema_mapping_path, "w") as f:
        json.dump(onedrive_cohort_schema_mapping, f, indent=4)
