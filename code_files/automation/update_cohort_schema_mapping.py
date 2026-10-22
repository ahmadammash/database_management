# import the required libraries
import json

# Import required functions
from code_files.querying.fetch_information import fetch_schema_names
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

# load the cohort name to schema mapping json file
cohort_schema_mapping_path = project_root / "json_files" / "cohort_schema_mapping.json"
with open(cohort_schema_mapping_path, "r") as f:
    cohort_schema_mapping = json.load(f)

# load the cohort names mapping json file
cohort_mapping_path = project_root / "json_files" / "cohort_names_mapping.json"
with open(cohort_mapping_path, "r") as f:
    cohort_mapping = json.load(f)


def update_cohort_schema_mapping():
    """
    This function loops through the cohort names from the cohort mapping json file and the schema names (fetched directly from the database) - then it assigns schema names to each cohort
    either by relying on its name or on a specific case which applies to the cohort name. The function then dumps the changes in the json file.
    """
    # get the schema names from the database
    schema_names = fetch_schema_names()["schema_name"].to_list()

    # loop through the mapped cohort names from the values of the cohort mapping dict
    for cohort in cohort_mapping.values():
        # loop through the schema names
        for schema in schema_names:
            # deal with the cohorts which can be directly mapped to their schema names judging by their names - then exit the loop
            if schema.replace("_", "") in cohort and all(
                x not in cohort
                for x in [
                    "normal",
                    "pembro",
                    "nivo",
                    "biopsies",
                    "resections",
                    "5k",
                    "bladder",
                    "breast",
                    "pancreatic",
                    "longitudinal",
                ]
            ):
                cohort_schema_mapping[cohort] = schema
                break
            # map the specific cohorts to schema names based on what their name includes
            if "postsoc" in cohort:
                cohort_schema_mapping[cohort] = "postsoc"
            elif "longitudinal" in cohort:
                cohort_schema_mapping[cohort] = "longitudinal"
            elif "crc" in cohort:
                cohort_schema_mapping[cohort] = "crc"
            elif "normal" in cohort:
                cohort_schema_mapping[cohort] = "normaltissue"
            elif "pembro" in cohort or "nivo" in cohort:
                cohort_schema_mapping[cohort] = "pembro_nivo"
            elif "5k" in cohort:
                cohort_schema_mapping[cohort] = "5K_IO"
            elif "h&n" in cohort or "hn" in cohort:
                cohort_schema_mapping[cohort] = "h_and_n"
            elif "glioblastoma" in cohort:
                cohort_schema_mapping[cohort] = "giloblastoma"
            elif "seminal vesicle" in cohort:
                cohort_schema_mapping[cohort] = "seminal_vesicle"
            elif "biopsies and resections" in cohort:
                cohort_schema_mapping[cohort] = "biopsies_resections"
            elif "cholangiocaricinoma" in cohort:
                cohort_schema_mapping[cohort] = "cholangiocarcinoma"
            elif "nsclc" in cohort:
                cohort_schema_mapping[cohort] = "nsclc"
            elif "sclc" in cohort:
                cohort_schema_mapping[cohort] = "sclc"
            elif "postsoc" in cohort:
                cohort_schema_mapping[cohort] = "postsoc"
            elif "biopsies" in cohort and "resections" in cohort:
                cohort_schema_mapping[cohort] = "biopsies_resections"
            elif "neuroendocrine" in cohort:
                cohort_schema_mapping[cohort] = "neuroendocrine"
            elif "gallbladder" in cohort:
                cohort_schema_mapping[cohort] = "gallbladder"
            elif "bladder" in cohort:
                cohort_schema_mapping[cohort] = "bladder"
            elif "pancreatic" in cohort:
                cohort_schema_mapping[cohort] = "pancreatic"
            elif "brain" in cohort:
                cohort_schema_mapping[cohort] = "brain"
            elif "breast" in cohort:
                cohort_schema_mapping[cohort] = "breast"
            else:
                cohort_schema_mapping[cohort] = "temp"
    with open(cohort_schema_mapping_path, "w") as f:
        json.dump(cohort_schema_mapping, f, indent=4)
