import json
import os
from datetime import date, datetime

import pandas as pd
from sqlalchemy.types import (
    CHAR,
    DATE,
    DECIMAL,
    INTEGER,
    VARCHAR,
    Boolean,
    SmallInteger,
)

# import the required functions
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

from code_files.utilities.config.mappings import create_column_map, set_data_types

# JSON file for the cohort_mapping
json_path = project_root / "json_files" / "cohort_names_mapping.json"
with open(json_path, "r") as f:
    cohort_mapping = json.load(f)

# JSON file for the onedrive connection
json_mapping_path = project_root / "json_files" / "clinical_metadata_mapping.json"

# JSON file for the patient list connection
json_mapping_path_patientList = project_root / "json_files" / "patientList_config.json"

# JSON file for the cohort name to cohort schema mapping
cohort_schema_mapping_path = project_root / "json_files" / "cohort_schema_mapping.json"

# Import required functions

# ----------------------------------------------------------------------------------------------------------------------------------------


def load_settings(settings_file_name):
    """
    This function loads the information from the file upload settings JSON file
    """
    settings_file_path = project_root / "json_files" / settings_file_name
    try:
        with open(settings_file_path, "r") as f:
            return json.load(f), settings_file_path
    # if the read-in fails, return a message accordingly
    except (FileNotFoundError, json.JSONDecodeError):
        print("The configuration file is not found or is corrupted.")
        return None, settings_file_path


# ----------------------------------------------------------------------------------------------------------------------------------------


def save_settings_generic(
    settings_file_path,
    file_path,
    sheet_name,
    sql_table_name,
    header_index,
    schema_name,
    data_types,
    column_names_mapping,
    dataset_type,
    upload_destination,
    **kwargs,
):
    """
    This function saves the config settings of a given table on the database
    """
    # Handle special date conversion (where needed) - unanonymised data will include actual dates
    for key, value in data_types.items():
        if isinstance(value, (pd.Timestamp, datetime, date)):
            data_types[key] = value.isoformat()

    # Serialize the types or objects to strings so they can be accepted in the JSON structure
    serializable_data_types = {k: str(v) for k, v in data_types.items()}

    # create the new settings in the JSON file
    new_settings = {
        "local_file_path": file_path,
        "sheet_name": sheet_name,
        "sql_table_name": sql_table_name,
        "header_row_index": header_index,
        "schema_name": schema_name,
        "data_types_serialized": serializable_data_types,
        "column_name_mapping": column_names_mapping,
        "dataset_type": dataset_type,
        "last_upload_timestamp": None,
        "upload_destination": upload_destination,
    }

    # Add optional fields from kwargs (i.e., cohort_name where applicable)
    if "cohort_name" in kwargs:
        new_settings["cohort_name"] = kwargs["cohort_name"]

    # Load existing settings safely
    existing_data = {}
    if os.path.exists(settings_file_path):
        try:
            with open(settings_file_path, "r") as f:
                loaded_data = json.load(f)
                if type(loaded_data) is dict:
                    existing_data = loaded_data
        except json.JSONDecodeError:
            print(f"Warning: Settings file at {settings_file_path} is corrupted.")

    # Update/Insert the entry and save
    existing_data[sql_table_name] = new_settings
    with open(settings_file_path, "w") as f:
        json.dump(existing_data, f, indent=4)

    return True


# ----------------------------------------------------------------------------------------------------------------------------------------


def deserialize_data_types(data_types_serialized):
    """
    This function deserializes data types to enable retrieval from the JSON file and re-use for uploading the database
    """
    dtype_mapping = {}
    for col, dtype in data_types_serialized.items():
        dtype_str = str(dtype).upper().strip()
        if dtype_str.startswith("VARCHAR"):
            start = dtype_str.find("(") + 1
            end = dtype_str.find(")")
            size = int(dtype_str[start:end])
            dtype_mapping[col] = VARCHAR(size)

        elif dtype_str.startswith("CHAR"):
            start = dtype_str.find("(") + 1
            end = dtype_str.find(")")
            size = int(dtype_str[start:end])
            dtype_mapping[col] = CHAR(size)

        elif "BIT" in dtype_str:
            dtype_mapping[col] = Boolean

        elif "TINYINT" in dtype_str:
            dtype_mapping[col] = SmallInteger

        elif dtype_str.startswith("DECIMAL"):
            content = dtype_str[dtype_str.find("(") + 1 : dtype_str.find(")")].split(
                ","
            )
            precision = int(content[0].strip())
            scale = 0
            if len(content) > 1:
                scale = int(content[1].strip())
            dtype_mapping[col] = DECIMAL(precision, scale)
        elif "DATE" in dtype_str:
            dtype_mapping[col] = DATE

        elif "INTEGER" in dtype_str:
            dtype_mapping[col] = INTEGER

    return dtype_mapping


# ----------------------------------------------------------------------------------------------------------------------------------------


def format_json():
    """
    This function re-formats the file upload settings JSON file and runs after every batch upload.
    """
    settings_file_name = project_root / "json_files" / "file_upload_settings.json"
    json_file, file_path = load_settings(settings_file_name)

    column_name_map_scans = {
        "Lh Number": "LH_Number",
        "Part Number": "Part_Number",
        "Anonymised Code": "Anonymised_Code",
        "Svs Number To Transfer": "Svs_Number_To_Transfer",
        "Scan 1 Qc": "Scan1_QC",
        "Scan 2 Qc": "Scan2_QC",
        "Error Code": "ErrorCode_Scan1",
        "Error Code.1": "ErrorCode_Scan2",
        "Notes": "Notes_Scan1",
        "Notes.1": "Notes_Scan2",
    }

    for cohort, config in json_file.items():
        # re-define the column name mapping in the json file
        if config["dataset_type"] == "scan logs":
            json_file[cohort]["column_name_mapping"] = column_name_map_scans
        else:
            config["column_name_mapping"] = create_column_map()

        # re-evaluate the cohort name for each dataset using the filepath to cohortname mapping (for box files) and onedrive tbl to cohort mapping (for onedrive files)
        # define the required file paths
        filepath_cohortname_mapping_path = (
            project_root / "json_files" / "filepath_cohortname_mapping.json"
        )

        # Load mappings
        with open(filepath_cohortname_mapping_path, "r") as f:
            filepath_cohortname_mapping = json.load(f)

        with open(cohort_schema_mapping_path, "r") as f:
            cohort_schema_mapping = json.load(f)

        # Temporary version of local_file_path for comparison logic
        current_path = config["local_file_path"]

        # extract the last part of the file path
        name = current_path.split("\\")[-1]

        # Create temp path replacing 'Normal' with 'Normal tissue' if needed
        # and replacing __ with _ if needed
        if "normal" in name.lower() and "tissue" not in name.lower():
            temp_path_for_matching = current_path.replace("Normal", "Normal tissue")
        else:
            temp_path_for_matching = current_path.replace("__", "_").replace(
                "Cholangiocaricinomas", "Cholangiocarcinoma"
            )

        # Loop through the filepath to cohort name mapping
        for filepath, cohortname in filepath_cohortname_mapping.items():
            if filepath in temp_path_for_matching:
                config["cohort_name"] = cohortname

        # Loop through the onedrive table to cohort name mapping
        for onedrive_tbl, cohortname in cohort_mapping.items():
            if onedrive_tbl in temp_path_for_matching:
                config["cohort_name"] = cohortname

        # Set the schema name for each table based on the cohort name
        try:
            config["schema_name"] = cohort_schema_mapping[config["cohort_name"]]
        except KeyError:
            config["schema_name"] = "temp"
            print(config["cohort_name"])

        # import the dataframe
        try:
            df = pd.read_excel(
                config["local_file_path"],
                sheet_name=config["sheet_name"],
                header=config["header_row_index"],
            )
            data_types = set_data_types(df, config["dataset_type"])
            serializable_data_types = {k: str(v) for k, v in data_types.items()}
            config["data_types_serialized"] = serializable_data_types
            config["last_upload_timestamp"] = datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            )
        except (
            FileNotFoundError,
            OSError,
            ValueError,
            TypeError,
            pd.errors.EmptyDataError,
            pd.errors.ParserError,
        ) as e:
            print(f"Error processing cohort {cohort}: {e}")
            continue

    with open(file_path, "w") as f:
        json.dump(json_file, f, indent=4)


def add_to_json_onedrive_cnxn(
    onedrivebox_filepath,
    onedrivebox_sheet_name,
    onedrivebox_header_row,
    onedrivebox_filename,
    onedrive_filepath,
    onedrive_sheet_name,
    onedrive_header_row,
    onedrive_filename,
    selected_cohort,
    dataset_type,
):
    """
    This function adds the configuration data associated with a dataset to the onedrive connection JSON file
    """

    data = {
        "onedrivebox_filepath": onedrivebox_filepath,
        "onedrive_filepath": onedrive_filepath,
        "onedrivebox_filename": onedrivebox_filename,
        "onedrive_filename": onedrive_filename,
        "onedrivebox_sheet_name": onedrivebox_sheet_name,
        "onedrive_sheet_name": onedrive_sheet_name,
        "onedrivebox_header_number": onedrivebox_header_row,
        "onedrive_header_number": onedrive_header_row,
        "cohort_name": cohort_mapping.get(
            selected_cohort["value"], selected_cohort["value"]
        ),
        "dataset_type": dataset_type,
        "last_upload_timestamp": None,
        "connection_type": "onedrive_cnxn",
    }

    # Load the existing data from the JSON file
    existing_data = {}

    # create a dictionary to map the dataset type to the suffix added to the table name in the json file
    suffix_mapping = {
        "anonymised prelim": "_Prelim",
        "anonymised follow up": "_FollowUp",
        "anonymised supplementary": "_Supplementary",
    }

    if os.path.exists(json_mapping_path):
        try:
            with open(json_mapping_path, "r") as json_file:
                loaded_data = json.load(json_file)

                # only use the loaded data if it's a dictionary
                if type(loaded_data) is dict:
                    existing_data = loaded_data
        except json.JSONDecodeError:
            # Give a warning if the JSON file is corrupted or empty
            print(
                "Warning: The JSON file is corrupted or empty. A new file will be created."
            )

    # Update the relevant settings in the JSON file but keep the other settings (for the associate onedrive connection unchanged)
    entry_key = f"{selected_cohort['value']}{suffix_mapping.get(dataset_type, '')}_onedrive_cnxn"

    # If the key already exists in the settings file, update the relevant information
    if entry_key in existing_data:
        existing_data[entry_key].update(data)
    else:
        existing_data[entry_key] = data

    # Save back to the file
    with open(json_mapping_path, "w") as json_file:
        json.dump(existing_data, json_file, indent=4)

    return True


def add_to_json_box_cnxn(
    box_filepath,
    box_sheet_name,
    box_header_row,
    box_filename,
    onedrive_filepath,
    onedrive_sheet_name,
    onedrive_header_row,
    onedrive_filename,
    selected_cohort,
    dataset_type
):
    """
    This function adds the configuration information of a dataset to the box cnxn JSON file
    """
    data = {
        "box_filepath": box_filepath,
        "onedrivebox_filepath": onedrive_filepath,
        "box_filename": box_filename,
        "onedrivebox_filename": onedrive_filename,
        "box_sheet_name": box_sheet_name,
        "onedrivebox_sheet_name": onedrive_sheet_name,
        "box_header_number": box_header_row,
        "onedrivebox_header_number": onedrive_header_row,
        "cohort_name": cohort_mapping.get(
            selected_cohort["value"], selected_cohort["value"]
        ),
        "dataset_type": dataset_type,
        "last_upload_timestamp": None,
        "connection_type": "box_cnxn",
    }

    # Load the existing data from the JSON file
    existing_data = {}

    # create a dictionary to map the dataset type to the suffix added to the table name in the json file
    suffix_mapping = {
        "anonymised prelim": "_Prelim",
        "anonymised follow up": "_FollowUp",
        "anonymised supplementary": "_Supplementary",
    }

    if os.path.exists(json_mapping_path):
        try:
            with open(json_mapping_path, "r") as json_file:
                loaded_data = json.load(json_file)

                # only use the loaded data if it's a dictionary
                if type(loaded_data) is dict:
                    existing_data = loaded_data
        except json.JSONDecodeError:
            # Give a warning if the JSON file is corrupted or empty
            print(
                "Warning: The JSON file is corrupted or empty. A new file will be created."
            )

    # Update the relevant settings in the JSON file but keep the other settings (for the associate onedrive connection unchanged)
    entry_key = (
        f"{selected_cohort['value']}{suffix_mapping.get(dataset_type, '')}_box_cnxn"
    )

    # If the key already exists in the settings file, update the relevant information
    if entry_key in existing_data:
        existing_data[entry_key].update(data)
    else:
        existing_data[entry_key] = data

    # Save back to the file
    with open(json_mapping_path, "w") as json_file:
        json.dump(existing_data, json_file, indent=4)

    return True


def add_to_json_patientList(file_path, selected_sheet, df_name):
    """
    This function adds the configuration information of the patient list to the associated JSON file
    """
    data = {
        "file_path": file_path,
        "sheet_name": selected_sheet,
        "df_name": df_name,
        "last_upload_timestamp": None,
    }
    # Load the existing data from the JSON file
    existing_data = {}

    if os.path.exists(json_mapping_path_patientList):
        try:
            with open(json_mapping_path_patientList, "r") as json_file:
                loaded_data = json.load(json_file)

                # only use the loaded data if it's a dictionary
                if type(loaded_data) is dict:
                    existing_data = loaded_data
        except json.JSONDecodeError:
            # Give a warning if the JSON file is corrupted or empty
            print(
                "Warning: The JSON file is corrupted or empty. A new file will be created."
            )

    # Add/update the new settings
    # This correctly adds a new entry if sql_table_name is new, or replaces it if it already exists
    existing_data["PatientList"] = data

    # Save back to the file
    with open(json_mapping_path_patientList, "w") as json_file:
        json.dump(existing_data, json_file, indent=4)

    return True


# %%
