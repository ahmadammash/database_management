# Import the required libraries
import json
from datetime import datetime

from code_files.utilities.io.json_manager import load_settings

# Import the required functions
from code_files.utilities.io.setup_db_cnxn import setup_cnxn
from code_files.utilities.pipelines.upload_pipeline import import_worksheet_patientList

# setup a database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()


# Define the upload process for the patient list
def run_upload_process_patientList():
    """
    This function runs the upload process for the patient list from the NHS OneDrive to the database using the configuration settings stored
    in patientList_config.json.
    """
    # Define the patient list configuration file
    config_file = "patientList_config.json"
    data, file_path = load_settings(config_file)

    # Read in the config settings from the json file
    try:
        with open(file_path, "r") as file:
            config = data["PatientList"]
            path = config["file_path"]
            sheet_name = config["sheet_name"]
            df_name = config["df_name"]

            # Process and upload the data to the database
            df = import_worksheet_patientList(path, sheet_name, 1, df_name, engine)

        return True
    except Exception:
        return False


# Define the function to start the batch upload procedure
def start_batch_upload_patientList():
    run_upload_process_patientList()

    # Define the patient list configuration file
    config_file = "patientList_config.json"
    data, file_path = load_settings(config_file)

    # Define the new settings for the JSON file with the updated timestamp
    new_settings = {
        "file_path": "D:/AmmashA.IA/OneDrive - NHS/TrisStar metadata/Important documents/Patient List.xlsm",
        "sheet_name": "Patient inventory",
        "df_name": "PatientList",
        "last_upload_timestamp": datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
    }

    # Update the JSON file
    data["PatientList"] = new_settings
    with open(file_path, "w") as file:
        # Update the last upload timestamp
        json.dump(data, file, indent=4)


if __name__ == "__main__":
    start_batch_upload_patientList()
