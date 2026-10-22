# Import the required libraries
import json
import os
import sys
import time
import warnings

import pandas as pd
from sqlalchemy.exc import SQLAlchemyError

from code_files.utilities.io.json_manager import (
    create_column_map,
    deserialize_data_types,
    set_data_types,
)

# Import the required functions
from code_files.utilities.io.setup_db_cnxn import setup_cnxn
from code_files.utilities.processing.cleaners import (
    rename_columns,
    replace_z_s,
    title_case,
    trim_column_names,
    trim_spaces,
)
from code_files.utilities.processing.formatters import (
    process_bit_columns,
    process_date_columns,
    process_dec_columns,
)

# Establish a connection to the database
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()    

# JSON file for the cohort mappings
file_path = project_root / "json_files" / "cohort_names_mapping.json"

with open(file_path, "r") as f:
    cohort_mapping = json.load(f)

# Suppress the python warning regarding the downcasting behaviour in the replace function
pd.set_option("future.no_silent_downcasting", True)
warnings.filterwarnings(
    "ignore",
    category=FutureWarning,
    message="Downcasting behavior in `replace` is deprecated",
)

# Define the settings file path
settings_file_path = project_root / "json_files" / "file_upload_settings.json"

# get the column name map
column_name_map = create_column_map()

# determine where the script is running
if "AmmashA.IA" in str(current_dir):
    server = "mssql"
elif "medaamm" in str(current_dir) or "ahmad" in str(current_dir):
    server = "gcp"

# -----------------------------------------------------------------------------------------------------------------------------------------------------#

# Define the number of maximum retries for reading the file from the configuration JSON file
# Define the delay between retries
MAX_RETRIES = 5
RETRY_DELAY = 2  # seconds

# define the separator line
separator = "\n" + "=" * 156 + "\n"

# -----------------------------------------------------------------------------------------------------------------------------------------------------#


# Function to run the upload process
def run_upload_process_supp_anon_unanon(config, engine):
    """
    This function runs the upload process for the anonymised and unanonymised supplementary data files by applying the processing steps
    using the configuration information given to the function as an input argument.
    """
    # define the global ID column to distinguish between different dataset types
    global id_column

    # Get the Settings from the configuration file
    file_path = config["local_file_path"]
    sheet_name = config["sheet_name"]
    sql_table_name = config["sql_table_name"]
    header_row_index = config["header_row_index"]
    schema_name = config["schema_name"]
    cohort_name = config["cohort_name"]
    data_types = config["data_types_serialized"]
    column_name_mapping = config["column_name_mapping"]
    dataset_type = config["dataset_type"]
    upload_destination = config["upload_destination"]

    # if running from ahmad's pc, modify the file path
    if "ahmad" in str(current_dir):
        file_path = file_path.replace("AmmashA.IA", "ahmad")

    # set the value of the id_column variable based on the dataset type from the configuration file
    if dataset_type == "anonymised supplementary":
        id_column = "DonorID"
    else:
        df_temp = pd.read_excel(file_path, sheet_name)
        id_column = "DonorID" if "DonorID" in df_temp.columns else "NHS_Number"

    # Normalise the file path
    file_path = os.path.normpath(file_path)

    # if the script is running from the UOL PC, modify the file_path variable
    if "medaamm" in str(current_dir):
        file_path = file_path.replace("AmmashA.IA", "medaamm")

    # if the dataset should only be uploaded to mssql (i.e. scan logs or unanonymised data) and the script is running from the UOL PC, skip the table by exiting the function (before starting the upload)
    if upload_destination == "mssql" and "medaamm" in str(current_dir):
        return True

    # print a message saying the upload started for the specific file path
    print(f"✅ Starting upload for: {file_path}")

    # Initialise the main_workbook variable as None
    main_workbook = None

    # Retry the read-in of the file within the maximum retries range
    for attempt in range(MAX_RETRIES):
        # Extract and Process the data
        try:
            # Read in the updated worksheet
            main_workbook = pd.read_excel(
                str(file_path), sheet_name=sheet_name, header=header_row_index
            )

            # if read-in successful, break the loop and exit
            print(f"✅ File read successfully on attempt {attempt + 1}.")
            break
        except (OSError, ValueError, TypeError, ImportError) as e:
            last_exception = e
            print(f"❌ Attempt {attempt + 1} failed to read the file")
            # if the last retry is reached and fails, exit and indicate that the read-in failed
            if attempt == MAX_RETRIES - 1:
                sys.stderr.write(separator)
                sys.stderr.write(
                    f"❌ CRITICAL ERROR: Max retries reached for file: {file_path}\n"
                )
                sys.stderr.write(
                    f"🔎 REASON: {type(last_exception).__name__}: {last_exception!s}\n"
                )
                sys.stderr.write("📝 Max retries reached. File read failed")
                sys.stderr.write(separator)
                return False
        time.sleep(RETRY_DELAY)
    if main_workbook is None:
        return False
    try:
        # Apply the processing steps
        main_workbook = trim_column_names(main_workbook)
        main_workbook = trim_spaces(main_workbook)
        main_workbook = replace_z_s(main_workbook)
        main_workbook = title_case(main_workbook)
        main_workbook = rename_columns(main_workbook, column_name_mapping)
        main_workbook = process_date_columns(main_workbook, dataset_type)
        main_workbook = process_dec_columns(main_workbook)
        main_workbook = process_bit_columns(main_workbook)

        # print a message saying processing is successful
        print("✅ Data processing succeeded.")
    # if the processing steps fail, return a message accordingly
    except (
        AttributeError,
        TypeError,
        ValueError,
        KeyError,
        OSError,
        ImportError,
        pd.errors.EmptyDataError,
        pd.errors.ParserError,
    ) as e:
        last_exception = e
        # print the error to the log file
        sys.stderr.write(separator)
        sys.stderr.write(
            f"❌ CRITICAL ERROR: Data processing failed for file: {file_path}"
        )
        sys.stderr.write(
            f"🔎 REASON: {type(last_exception).__name__}: {last_exception!s}\n"
        )
        sys.stderr.write("📝 Data processing steps failed.")
        sys.stderr.write(separator)
        return False

    # Add the cohort name column
    main_workbook = main_workbook.assign(
        CohortName=cohort_mapping.get(cohort_name, cohort_name)
    )

    # Print a message saying the cohort name column was added
    print("✅ Cohort Name column added.")

    # re-order the dataframe columns to place the cohort name column at the start
    main_workbook = main_workbook[
        ["CohortName"]
        + [col for col in main_workbook.columns if col not in ["CohortName"]]
    ]

    # Upload to SQL Server
    try:
        with engine.begin() as conn:
            # update the data types variable by running the set_data_types function
            data_types = set_data_types(main_workbook, dataset_type)

            main_workbook.to_sql(
                sql_table_name,
                conn,
                schema=schema_name,
                if_exists="replace",
                dtype=deserialize_data_types(data_types),
                index=False,
            )
            print(
                f"✅ Dataframe '{sql_table_name}' successfully uploaded to schema '{schema_name}'."
            )
            print(separator)
        return True

    # if the upload to SQL fails, return a message accordingly
    except (SQLAlchemyError, TypeError, ValueError, AttributeError, OSError) as e:
        last_exception = e
        # print the error to the log file
        sys.stderr.write(separator)
        sys.stderr.write(f"❌ CRITICAL ERROR: Upload failed for file: {file_path}")
        sys.stderr.write(
            f"🔎 REASON: {type(last_exception).__name__}: {last_exception!s}\n"
        )
        sys.stderr.write("📝 Upload failed.")
        sys.stderr.write(separator)
        return False


# -----------------------------------------------------------------------------------------------------------------------------------------------------#
# %%
