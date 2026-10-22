# Import the required libraries
import os
import sys
import time
import warnings
from pathlib import Path

import openpyxl
import pandas as pd

# Import the required functions
from code_files.utilities.config.mappings import set_data_types_scan_logs
from code_files.utilities.io.json_manager import deserialize_data_types
from code_files.utilities.io.setup_db_cnxn import setup_cnxn
from code_files.utilities.processing.cleaners import (
    analyse_cell_colors,
    remove_empty_and_complete,
    rename_columns,
    title_case,
    trim_column_names,
    trim_spaces,
)

# Suppress the python warning regarding the downcasting behaviour in the replace function
pd.set_option("future.no_silent_downcasting", True)
warnings.filterwarnings(
    "ignore",
    category=FutureWarning,
    message="Downcasting behavior in `replace` is deprecated",
)

# get the current directory
current_dir = Path(__file__).resolve()

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

# -----------------------------------------------------------------------------------------------------------------------------------------------------#

# Define the number of maximum retries for reading the file from the configuration JSON file
# Define the delay between retries
MAX_RETRIES = 5
RETRY_DELAY = 2  # seconds
# define the separator line
separator = "\n" + "=" * 156 + "\n"
# -----------------------------------------------------------------------------------------------------------------------------------------------------#


# Functiion to run the upload process
def run_upload_process_scan_logs(config, engine):
    """
    This function runs the upload process for scan logs datasets after processing them using the configuration information associated with
    each table and given to the function as an input argument.
    """
    # Get the settings from the configuration file
    file_path = config["local_file_path"]
    sheet_name = config["sheet_name"]
    sql_table_name = config["sql_table_name"]
    header_row_index = config["header_row_index"]
    schema_name = config["schema_name"]
    data_types = config["data_types_serialized"]
    column_name_mapping = config["column_name_mapping"]
    dataset_type = config["dataset_type"]
    last_timestamp = config["last_upload_timestamp"]
    upload_destination = config["upload_destination"]

    # Normalise the file path
    file_path = os.path.normpath(file_path)

    # ------------------------------------------------------------------------------

    wb = openpyxl.load_workbook(file_path, data_only=True)

    # Try to read the file with retries
    for attempt in range(1, MAX_RETRIES + 1):
        # Extract and Process the data
        try:
            # Read in the updated worksheet
            main_workbook = pd.read_excel(
                str(file_path), sheet_name=sheet_name, header=header_row_index
            )
            # if read-in successful, break the loop and exit
            print(f"✅ File read successfully on attempt {attempt}.")
            break
        except Exception as e:
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

    main_workbook = analyse_cell_colors(
        filename_main=file_path, selected_sheet=sheet_name, header_row=header_row_index
    )

    # ------------------------------------------------------------------------------

    # if the dataset should only be uploaded to mssql (i.e. scan logs or unanonymised data) and the script is running from the UOL PC, skip the table by exiting the function (before starting the upload)
    if upload_destination == "mssql" and "medaamm" in str(current_dir):
        return True

    # print a message saying the upload started for the specific file path
    print(f"✅ Starting upload for: {file_path}")

    try:
        # Apply the processing steps
        main_workbook = trim_column_names(main_workbook)
        main_workbook = trim_spaces(main_workbook)
        main_workbook = title_case(main_workbook)
        main_workbook = rename_columns(main_workbook, column_name_mapping)
        main_workbook = remove_empty_and_complete(main_workbook)

        # print a message saying processing is successful
        print("✅ Data processing succeeeded.")

    # if the processing steps fail, return a message accordingly
    except Exception as e:
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

    # Upload to SQL Server
    try:
        with engine.begin() as conn:
            # update the data types variable by running the set_data_types_scan_logs function
            data_types = set_data_types_scan_logs(main_workbook)

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
    except Exception as e:
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
