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

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()
json_files_dir = project_root / "json_files"

# JSON file for the cohort mappings
json_path = project_root / "json_files" / "cohort_names_mapping.json"
with open(json_path, "r") as f:
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

# Get the column name map
column_name_map = create_column_map()

# determine where the script is running
if "AmmashA.IA" in str(current_dir):
    server = "mssql"
elif "medaamm" in str(current_dir):
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
def run_upload_process_prelim(config, engine):
    """
    This function runs the upload process for prelim data after processing the data using the configuration information provided
    as an input argument from the file upload settings JSON file.
    """
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

    # Normalise the file path
    file_path = os.path.normpath(file_path)

    # if the script is running from the UOL PC, modify the file_path variable
    if server == "gcp":
        file_path = file_path.replace("AmmashA.IA", "medaamm")

    # if the dataset should only be uploaded to mssql (i.e. scan logs or unanonymised data) and the script is running from the UOL PC, skip the table by exiting the function (before starting the upload)
    if upload_destination == "mssql" and server == "gcp":
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
        except (
            FileNotFoundError,
            OSError,
            ValueError,
            TypeError,
            ImportError,
            pd.errors.ParserError,
        ) as e:
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

        # Convert the date columns to date type
        if "unanonymised" in dataset_type:
            for col in main_workbook.columns:
                if ("Date" in col or "DOB" in col or "Dob" in col) and (
                    "TTE" not in col
                ):
                    main_workbook[col] = pd.to_datetime(
                        main_workbook[col], errors="coerce"
                    ).dt.strftime("%Y-%m-%d")
                    # Handle NaT values
                    main_workbook[col] = main_workbook[col].where(
                        main_workbook[col].notnull(), None
                    )

        # print a message saying processing is successful
        print("✅ Data processing succeeded.")
    # if the processing steps fail, return a message accordingly
    except (
        ValueError,
        TypeError,
        KeyError,
        IndexError,
        AttributeError,
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
    main_workbook = main_workbook.copy()
    main_workbook["CohortName"] = cohort_mapping.get(cohort_name, cohort_name)

    # Replace N/A values with actual nulls in the ID columns
    id_cols_to_check = [
        "SelSpecAnonymisedID",
        "AnonymisedID_MatchedMet1",
        "AnonymisedID_NormalTissue",
        "SelBiopsyAnonymisedID",
        "ResectionAnonymisedID",
    ]

    block_cols_to_check = [
        "SpecBlockNum1",
        "MatchedMet1_BlockNumber",
        "MatchedNormalTissue_BlockNumber",
        "LowGrade_Dysplasia_Block1",
        "LowGrade_Dysplasia_Block2",
        "HighGrade_Dysplasia_Block1",
        "HighGrade_Dysplasia_Block2",
        "FocalHighGrade_Dysplasia_Block1",
        "FocalHighGrade_Dysplasia_Block2",
        "BiopsyBlockNum1",
        "BiopsyBlockNum2",
        "ResectionBlockNum1",
        "ResectionBlockNum2",
        "SpecBlockNum",
    ]

    # Filter the lists of cols to include those that are in the main_workbook
    existing_id_cols = [col for col in id_cols_to_check if col in main_workbook.columns]
    existing_block_cols = [
        col for col in block_cols_to_check if col in main_workbook.columns
    ]

    for col in existing_id_cols + existing_block_cols:
        main_workbook[col] = (
            main_workbook[col]
            .astype(str)
            .replace(["N/A", "n/a", "NA", "nan", "None"], pd.NA)
        )

    if dataset_type in [
        "anonymised follow up",
        "anonymised prelim",
        "unanonymised follow up",
        "unanonymised prelim",
    ]:
        # Work with the ID columns
        if existing_id_cols:
            ids = main_workbook[existing_id_cols].bfill(axis=1).iloc[:, 0]
        else:
            ids = pd.Series("UnknownID", index=main_workbook.index)

        # Work with the block numbers
        if existing_block_cols:
            blocks = main_workbook[existing_block_cols].bfill(axis=1).iloc[:, 0]
        else:
            blocks = pd.Series("UnknownBlock", index=main_workbook.index)

        # Clean up the ids and block numbers
        ids = ids.astype(str).replace(["<NA>", "nan"], "UnknownID")
        blocks = blocks.astype(str).replace(["<NA>", "nan"], "UnknownBlock")

    # re-order the dataframe columns to place the cohort name at the start
    main_workbook = main_workbook[
        ["CohortName"] + [col for col in main_workbook.columns if col != "CohortName"]
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
    except (SQLAlchemyError, ValueError, TypeError) as e:
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
