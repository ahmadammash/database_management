# Import the required libraries
import os
import sys
import time
import warnings

import pandas as pd

from code_files.automation.anonymiser import anonymiser

# Suppress the python warning regarding the downcasting behaviour in the replace function
pd.set_option("future.no_silent_downcasting", True)
warnings.filterwarnings(
    "ignore",
    category=FutureWarning,
    message="Downcasting behavior in `replace` is deprecated",
)

# disable warning errors
warnings.filterwarnings("ignore", category=UserWarning, message=".*")
warnings.filterwarnings("ignore", category=FutureWarning, message=".*")

# -----------------------------------------------------------------------------------------------------------------------------------------------------#

# Define the number of maximum retries for reading the file from the configuration JSON file
# Define the delay between retries
MAX_RETRIES = 5
RETRY_DELAY = 2  # seconds

# define the separator line
separator = "\n" + "=" * 156 + "\n"

# -----------------------------------------------------------------------------------------------------------------------------------------------------#


# Function to run the upload process
def run_upload_process_onedrive_cnxn(cohort_config, cohort_name):
    """
    This function runs the upload process for unanonymised data from NHS OneDrive to the NHS OneDrive Box uploads folder. It anonymises the
    data using the anonymiser function before moving it to the Box Uploads folder.
    """
    # Get the Settings from the configuration file
    onedrivebox_filepath = cohort_config["onedrivebox_filepath"]
    onedrive_filepath = cohort_config["onedrive_filepath"]
    onedrivebox_filename = cohort_config["onedrivebox_filename"]
    onedrive_filename = cohort_config["onedrive_filename"]
    onedrivebox_sheet_name = cohort_config["onedrivebox_sheet_name"]
    onedrive_sheet_name = cohort_config["onedrive_sheet_name"]
    onedrivebox_header_number = cohort_config["onedrivebox_header_number"]
    onedrive_header_number = cohort_config["onedrive_header_number"]
    dataset_type = cohort_config["dataset_type"]
    last_upload_timestamp = cohort_config["last_upload_timestamp"]

    # Normalise the OneDrive file path
    onedrivebox_filepath = os.path.normpath(onedrivebox_filepath)

    # print a message saying the anonymisation and upload started for the specific file path
    print(f"✅ Starting the anonymisation and upload process for {onedrive_filepath}")

    # Initialise the main_workbook variable to None
    main_workbook = None

    # Retry the read-in of the file within the maximum retries range
    for attempt in range(MAX_RETRIES):
        # Extract and Process the data
        try:
            # Read in the updated worksheet
            main_workbook = pd.read_excel(
                str(onedrive_filepath),
                sheet_name=onedrive_sheet_name,
                header=onedrive_header_number,
                dtype=object,
            )
            # if read-in successful, break the loop and exit
            print(f"✅ File read successfully on attempt {attempt + 1}.")
            break
        except Exception as e:
            last_exception = e
            print(f"❌ Attempt {attempt + 1} failed to read the file")
            # if the last retry is reached and fails, exit and indicate that the read-in failed
            if attempt == MAX_RETRIES - 1:
                sys.stderr.write(separator)
                sys.stderr.write(
                    f"❌ CRITICAL ERROR: Max retries reached for file: {onedrive_filepath}\n"
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
        # run the anonymisation
        if dataset_type in ["anonymised prelim", "anonymised follow up"] or dataset_type == "anonymised supplementary":
            main_workbook_anonymised = anonymiser(
                onedrivebox_sheet_name, main_workbook, cohort_name, onedrivebox_filename
            )

        print(
            f"✅ Anonymisation and upload process completed successfully for {onedrive_filepath}"
        )

    except Exception as e:
        last_exception = e
        sys.stderr.write(separator)
        sys.stderr.write(
            f"❌ CRITICAL ERROR: An error occurred during the anonymisation and upload process for file: {onedrive_filepath}\n"
        )
        sys.stderr.write(
            f"🔎 REASON: {type(last_exception).__name__}: {last_exception!s}\n"
        )
        sys.stderr.write("📝 Anonymisation and upload process failed")
        sys.stderr.write(separator)
        return False
