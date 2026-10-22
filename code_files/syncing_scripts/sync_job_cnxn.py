# Import the required libraries
import sys
import time
import warnings
from datetime import datetime
from pathlib import Path

import pandas as pd

# Suppress the python warning regarding the downcasting behaviour in the replace function
pd.set_option("future.no_silent_downcasting", True)
warnings.filterwarnings(
    "ignore",
    category=FutureWarning,
    message="Downcasting behavior in `replace` is deprecated",
)

current_dir = Path(__file__).resolve()
project_root = next(
    (p for p in Path(".").resolve().parents if (p / "pyproject.toml").exists()),
    Path(".").resolve(),
)
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

# move to the code_files directory
code_files = project_root / "code_files"

# add the root path to the system
if str(code_files) not in sys.path:
    sys.path.insert(0, str(code_files))

from etl_logic.etl_logic_box_cnxn import run_upload_process_box_cnxn
from etl_logic.etl_logic_onedrive_cnxn import run_upload_process_onedrive_cnxn
from utilities.io.database import save_all_settings
from utilities.io.json_manager import load_settings

# -----------------------------------------------------------------------------------------------------------------------------------------------------#


# Define the batch upload function
def start_batch_upload():
    # define the configuration settings
    # to run the ETL process for every uploaded table
    print("Loading Configuration...")
    mapping_file = "clinical_metadata_mapping.json"
    config, settings_file_path = load_settings(mapping_file)

    # display a message if the configuration file failed to load
    if not config:
        sys.stderr.write(
            "\n===================================================================================================================\n"
        )
        sys.stderr.write(
            "❌ CRITICAL ERROR: Configuration file is missing or has failed to load.\n"
        )
        sys.stderr.write(
            "\n===================================================================================================================\n"
        )
        sys.exit(1)  # exit with an error code
        return

    # define the number of datasets already on the database and the number of those that have been successfully processed
    upload_count = 0
    success_count = 0

    print(
        f"Starting the batch re-upload process for {len(config)} configured datasets..."
    )
    print("-" * 50)

    for table_name, cohort_config in config.items():
        # check if the configuration associated with that table is valid
        if not isinstance(cohort_config, dict):
            print(f"Skipping key '{table_name}': Invalid configuration format.")
            print("-" * 75)
            upload_count += 1
            continue
        upload_count += 1

        # process the table if the configuration is valid
        print(f"\n[{upload_count}/{len(config)}] Processing table: **{table_name}**")
        file_path = cohort_config.get("local_file_path", "N/A")

        # get the cohort name
        cohort_name = cohort_config["cohort_name"]

        # add a short delay
        time.sleep(2)

        # run the upload process depending on the connection type
        try:
            if cohort_config["connection_type"] == "onedrive_cnxn":
                run_upload_process_onedrive_cnxn(cohort_config, cohort_name)
                cohort_config["last_upload_timestamp"] = datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                )  # update the last upload timestamp in the config
                success_count += 1

            elif cohort_config["connection_type"] == "box_cnxn":
                run_upload_process_box_cnxn(cohort_config, cohort_name)
                cohort_config["last_upload_timestamp"] = datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                )  # update the last upload timestamp in the config
                success_count += 1

            # if the upload process runs successfully, update the success count and print a success message
            print(f"✅ Successfully processed and uploaded dataset for {cohort_name}.")

        except Exception as e:
            # write the error to the log file
            last_exception = e
            sys.stderr.write(
                "\n========================================================================================================================================================\n"
            )
            sys.stderr.write(f"\n❌ ERROR PROCESSING TABLE: {cohort_name} ---\n")
            sys.stderr.write(
                f"🔎 REASON: {type(last_exception).__name__}: {last_exception!s}\n"
            )
            sys.stderr.write("📝 Error in processing dataset.")
            sys.stderr.write(
                "\n========================================================================================================================================================\n"
            )

    save_all_settings(config, settings_file_path)
    print("Configuration file updated with new timestamps.")

    if success_count < upload_count:
        print(f"\nCompleted with {success_count}/{upload_count} successful uploads.")
        sys.exit(1)

    else:
        print(f"\nAll {upload_count} uploads completed successfully.")
        sys.exit(0)


if __name__ == "__main__":
    start_batch_upload()
