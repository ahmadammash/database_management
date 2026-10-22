# Import the required libraries
import sys
import warnings
from datetime import datetime

# Suppress the python warning regarding the downcasting behaviour in the replace function
categories = [UserWarning, DeprecationWarning, FutureWarning]
for cat in categories:
    warnings.filterwarnings("ignore", category=cat)

# Import functions from the project directory
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()
json_files_dir = project_root / "json_files"

from code_files.automation.add_indications import update_indications
from code_files.automation.auto_base_uploads import start_auto_base_uploads
from code_files.automation.update_cohort_mapping import update_cohort_mapping
from code_files.automation.update_cohort_schema_mapping import (
    update_cohort_schema_mapping,
)
from code_files.automation.update_cohortname_cohortID_mapping import (
    update_cohortname_cohortID_mappings,
)
from code_files.automation.update_filepath_cohortname_mapping import (
    update_filepath_cohortname_mappings,
)
from code_files.automation.update_filepath_tablename_mapping import (
    update_filepath_tablename_mappings,
)
from code_files.automation.update_onedrive_cohort_schema_mapping import (
    update_onedrive_cohort_schema_mapping,
)
from code_files.base_scripts.creating_long_master_list import create_long_master_list
from code_files.base_scripts.creating_long_master_list_unanon import (
    create_long_master_list_unanon,
)
from code_files.etl_logic.etl_logic_followup_anon_unanon import (
    run_upload_process_followup_anon_unanon,
)
from code_files.etl_logic.etl_logic_patientList import start_batch_upload_patientList
from code_files.etl_logic.etl_logic_prelim_anon_unanon import run_upload_process_prelim
from code_files.etl_logic.etl_logic_scan_logs import run_upload_process_scan_logs
from code_files.etl_logic.etl_logic_supp_anon_unanon import (
    run_upload_process_supp_anon_unanon,
)
from code_files.utilities.io.database import (
    create_gcloud_engine,
    create_sql_engine,
    save_all_settings,
)
from code_files.utilities.io.json_manager import format_json, load_settings

# -----------------------------------------------------------------------------------------------------------------------------------------------------#


# Define the batch upload function
def start_batch_upload():
    # define the configuration settings
    # to run the ETL process for every uploaded table
    print("Loading Configuration...")
    upload_settings = json_files_dir / "file_upload_settings.json"
    config, settings_file_path = load_settings(upload_settings)

    # define a list for the failed table names
    failed_tables = []

    # Flag to decide whether there is a need to write to disk at the end
    should_save_to_disk = False

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

    # Establish connection to the database (GCP or MSSQL - depending on where the script is running)
    print("Establishing connection to the database...")
    # when the script is running from the NHS remote device, connect to SQL server
    if "AmmashA.IA" in str(current_dir):
        # create the SQL engine for the SQL database
        engine = create_sql_engine()

        # display a message if the engine failed to be created
        if engine is None:
            print("Cannot start upload without a valid database connection.")
            return

        # determine the number of datasets to process and upload to the database
        num_config_datasets = len(config)  # process all the datasets

    # when the script is running from the UOL PC, connect to GCP
    if "medaamm" in str(current_dir) or "ahmad" in str(current_dir):
        # create the connection engine to GCP
        env_path = project_root / "code_files" / "utilities" / "io" / ".env"
        engine = create_gcloud_engine(path=env_path)

        # display a message if the engine failed to be created
        if engine is None:
            print("Cannot start upload without a valid database connection.")
            return

        # determine the number of datasets to process and upload to the database
        num_config_datasets = 0
        # only count the datasets which are to be uploaded to GCP
        for info in config.values():
            if info["upload_destination"] == "all":
                num_config_datasets += 1

    # define the number of datasets already on the database and the number of those that have been successfully processed
    upload_count = 0
    success_count = 0

    print(
        f"Starting the batch re-upload process for {num_config_datasets} configured datasets..."
    )
    print("-" * 50)

    for table_name, table_config in config.items():
        # process the table if the configuration is valid
        # if the script is running on the UOL PC, skip the tables that should not be uploaded to it
        if (table_config["upload_destination"] == "mssql") and (
            "medaamm" in str(current_dir)
        ):
            continue

        # check if the configuration associated with that table is valid
        if not isinstance(table_config, dict):
            print(f"Skipping key '{table_name}': Invalid configuration format.")
            print("-" * 75)
            upload_count += 1
            continue
        upload_count += 1
        # Otherwise, process the table
        print(
            f"\n[{upload_count}/{num_config_datasets}] Processing table: **{table_name}**"
        )

        # run the appropriate ETL logic for the given dataset type
        # run the appropriate ETL logic for the given dataset type
        try:
            # if the upload process runs successfully, increment the success count variable
            if table_config["dataset_type"] in [
                "anonymised follow up",
                "unanonymised follow up",
            ]:
                if run_upload_process_followup_anon_unanon(table_config, engine):
                    success_count += 1
                    should_save_to_disk = True  # to mark that a change was made
                # otherwise, print a message saying the upload failed
                else:
                    print(
                        f"Upload failed for {table_name}. Continuing with the next table."
                    )
                    # add the table name to the failed uploads
                    failed_tables.append(table_name)
            elif table_config["dataset_type"] in [
                "anonymised prelim",
                "unanonymised prelim",
            ]:
                if run_upload_process_prelim(table_config, engine):
                    success_count += 1
                    should_save_to_disk = True  # to mark that a change was made
                # otherwise, print a message saying the upload failed
                else:
                    print(
                        f"Upload failed for {table_name}. Continuing with the next table."
                    )
                    # add the table name to the failed uploads
                    failed_tables.append(table_name)
            elif table_config["dataset_type"] in [
                "anonymised supplementary",
                "unanonymised supplementary",
            ]:
                if run_upload_process_supp_anon_unanon(table_config, engine):
                    success_count += 1
                    should_save_to_disk = True  # to mark that a change was made
                # otherwise, print a message saying the upload failed
                else:
                    print(
                        f"Upload failed for {table_name}. Continuing with the next table."
                    )
                    # add the table name to the failed uploads
                    failed_tables.append(table_name)
            elif table_config["dataset_type"] == "scan logs":
                if run_upload_process_scan_logs(table_config, engine):
                    success_count += 1
                    should_save_to_disk = True  # to mark that a change was made
                # otherwise, print a message saying the upload failed
                else:
                    print(
                        f"Upload failed for {table_name}. Continuing with the next table."
                    )
                    # add the table name to the failed uploads
                    failed_tables.append(table_name)
            # update the timestamp in memory
            table_config["last_upload_timestamp"] = datetime.now().strftime(
                "%d-%m-%Y %H:%M:%S"
            )

        except (AttributeError, KeyError, TypeError, ValueError, OSError) as e:
            # write the error to the log file
            last_exception = e
            sys.stderr.write(
                "\n========================================================================================================================================================\n"
            )
            sys.stderr.write(f"\n❌ ERROR PROCESSING TABLE: {table_name} ---\n")
            sys.stderr.write(
                f"🔎 REASON: {type(last_exception).__name__}: {last_exception!s}\n"
            )
            sys.stderr.write("📝 Error in processing dataset.")
            sys.stderr.write(
                "\n========================================================================================================================================================\n"
            )
    # Print a message to say that the indications fields are being updated
    print("Adding the Indications to the Supplementary datasets...")

    # Update the indications fields in the anonymised supplementary datasets
    update_indications()

    # Print a message to say that the JSON file is being updated
    print("Updating the JSON configuration file...")

    # Update the json configuration file
    format_json()

    # Print a message on the log file showing the number of successful uploads
    sys.stderr.write(
        "\n\n========================================================================================================================================================\n\n"
    )
    sys.stderr.write(
        f"Number of Successful Uploads: {success_count}/{upload_count} tables."
    )
    sys.stderr.write(
        "\n\n========================================================================================================================================================\n\n"
    )
    sys.stderr.flush()

    # Print a message to the log file showing the tables that failed to upload (when there are any)
    if success_count < upload_count:
        sys.stderr.write(
            "\n\n========================================================================================================================================================\n\n"
        )
        sys.stderr.write(f"Upload failed for the tables: {failed_tables}")
        sys.stderr.write(
            "\n\n========================================================================================================================================================\n\n"
        )
        sys.stderr.flush()

    if should_save_to_disk:
        save_all_settings(config, settings_file_path)

    if success_count < upload_count:
        print(f"\n⚠️ Completed with {success_count}/{upload_count} successful uploads.")
    else:
        print(f"\n✅ All {upload_count} uploads completed successfully.")

    return success_count == upload_count


if __name__ == "__main__":
    print(
        "\n\n========================================================================================================================================================\n\n"
    )
    print("✅ Updating the Patient List and TriStar Pathway dataset...")
    start_batch_upload_patientList()
    print("✅ Patient List and TriStar Pathway updated.")
    print(
        "\n\n========================================================================================================================================================\n\n"
    )

    print(
        "\n\n========================================================================================================================================================\n\n"
    )
    print("✅ Starting batch upload for datasets in the JSON configuration file...")
    start_batch_upload()
    print("✅ Completed the batch upload for datasets in the JSON configuration file.")
    print(
        "\n\n========================================================================================================================================================\n\n"
    )

    print(
        "\n\n========================================================================================================================================================\n\n"
    )
    print("✅ Updating cohort mapping JSON file...")
    update_cohort_mapping()
    print("✅ Updated cohort mapping JSON file.")
    print(
        "\n\n========================================================================================================================================================\n\n"
    )

    print(
        "\n\n========================================================================================================================================================\n\n"
    )
    print("✅ Updating cohort schema mapping JSON file...")
    update_cohort_schema_mapping()
    print("✅ Updated cohort schema mapping JSON file.")
    print(
        "\n\n========================================================================================================================================================\n\n"
    )

    print(
        "\n\n========================================================================================================================================================\n\n"
    )
    print("✅ Updating cohort name to cohort ID mapping JSON file...")
    update_cohortname_cohortID_mappings()
    print("✅ Updated cohort name to cohort ID mapping JSON file.")
    print(
        "\n\n========================================================================================================================================================\n\n"
    )

    print(
        "\n\n========================================================================================================================================================\n\n"
    )
    print("✅ Updating file path to cohort name mapping JSON file...")
    update_filepath_cohortname_mappings()
    print("✅ Updated file path to cohort name mapping JSON file.")
    print(
        "\n\n========================================================================================================================================================\n\n"
    )

    print(
        "\n\n========================================================================================================================================================\n\n"
    )
    print("✅ Updating file path to table name mapping JSON file...")
    update_filepath_tablename_mappings()
    print("✅ Updated file path to table name mapping JSON file.")
    print(
        "\n\n========================================================================================================================================================\n\n"
    )

    print(
        "\n\n========================================================================================================================================================\n\n"
    )
    print("✅ Updating onedrive to cohort schema mapping JSON file...")
    update_onedrive_cohort_schema_mapping()
    print("✅ Updated onedrive to cohort schema mapping JSON file.")
    print(
        "\n\n========================================================================================================================================================\n\n"
    )

    print(
        "\n\n========================================================================================================================================================\n\n"
    )
    print("✅ Running auto base uploads from the Box folder...")
    start_auto_base_uploads()
    print("✅ Completed auto base uploads from the Box folder.")
    print(
        "\n\n========================================================================================================================================================\n\n"
    )

    print(
        "\n\n========================================================================================================================================================\n\n"
    )
    print("✅ Creating the anonymised master list...")
    create_long_master_list()
    print("✅ Created and uploaded the anonymised master list.")
    print(
        "\n\n========================================================================================================================================================\n\n"
    )

    print(
        "\n\n========================================================================================================================================================\n\n"
    )
    print("✅ Creating the de-anonymised master list...")
    create_long_master_list_unanon()
    print("✅ Created the de-anonymised master list")
    print(
        "\n\n========================================================================================================================================================\n\n"
    )


# -----------------------------------------------------------------------------------------------------------------------------------------------------#
