# Import required libraries
import json
import re

import ipywidgets as widgets
import pandas as pd
from IPython.display import clear_output, display
from sqlalchemy.exc import SQLAlchemyError

# Import the required functions
from code_files.automation.anonymiser import anonymiser
from code_files.utilities.config.mappings import (
    add_cohortName_column,
    create_column_map,
    map_cohort_names_patientList,
    set_data_types,
    set_data_types_patientList,
)
from code_files.utilities.io.json_manager import save_settings_generic
from code_files.utilities.io.setup_db_cnxn import setup_cnxn
from code_files.utilities.processing.cleaners import (
    rename_columns,
    rename_PatientInventory_Columns,
    replace_z_s,
    standardise_cohortName_patientList,
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

# JSON file for the data uploaded to the database (anonymised and unanonymised prelim, follow up, and supplementary data) + scan logs
settings_file_path = json_files_dir / "file_upload_settings.json"

# JSON file for the cohort_mapping
json_path = project_root / "json_files" / "cohort_names_mapping.json"
with open(json_path, "r") as f:
    cohort_mapping = json.load(f)

# define the username and server variables
if "D:" in str(current_dir) or "AmmashA.IA" in str(current_dir):
    server = "SQL"
elif "medaamm" in str(current_dir):
    server = "GCP"


def import_worksheet(
    filename_main,
    selected_sheet,
    header_row_main,
    df_name_main,
    engine,
    column_name_map,
    dataset_type,
    id_column,
    auto_mode=False,
    cohort_name=None,
    schema_name=None,
    df_in=None,
):
    """
    This function handles the processing and upload of all datasets to the database - both in the automated and manual pipelines
    """
    # Load the main workbook
    main_workbook = pd.read_excel(
        filename_main, sheet_name=selected_sheet, header=header_row_main
    )

    # When the function is not running automatically, check if the user wants to anonymise the dataset before proceeding
    if auto_mode == False:
        to_anonymise_answer = input(
            "Would you like to anonymise this unanonymised dataset before proceeding with the database upload? (Y/N):"
        )
        if to_anonymise_answer == "Y":
            main_workbook = anonymiser(
                source_sheet_name=selected_sheet,
                df_onedrive=main_workbook,
                cohort_name=cohort_name,
                source_file_name=df_name_main,
                testing_mode=False,
                skip_box_upload=True,
            )
            # in case the dataframe is anonymised, reset the dataset type to anonymised, the id_column to DonorID, and the df_name_main to Anon
            dataset_type = dataset_type.replace("unanonymised", "anonymised")
            df_name_main = df_name_main.replace("Unanon", "Anon")

    # Process the main workbook
    main_workbook = trim_column_names(main_workbook)
    main_workbook = trim_spaces(main_workbook)
    main_workbook = replace_z_s(main_workbook)
    main_workbook = title_case(main_workbook)
    main_workbook = rename_columns(main_workbook, create_column_map())
    data_types = set_data_types(main_workbook, dataset_type)
    main_workbook = process_date_columns(main_workbook, dataset_type)
    main_workbook = process_bit_columns(main_workbook)
    main_workbook = process_dec_columns(main_workbook)

    # In the case of preliminary data, make sure not all the columns are selected in the case where the full follow up data columns are present
    # but are blank - which would throw an error
    if dataset_type in ["anonymised prelim", "unanonymised prelim"] and "StageAtDX" in main_workbook.columns:
            main_workbook = main_workbook.loc[:, :"StageAtDX"]

    # Replace N/A values with actual nulls in the columns in the ID columns
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

    # Read schema list
    schemas_list_path = project_root / "text_files" / "schemas_list.txt"
    with open(schemas_list_path, "r") as file:
        schemas_list = [line.strip() for line in file]

    for col in main_workbook.columns:
        col_name = str(col).strip()

        # Match columns containing 'date' or 'dob', but ignore Time-To-Event ('tte') metrics
        if re.search(r"Date|DOB", col_name, re.IGNORECASE) and not re.search(
            r"^TTE", col_name, re.IGNORECASE
        ):
            # check if the column contains non-standard text patterns like quarter
            contains_quarter = (
                main_workbook[col]
                .astype(str)
                .str.contains(r"Q[1-4]\s*[-/]", regex=True, na=False)
                .any()
            )
            if not contains_quarter:
                # errors='coerce' turns empty spaces, strings, or bad floats into NaT (Not a Time)
                main_workbook[col] = pd.to_datetime(main_workbook[col], errors="coerce")
            else:
                main_workbook[col] = (
                    main_workbook[col]
                    .astype(str)
                    .replace(["nan", "NaN", "None", "<NA>"], None)
                )

    # when the script is running automatically, don't display the widgets and provide the cohort and schema names as
    # arguments to the function
    if auto_mode:
        selected_schema = schema_name
        if isinstance(cohort_name, dict):
            current_cohort_string = cohort_name.get("value")
        else:
            current_cohort_string = str(cohort_name)
        main_workbook["CohortName"] = current_cohort_string

        if dataset_type in ["anonymised supplementary", "unanonymised supplementary"]:
            try:
                main_workbook.to_sql(
                    df_name_main,
                    engine,
                    if_exists="replace",
                    schema=selected_schema,
                    index=False,
                    dtype=data_types,
                )
                print(f"Uploading the supplementary dataset {df_name_main}")

                # Set the destination variable based on the dataset type
                destination = "mssql" if "unaonymised" in dataset_type else "all"

                # Set the template version variable based on the column count and dataset type

                # Save the settings logic while factoring in the dataset type
                save_settings_generic(
                    settings_file_path=settings_file_path,
                    file_path=filename_main,
                    sheet_name=selected_sheet,
                    sql_table_name=df_name_main,
                    header_index=header_row_main,
                    schema_name=selected_schema,
                    data_types=data_types,
                    column_names_mapping=column_name_map,
                    dataset_type=dataset_type,
                    upload_destination=destination,
                    cohort_name=current_cohort_string,
                )
            except (SQLAlchemyError, ValueError, OSError) as e:
                print(f"Dataset {df_name_main} failed to upload: {e}")

        elif dataset_type in [
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

            # re-order the dataframe columns to place the CohortName at the start
            ["CohortName"] + [
                col for col in main_workbook.columns if col not in ["CohortName"]
            ]

            # convert the date columns to the correct format before uploading to the database - when the dataset is of the unanonymised type
            if "unanonymised" in dataset_type:
                for col in main_workbook.columns:
                    if ("date" in col.lower() or "dob" in col.lower()) and (
                        "tte" not in col.lower()
                    ):
                        # Clear any hidden float garbage or string artifacts by parsing to datetime first
                        parsed_dates = pd.to_datetime(
                            main_workbook[col], errors="coerce"
                        )

                        # Convert valid datetimes directly to clean text string formats ('YYYY-MM-DD')
                        formatted_strings = parsed_dates.dt.strftime("%Y-%m-%d").astype(
                            object
                        )

                        # Swap out NaT null values with clean Python Nones for database NULL compatibility
                        main_workbook[col] = formatted_strings.where(
                            parsed_dates.notna(), None
                        )
            try:
                main_workbook.to_sql(
                    df_name_main,
                    engine,
                    if_exists="replace",
                    schema=selected_schema,
                    index=False,
                    dtype=data_types,
                )
                print(f"Uploading dataset {df_name_main}")

                # Set the destination variable based on the dataset type
                destination = "mssql" if "unanonymised" in dataset_type else "all"

                # Set the template version variable based on the column count and dataset type

                # Save the settings logic while factoring in the dataset type
                save_settings_generic(
                    settings_file_path=settings_file_path,
                    file_path=filename_main,
                    sheet_name=selected_sheet,
                    sql_table_name=df_name_main,
                    header_index=header_row_main,
                    schema_name=selected_schema,
                    data_types=data_types,
                    column_names_mapping=column_name_map,
                    dataset_type=dataset_type,
                    upload_destination=destination,
                    cohort_name=current_cohort_string,
                )
            except (SQLAlchemyError, ValueError, OSError) as e:
                print(f"Dataset {df_name_main} failed to upload: {e}")

    else:  # auto_mode = off
        # Add the cohort name column
        selected_cohort = add_cohortName_column()

        # Create widgets
        dropdown = widgets.Dropdown(
            options=schemas_list, description="Schema:", value=schemas_list[0]
        )

        submit_button = widgets.Button(
            description=f"Upload to {server}", button_style="success"
        )
        output = widgets.Output()

    # Define upload logic
    def on_submit_clicked(b):
        with output:
            nonlocal main_workbook

            clear_output()
            if not auto_mode:
                selected_schema = dropdown.value

            # Set the cohort name immediately before the upload
            if selected_cohort["value"] is None:
                print("❌ Upload failed: Please select a valid cohort name.")
                return

            # print a message saying the cohort name column was successfully added
            print("✅ Cohort Name column added.")

            try:
                # map the selected cohort to the name from the cohort names mapping dictionary
                main_workbook["CohortName"] = cohort_mapping[selected_cohort["value"]]
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
                        blocks = (
                            main_workbook[existing_block_cols].bfill(axis=1).iloc[:, 0]
                        )
                    else:
                        blocks = pd.Series("UnknownBlock", index=main_workbook.index)

                    # Clean up the ids and block numbers
                    ids = ids.astype(str).replace(["<NA>", "nan"], "UnknownID")
                    blocks = blocks.astype(str).replace(["<NA>", "nan"], "UnknownBlock")

            except KeyError as e:
                print(f"❌ An error occurred during processing: {e}")

            # re-order the dataframe columns to place the CohortName at the start
            if dataset_type in [
                "anonymised follow up",
                "anonymised prelim",
                "unanonymised follow up",
                "unanonymised prelim",
            ]:
                new_columns = ["CohortName"] + [
                    col for col in main_workbook.columns if col != "CohortName"
                ]
            main_workbook = main_workbook[new_columns]

            # convert the date columns to the correct format before uploading to the database - when the dataset is of the unanonymised type
            if "unanonymised" in dataset_type:
                for col in main_workbook.columns:
                    # Use a case-insensitive check matching your type-setting logic
                    if ("date" in col.lower() or "dob" in col.lower()) and (
                        "tte" not in col.lower()
                    ):
                        # Force parse to datetime to eliminate epoch integer conversions
                        parsed_dates = pd.to_datetime(
                            main_workbook[col], errors="coerce"
                        )

                        # Keep as an object layout and insert explicit clean Python Nones for SQL NULL compatibility
                        main_workbook[col] = parsed_dates.dt.strftime(
                            "%Y-%m-%d"
                        ).astype(object)
                        main_workbook[col] = main_workbook[col].where(
                            main_workbook[col].notna(), None
                        )

            # try uploading the dataframe to the database
            try:
                print(
                    f"✅ Uploading dataframe '{df_name_main}' to schema '{selected_schema}"
                )
                main_workbook.to_clipboard(index=False)
                main_workbook.to_sql(
                    df_name_main,
                    engine,
                    if_exists="replace",
                    schema=selected_schema,
                    index=False,
                    dtype=data_types,
                )
                print(
                    f"✅ Dataframe '{df_name_main}' uploaded to schema '{selected_schema}' successfully."
                )

                # Set the destination variable based on the dataset type
                destination = "mssql" if "unanonymised" in dataset_type else "all"

                # Set the template version variable based on the column count and dataset type

                # Save the settings logic while factoring in the dataset type
                save_settings_generic(
                    settings_file_path=settings_file_path,
                    file_path=filename_main,
                    sheet_name=selected_sheet,
                    sql_table_name=df_name_main,
                    header_index=header_row_main,
                    schema_name=selected_schema,
                    data_types=data_types,
                    column_names_mapping=column_name_map,
                    dataset_type=dataset_type,
                    upload_destination=destination,
                    cohort_name=selected_cohort.get("value"),
                )
            except (SQLAlchemyError, ValueError, OSError) as e:
                print(f"❌ Upload failed: {e}")

    if not auto_mode:
        submit_button.on_click(on_submit_clicked)

        # Display UI
        display(widgets.VBox([dropdown, submit_button, output]))

    # Return the processed DataFrame for further use
    return main_workbook


def run_processing_flow(selected_cohort_value, box_data_package, dataset_type):
    (
        _onedrivebox_filepath,
        selected_sheetName_onedrivebox,
        _header_row_onedrivebox,
        onedrivebox_filename,
        onedrive_box_filepath,
        selected_sheetName_onedrive,
        header_row_onedrive,
        _onedrive_filename_unused,
    ) = box_data_package
    """
    This function runs the anonymiser function which anonymises prelim, follow up, and supplementary datasets
    """

    print(f"Processing the data for cohort: {selected_cohort_value}...")

    # read in the OneDrive file
    onedrive_data = pd.read_excel(
        onedrive_box_filepath,
        sheet_name=selected_sheetName_onedrive,
        header=header_row_onedrive,
    )

    # apply the anonymisation based on the dataset_type global variable value
    if dataset_type in ["anonymised prelim", "anonymised follow up"] or dataset_type in ["anonymised supplementary"]:
        anonymiser(
            selected_sheetName_onedrivebox,
            onedrive_data,
            selected_cohort_value,
            onedrivebox_filename,
        )

    print(f"✅ Successfully anonymised data for the cohort {selected_cohort_value}.")


def import_worksheet_patientList(
    filename_main, selected_sheet, header_row_main, df_name_main, engine
):
    """
    This function processes and uploads the patient list to the database
    """
    main_workbook = pd.read_excel(
        filename_main, sheet_name=selected_sheet, header=header_row_main
    )

    # Process the dataframe
    main_workbook = trim_column_names(main_workbook)
    main_workbook = trim_spaces(main_workbook)
    main_workbook = title_case(main_workbook)
    main_workbook = standardise_cohortName_patientList(main_workbook)
    main_workbook = rename_PatientInventory_Columns(main_workbook)
    data_types_PatientList = set_data_types_patientList(main_workbook)
    main_workbook = process_bit_columns(main_workbook)
    main_workbook = map_cohort_names_patientList(main_workbook)
    main_workbook.to_sql(
        df_name_main,
        engine,
        if_exists="replace",
        schema="imp_documents",
        index=False,
        dtype=data_types_PatientList,
    )
    print(f"✅ Dataframe '{df_name_main}' uploaded successfully.")

    return main_workbook, data_types_PatientList
