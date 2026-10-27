# Import the required libraries
import pandas as pd
from IPython.display import clear_output, display
from ipywidgets import widgets
from sqlalchemy.exc import SQLAlchemyError

from code_files.utilities.config.mappings import set_data_types_scan_logs
from code_files.utilities.io.json_manager import save_settings_generic
from code_files.utilities.io.readers import find_headerNumber, show_worksheet_names
from code_files.utilities.io.setup_db_cnxn import setup_cnxn
from code_files.utilities.processing.cleaners import (
    analyse_cell_colors,
    remove_empty_and_complete,
    title_case,
    trim_column_names,
    trim_spaces,
)

# Import the required functions
from code_files.utilities.ui.widgets import select_workbook_main, select_worksheet_main

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

# Define the json files directory
json_files_dir = project_root / "json_files"

# ------------------------------------------------------------------------------------------------------------------

# Define the file path to the JSON file
settings_file_path = project_root / "json_files" / "file_upload_settings.json"

# ------------------------------------------------------------------------------------------------------------------


# Define a function to ask the user for a dataframe name
def get_dataframe_name_main():
    """This function asks the user to enter a name for a dataframe and adds the Scan Log suffix to it."""

    # Ask the user to enter the dataframe name
    df_name = input("Enter a name for the selected dataframe (without the suffix)")

    # Add the suffix to the input dataframe name
    df_name = df_name + "_Scan_Logs"

    # Define the global dataset_type variable and assign based on the name provided
    # Define the global id_column variable and assign based on the name provided
    global dataset_type

    # Set the dataset type variable
    dataset_type = "scan logs"

    return df_name


# ------------------------------------------------------------------------------------------------------------------

# Define a Function to Standardise the Column Names
column_name_map = {
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


def rename_columns(main_workbook):
    """
    This function renames the columns in a scan log dataset using the pre-defined column name mapping dictionnary.
    """
    # rename the columns using the dictionary defined above
    main_workbook = main_workbook.rename(columns=column_name_map)
    # convert the column names to a list
    cols = list(main_workbook.columns)
    # define the occurence variable - to count the occurences of specific olumns
    occurence = 0
    # rename the columns which have the same name in the original worksheet
    # loop through the column names
    for i, name in enumerate(cols):
        # Rename the column according to the number of occurences of "Error Code"
        if name == "Error Code":
            occurence += 1
            if occurence == 1:
                cols[i] = "ErrorCode_Scan1"
            elif occurence == 2:
                cols[i] = "ErrorCode_Scan2"
        # Similarly, rename the columns depending on the number of occurences of "Notes"
        if name == "Notes":
            occurence += 1
            if occurence == 1:
                cols[i] = "Notes_Scan1"
            elif occurence == 2:
                cols[i] = "Notes_Scan2"

    main_workbook.columns = cols

    return main_workbook


# ------------------------------------------------------------------------------------------------------------------


# ------------------------------------------------------------------------------------------------------------------


# Define a function to import the worksheet
def import_worksheet(
    filename_main, selected_sheet, header_row_main, df_name_main, engine
):
    """
    This function enables the processing and manual transfer of scan logs sheets to the database.
    """
    # Load and process the main workbook
    main_workbook = pd.read_excel(
        filename_main, sheet_name=selected_sheet, header=header_row_main
    )
    main_workbook = trim_column_names(main_workbook)
    main_workbook = trim_spaces(main_workbook)
    main_workbook = title_case(main_workbook)
    main_workbook = rename_columns(main_workbook)
    main_workbook = remove_empty_and_complete(main_workbook)
    main_workbook = analyse_cell_colors(
        filename_main, selected_sheet, header_row_main, main_workbook
    )
    data_types = set_data_types_scan_logs(main_workbook)

    # Read the schema list
    schemas_list_path = project_root / "text_files" / "schemas_list.txt"
    with open(schemas_list_path, "r") as file:
        schemas_list = [line.strip() for line in file]

    # Create widgets
    dropdown = widgets.Dropdown(
        options=schemas_list, description="Schema:", value=schemas_list[0]
    )

    # create the "Submit" button
    submit_button = widgets.Button(description="Upload to SQL", button_style="success")
    # create the container for any messages the function prints
    output = widgets.Output()

    # Define upload logic for when the submit button is clicked
    def on_submit_clicked(b):
        """
        This function runs when the "Submit" button is clicked in order to upload a scan log dataset to the database.
        """
        with output:
            # define the main workbook variable used from the import_worksheet function (this function is nested inside it)
            nonlocal main_workbook
            # clear the output of the widget so that previous messages are erased
            clear_output()
            # set the selected schema equal to the value selected from the dropdown (after the submit button is clicked)
            selected_schema = dropdown.value

            try:
                print(
                    f"✅ Uploading dataframe '{df_name_main}' to schema '{selected_schema}'"
                )
                # upload the main workbook to the database
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

                # Save settings
                save_settings_generic(
                    settings_file_path,
                    filename_main,
                    selected_sheet,
                    df_name_main,
                    header_row_main,
                    selected_schema,
                    data_types,
                    column_name_map,
                    "scan logs",
                    "mssql",
                )

            except (ValueError, TypeError, OSError, SQLAlchemyError) as e:
                print(f"❌ Upload failed: {e}")

    submit_button.on_click(on_submit_clicked)

    # Display UI
    display(widgets.VBox([dropdown, submit_button, output]))

    # Return the main processed dataframe for further use
    return main_workbook


# ------------------------------------------------------------------------------------------------------------------


# Define a main function combining the functions above
def main():
    # Select workbook and get the sheet names
    filename_main = select_workbook_main(
        "Please select a scan log", "D:/AmmashA.IA/OneDrive - NHS"
    )
    sheet_names_main = show_worksheet_names(filename_main=filename_main)

    # Call the function to select a worksheet
    selected_sheet = select_worksheet_main(sheet_names_main=sheet_names_main)

    # Get the header row of the selected worksheet
    header_row = find_headerNumber(filename_main, selected_sheet, "scan logs")

    # Get the name of the dataframe from the user
    df_name = get_dataframe_name_main()

    # Import and upload the selected dataframe after processing
    import_worksheet(filename_main, selected_sheet, header_row, df_name, engine)


if __name__ == "__main__":
    main()
