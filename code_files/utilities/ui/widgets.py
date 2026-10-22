# Import the required libraries
import json
import tkinter as tk
from tkinter.filedialog import askopenfilename

import ipywidgets as widgets
import pandas as pd
from IPython.display import display

from code_files.utilities.io.json_manager import (
    add_to_json_box_cnxn,
    add_to_json_onedrive_cnxn,
)
from code_files.utilities.io.readers import find_headerNumber
from code_files.utilities.io.setup_db_cnxn import setup_cnxn
from code_files.utilities.pipelines.upload_pipeline import run_processing_flow

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

# define the username and server variables
if "D:" in str(current_dir):
    username = "AmmashA.IA"
    server = "SQL"
elif "medaamm" in str(current_dir):
    username = "medaamm"
    server = "GCP"

# load the json file for the cohort mapping
json_path = project_root / "json_files" / "cohort_names_mapping.json"
with open(json_path, "r") as f:
    cohort_mapping = json.load(f)


def select_workbook_main(title, initial_dir):
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    filename_main = askopenfilename(
        title=title,
        filetypes=[("Excel files", "*.xlsx;*.xls;*.xlsm")],
        initialdir=initial_dir,
    )  # define the accepted file types
    return filename_main


def select_worksheet_main(sheet_names_main):
    selected_sheetNumber = input(
        "Using the numbers of the sheets listed above, enter the number of the worksheet to import."
    )
    selected_sheetName = sheet_names_main[int(selected_sheetNumber) - 1]
    return selected_sheetName


def get_dataframe_name_main():

    # Define a list for the dataset type options
    data_types = [
        "Anonymised Follow Up",
        "Unanonymised Follow Up",
        "Anonymised Prelim",
        "Unanonymised Prelim",
        "Anonymised Supplementary",
        "Unanonymised Supplementary",
        "Scan Logs",
    ]
    # Show the options for the dataset type
    print("Below are the possible data types:")
    for i, data_type in zip(range(1, 8), data_types):
        print(f"{i}: {data_type}")

    # Ask the user to enter the dataframe name
    df_name = input(
        "Enter a name for the selected dataframe (without the dataset type suffix)."
    )

    # Ask the user to enter a number for the dataset type
    selected_dataset_type = input(
        "Using the dataset types and numbers shown, enter the relevant number matching the selected dataset's type."
    )

    # Define the global dataset_type variable and assign based on the name provided
    # Define the global id_column variable and assign based on the name provided
    global dataset_type
    global id_column

    # Initialise the variables so that we don't get an error if their value is None at the end of the function
    id_column = None
    dataset_type = None

    # Define the suffix variable based on the number of the dataset type input by the user
    match selected_dataset_type:
        case "1":
            suffix = "_FollowUp_Anon"
            dataset_type = "anonymised follow up"
            id_column = "DonorID"

        case "2":
            suffix = "_FollowUp_Unanon"
            dataset_type = "unanonymised follow up"
            id_column = "NHS_Number"

        case "3":
            suffix = "_Prelim_Anon"
            dataset_type = "anonymised prelim"
            id_column = "DonorID"

        case "4":
            suffix = "_Prelim_Unanon"
            dataset_type = "unanonymised prelim"
            id_column = "NHS_Number"

        case "5":
            suffix = "_Supplementary_Anon"
            dataset_type = "anonymised supplementary"

        case "6":
            suffix = "_Supplementary_Unanon"
            dataset_type = "unanonymised supplementary"

        case "7":
            suffix = "_Scan_Logs"

    # Append the suffix to the df_name
    df_name = df_name + suffix

    return df_name, dataset_type, id_column


def get_worksheets_onedrive_cnxn():
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    # Get the information on the OneDrive Box uploads workbook
    onedrivebox_filepath = select_workbook_main(
        "Please select the anonymised workbook from the Box uploads folder on OneDrive",
        f"C:/Users/{username}/OneDrive - NHS/TrisStar metadata/Box uploads/Updated template",
    )

    # Show the worksheet names
    xl = pd.ExcelFile(onedrivebox_filepath)
    sheetnames_onedrivebox = xl.sheet_names
    print(
        "Below are the worksheets in the imported workbook from OneDrive Box Uploads:"
    )
    for i, sheet in enumerate(sheetnames_onedrivebox):
        print(f"{i + 1}: {sheet}")

    # Ask the user to enter the worksheet number for the Box data
    selected_sheetNumber_onedrivebox = input(
        "Using the numbers of the sheets listed above, please enter the number of the worksheet to import from OneDrive Box Uploads."
    )
    selected_sheetName_onedrivebox = sheetnames_onedrivebox[
        int(selected_sheetNumber_onedrivebox) - 1
    ]

    # Define a list for the dataset type options and ask the user for the dataset type
    print(
        "Using the numbers and dataset types shown below, please enter the number ressembling the selected dataset type:"
    )
    data_types = [
        "Anonymised Follow Up",
        "Anonymised Prelim",
        "Anonymised Supplementary",
    ]

    for i, data_type in zip(range(1, 7), data_types):
        print(f"{i}: {data_type}")

    selected_dataset_type = input(
        "Using the dataset types and numbers shown, enter the relevant number matching the selected dataset's type."
    )

    # assign the global dataset_type variable using the selected dataset type number
    global dataset_type
    match selected_dataset_type:
        case "1":
            dataset_type = "anonymised follow up"
        case "2":
            dataset_type = "anonymised prelim"
        case "3":
            dataset_type = "anonymised supplementary"

    # Get the header number of the Box worksheet using the defined function
    header_row_onedrivebox = find_headerNumber(
        onedrivebox_filepath, selected_sheetName_onedrivebox, dataset_type
    )

    # -------------------------------------------------------------------------------------------------------------

    # Get the information on the OneDrive workbook
    onedrive_box_filepath = select_workbook_main(
        "Please select the unanonymised data workbook from the OneDrive folder.",
        f"C:/Users/{username}/OneDrive - NHS/TrisStar metadata",
    )

    # Show the worksheet names
    xl = pd.ExcelFile(onedrive_box_filepath)
    sheet_names_onedrive = xl.sheet_names
    print(
        "Below are the worksheets in the imported unanonymised data workbook from the OneDrive folder:"
    )
    for i, sheet in enumerate(sheet_names_onedrive):
        print(f"{i + 1}: {sheet}")

    # Ask the user to enter the worksheet number for the OneDrive data
    selected_sheetNumber_onedrive = input(
        "Using the numbers of the sheets listed above, please enter the number of the unanonymised data worksheet to import from the OneDrive folder."
    )
    selected_sheetName_onedrive = sheet_names_onedrive[
        int(selected_sheetNumber_onedrive) - 1
    ]

    # assign a temporary dataset_type_unanon variable based on the global dataset_type variable
    match dataset_type:
        case "anonymised follow up":
            dataset_type_unanon = "unanonymised follow up"
        case "anonymised prelim":
            dataset_type_unanon = "unanonymised prelim"
        case "anonymised supplementary":
            dataset_type_unanon = "unanonymised supplementary"

    # Get the header row number for the unanonymised worksheet
    header_row_onedrive = find_headerNumber(
        onedrive_box_filepath, selected_sheetName_onedrive, dataset_type_unanon
    )

    return (
        onedrivebox_filepath,
        selected_sheetName_onedrivebox,
        header_row_onedrivebox,
        onedrive_box_filepath,
        selected_sheetName_onedrive,
        header_row_onedrive,
    )


# -------------------------------------------------------------------------------------------------------------

# In[30]: Function to get information on the Box and OneDrive worksheets (box connection)


def get_worksheets_box_cnxn():
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    # Get the information on the Box workbook
    box_filepath = select_workbook_main(
        "Please select the anonymised workbook from Box.", f"C:/Users/{username}/Box"
    )

    # Show the worksheet names
    xl = pd.ExcelFile(box_filepath)
    sheetnames_box = xl.sheet_names
    print("Below are the worksheets in the imported workbook from Box:")
    for i, sheet in enumerate(sheetnames_box):
        print(f"{i + 1}: {sheet}")

    # Ask the user to enter the worksheet number for the Box data
    selected_sheetNumber_box = input(
        "Using the numbers of the sheets listed above, please enter the number of the worksheet to import from Box."
    )
    selected_sheetName_box = sheetnames_box[int(selected_sheetNumber_box) - 1]

    # Define a list for the dataset type options and ask the user for the dataset type
    print(
        "Using the numbers and dataset types shown below, please enter the number ressembling the selected dataset type:"
    )
    data_types = [
        "Anonymised Follow Up",
        "Anonymised Prelim",
        "Anonymised Supplementary",
    ]

    for i, data_type in zip(range(1, 7), data_types):
        print(f"{i}: {data_type}")

    selected_dataset_type = input(
        "Using the dataset types and numbers shown, enter the relevant number matching the selected dataset's type."
    )

    # assign the global dataset_type variable using the selected dataset type number
    global dataset_type
    match selected_dataset_type:
        case "1":
            dataset_type = "anonymised follow up"
        case "2":
            dataset_type = "anonymised prelim"
        case "3":
            dataset_type = "anonymised supplementary"

    # Get the header number of the Box worksheet using the defined function
    header_row_box = find_headerNumber(
        box_filepath, selected_sheetName_box, dataset_type
    )

    # -------------------------------------------------------------------------------------------------------------

    # Get the information on the OneDrive Box workbook
    # define the folder within the Box uploads folder depending on the dataset type
    if dataset_type in ["anonymised follow up", "anonymised supplementary"]:
        sub_folder = "Follow up"
    else:
        sub_folder = "Prelim"
    # pick out the file name from the box file path
    parts = box_filepath.split("/")
    file_name = "/".join(parts[6:])
    onedrive_box_filepath = f"C:/Users/{username}/OneDrive - NHS/TrisStar metadata/Box uploads/Updated template/{sub_folder}/{file_name}"

    # Show the worksheet names
    xl = pd.ExcelFile(onedrive_box_filepath)
    sheet_names_onedrive_box = xl.sheet_names
    print(
        "Below are the worksheets in the imported unanonymised data workbook from the OneDrive Box folder:"
    )
    for i, sheet in enumerate(sheet_names_onedrive_box):
        print(f"{i + 1}: {sheet}")

    # Ask the user to enter the worksheet number for the OneDrive Box data
    selected_sheetNumber_onedrive_box = input(
        "Using the numbers of the sheets listed above, please enter the number of the anonymised data worksheet from the OneDrive Box folder."
    )
    selected_sheetName_onedrive_box = sheet_names_onedrive_box[
        int(selected_sheetNumber_onedrive_box) - 1
    ]

    # Get the header row number for the unanonymised worksheet
    header_row_onedrive_box = find_headerNumber(
        onedrive_box_filepath, selected_sheetName_onedrive_box, dataset_type
    )

    return (
        box_filepath,
        selected_sheetName_box,
        header_row_box,
        onedrive_box_filepath,
        selected_sheetName_onedrive_box,
        header_row_onedrive_box,
    )


def add_cohortName_column_live_cnxn(cohort_names, box_data_package, connection_type):
    # create an output widget to capture prints
    output_box = widgets.Output()

    sorted(cohort_names)

    # Add the cohort names to the dropdown menu
    dropdown = widgets.Dropdown(
        options=cohort_names,
        value=None,
        description="Please select the cohort: ",
        disabled=False,
        layout=widgets.Layout(width="450px"),
        style={"description_width": "initial"},
    )

    # Define a variable to store the selected cohort
    selected_cohort = {"value": dropdown.value, "confirmed": False}

    # observe the widget for changes in the selection and update accordingly
    def on_change(change):
        if change["type"] == "change" and change["name"] == "value":
            selected_cohort["value"] = change["new"]
            print(f"You selected the cohort named {selected_cohort['value']}.")

    # Create a confirm button
    confirm_button = widgets.Button(
        description="Confirm",
        button_style="success",
        tooltip="Click to confirm the selection",
        disabled=False,
        layout=widgets.Layout(width="100px"),
    )

    # Handle button click
    def on_button_click(button):
        with output_box:
            # Directly get the current value from the dropdown widget
            current_selection = dropdown.value

            if current_selection is not None:
                # Call the JSON function
                if connection_type == "onedrive_cnxn":
                    success = add_to_json_onedrive_cnxn(
                        *box_data_package[:8],
                        selected_cohort={
                            "value": cohort_mapping.get(
                                current_selection, current_selection
                            )
                        },
                        dataset_type=dataset_type,
                    )
                elif connection_type == "box_cnxn":
                    success = add_to_json_box_cnxn(
                        *box_data_package[:8],
                        selected_cohort={
                            "value": cohort_mapping.get(
                                current_selection, current_selection
                            )
                        },
                        dataset_type=dataset_type,
                    )

                if success:
                    print(f"✅ Success! Confirmed and saved: {current_selection}")
                    # Disable the dropdown and button so the selection cannot be changed after confirmation
                    button.disabled = True
                    dropdown.disabled = True

                    if connection_type == "onedrive_cnxn":
                        run_processing_flow(
                            cohort_mapping.get(current_selection, current_selection),
                            box_data_package,
                            dataset_type=dataset_type,
                        )

            else:
                print("⚠️ Please select a cohort before confirming.")

    confirm_button.on_click(on_button_click)

    ui = widgets.VBox([widgets.HBox([dropdown, confirm_button]), output_box])
    display(ui)


def get_dataframe_name_patientList():

    # Ask the user to enter the dataframe name
    df_name = input(
        "Enter a name for the selected dataframe (without the dataset type suffix)."
    )

    return df_name
