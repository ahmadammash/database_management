# Import the Required Libraries
import warnings

# Suppress warnings
categories = [UserWarning, DeprecationWarning, FutureWarning]
for cat in categories:
    warnings.filterwarnings("ignore", category=cat)

# Import the required functions
from code_files.utilities.config.mappings import create_column_map
from code_files.utilities.io.readers import find_headerNumber, show_worksheet_names
from code_files.utilities.io.setup_db_cnxn import setup_cnxn
from code_files.utilities.pipelines.upload_pipeline import import_worksheet
from code_files.utilities.ui.widgets import (
    get_dataframe_name_main,
    select_workbook_main,
    select_worksheet_main,
)

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

# ------------------------------------------------------------------------------------------------------------------

# Define the JSON File Location Where the SQL Table Settings Chosen by the User will be Stored
settings_file_path = project_root / "json_files" / "file_upload_settings.json"

# ------------------------------------------------------------------------------------------------------------------


# In[]:
# Define a Main Function Combining the Functions Above
def main():
    # Select Workbook and get Sheet Names (Static Step)
    filename_main = select_workbook_main(
        "Please select the workbook containing the worksheet to upload to the database.",
        "C:/Users/Ammash.IA",
    )
    sheet_names_main = show_worksheet_names(filename_main)

    # Get the Dropdown Widget and the Storage Dictionary
    selected_sheet = select_worksheet_main(sheet_names_main)

    # Get the name of the dataframe
    df_name, dataset_type, id_column = get_dataframe_name_main()

    # Identify the header row number in the selected sheet
    header_row = find_headerNumber(filename_main, selected_sheet, dataset_type)

    # Import and upload the selected dataframe with the selected header row and name
    import_worksheet(
        filename_main,
        selected_sheet,
        header_row,
        df_name,
        engine,
        create_column_map(),
        dataset_type,
        id_column,
        auto_mode=False,
    )


if __name__ == "__main__":
    main()

# %%
