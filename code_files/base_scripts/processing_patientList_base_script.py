# Import the Required Libraries
import pandas as pd

pd.set_option("display.max_columns", None)  # Ensure all columns are displayed

from code_files.utilities.io.json_manager import add_to_json_patientList
from code_files.utilities.io.setup_db_cnxn import setup_cnxn
from code_files.utilities.pipelines.upload_pipeline import import_worksheet_patientList
from code_files.utilities.ui.widgets import (
    get_dataframe_name_patientList,
    select_workbook_main,
)

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

# Define the patient list configuration file
config_file = project_root / "json_files" / "patientList_config.json"


# A Main Function Combining the Functions Above
def main():
    filename_main = select_workbook_main(
        "Please select the workbook containing the Patient List.",
        "D:/AmmashA.IA/OneDrive - NHS/TrisStar metadata/Important documents",
    )
    selected_sheet = "Patient inventory"
    header_row = 1
    df_name = get_dataframe_name_patientList()
    import_worksheet_patientList(
        filename_main, selected_sheet, header_row, df_name, engine
    )

    # Add the configuration settings to the json file
    add_to_json_patientList(filename_main, selected_sheet, df_name)


if __name__ == "__main__":
    main()
