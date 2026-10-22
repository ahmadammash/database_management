# Import the Required Libraries
import os
import sys

import numpy as np
import pandas as pd

# Import functions from the project directory
from code_files.utilities.io.setup_db_cnxn import setup_cnxn
from code_files.utilities.ui.widgets import (
    add_cohortName_column_live_cnxn,
    get_worksheets_onedrive_cnxn,
)

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

# move to the code_files directory
code_files = project_root / "code_files"

# add the root path to the system
if str(code_files) not in sys.path:
    sys.path.insert(0, str(code_files))


# Define the clinical metadata mapping json file and normalise the file path
json_mapping_path = project_root / "json_files" / "clinical_metadata_mapping.json"

# ------------------------------------------------------------------------------------------------------------------

# Import the patient inventory and get the unique cohort names to use in the cohorts dropdown menu later
# Import the patient inventory worksheet
inventory_excel_path = "C:/Users/AmmashA.IA/OneDrive - NHS/TrisStar metadata/Important documents/Patient list.xlsm"
patient_inventory = pd.read_excel(inventory_excel_path, sheet_name="Patient inventory")

# Adjust the column names of the Patient Inventory Dataset
column_name_map_PatientInventory = {
    "Donor ID": "DonorID",
    "Specimen Number": "SpecimenNumber",
    "NHS Number": "NHS_Number",
    "Sent to TriStar": "SentToTriStar",
    "No. tumour blocks sent to TriStar": "NumTumourBlocksSentToTriStar",
    "No. normal tissue blocks sent to TriStar": "NumNormalBlocksSentToTriStar",
    "Total number of blocks sent to TriStar": "TotalNumBlocksSentToTriStar",
    "Date blocks sent": "DateBlocksSent",
    "Manifest ID": "ManifestID",
    "Cohort": "CohortName",
    "Cohort ID": "CohortID",
    "File name": "FileName",
    "All blocks returned from TriStar?": "AllBlocksReturnedFromTriStar",
    "No. Blocks returned from TriStar": "NumBlocksReturnedFromTriStar",
    "Date blocks returned from TriStar": "DateBlocksReturnedFromTriStar",
    "Tracking number": "TrackingNumber",
    "Block return comments": "BlockReturnComments",
}

# Rename the columns in the Patient Inventory dataset
patient_inventory = patient_inventory.rename(columns=column_name_map_PatientInventory)

# Get the unique cohort names from the Patient List data
cohort_names = patient_inventory["CohortName"].unique().tolist()

# Keep the cohort names other than the ones below - then sort the cohort names
cohort_names_to_remove = [
    "Jess",
    "Katie",
    "Angie",
    "Yong + Inga",
    "Neuro (Dolapo)",
    "Barrett's",
    np.nan,
    "Alia",
    "Combined transfer 1",
    "Long term storage",
]
cohort_names = [name for name in cohort_names if name not in cohort_names_to_remove]
cohort_names = sorted(cohort_names)


# ------------------------------------------------------------------------------------------------------------------


# Define a function to run the upload process, which will be called in the main function
def main():

    # Get the initial file information from the user
    (
        onedrivebox_filepath,
        selected_sheetName_onedrivebox,
        header_row_onedrivebox,
        onedrive_filepath,
        selected_sheetName_onedrive,
        header_row_onedrive,
    ) = get_worksheets_onedrive_cnxn()

    # create the file names for the onedrive and onedrive box sheets from the file paths
    onedrive_filename = os.path.basename(onedrive_filepath)
    onedrivebox_filename = os.path.basename(onedrivebox_filepath)

    # create the box data package containing information on the onedrive and onedrive box data
    box_data_package = (
        onedrivebox_filepath,
        selected_sheetName_onedrivebox,
        header_row_onedrivebox,
        onedrivebox_filename,
        onedrive_filepath,
        selected_sheetName_onedrive,
        header_row_onedrive,
        onedrive_filename,
    )

    # add the cohort name column
    add_cohortName_column_live_cnxn(cohort_names, box_data_package, "onedrive_cnxn")


if __name__ == "__main__":
    main()
