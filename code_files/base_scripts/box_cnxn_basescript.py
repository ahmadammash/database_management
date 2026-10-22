# Import the Required Libraries
import os

import numpy as np
import pandas as pd

from code_files.utilities.io.setup_db_cnxn import setup_cnxn
from code_files.utilities.ui.widgets import (
    add_cohortName_column_live_cnxn,
    get_worksheets_box_cnxn,
)

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

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


#  Main function to run the functions above
def main():
    # call the function to get the worksheets from the user and return the file paths, sheet names, and header row numbers for the onedrive data and the corresponding Box data
    (
        box_filepath,
        selected_sheetName_box,
        header_row_box,
        onedrivebox_filepath,
        selected_sheetName_onedrivebox,
        header_row_onedrivebox,
    ) = get_worksheets_box_cnxn()

    # create the file name from the file path
    box_filename = os.path.basename(box_filepath)
    onedrivebox_filename = os.path.basename(onedrivebox_filepath)

    # create the package containing the information on the box and onedrive data
    box_data_package = (
        box_filepath,
        selected_sheetName_box,
        header_row_box,
        box_filename,
        onedrivebox_filepath,
        selected_sheetName_onedrivebox,
        header_row_onedrivebox,
        onedrivebox_filename,
    )

    # add the cohort name
    add_cohortName_column_live_cnxn(cohort_names, box_data_package, "box_cnxn")


if __name__ == "__main__":
    main()
# %%
