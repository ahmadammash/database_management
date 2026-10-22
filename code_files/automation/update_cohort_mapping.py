# import the required libraries
import json
from pathlib import Path

# Import required functions
from code_files.utilities.io.setup_db_cnxn import setup_cnxn
from code_files.utilities.processing.cleaners import clean_logic

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

# load the cohort mapping from the json file
cohort_names_mapping_path = project_root / "json_files" / "cohort_names_mapping.json"
with open(cohort_names_mapping_path, "r") as f:
    cohort_mapping = json.load(f)


# define a function to get the cohort names as a set from onedrive
def get_onedrive_tbl_names():
    """
    This function fetches the file names from the NHS OneDrive - which are then used to create the cohort names
    mapping json file.
    """
    # define the path to the nhs onedrive folder
    nhs_onedrive = Path("C:/Users/AmmashA.IA/OneDrive - NHS/TrisStar metadata")

    # define the list of paths to the excel workbooks in the NHS OneDrive
    excel_files = list(nhs_onedrive.rglob("*.xls*"))

    # exclude all the unwanted files
    exclusions = [
        excl.lower().strip()
        for excl in [
            "4D Path data",
            "Alia",
            "Angie",
            "Barrett's Cases-4D Path",
            "Box uploads",
            "CellNass requests",
            "Cytology",
            "IHC projects",
            "Important documents",
            "Jess",
            "Long Term Storage",
            "Multi tumour TMA",
            "Narinder",
            "Neuro (Dolapo)",
            "New searches",
            "Nic and Clare Access DB share",
            "Plans",
            "Racheal",
            "Recalled Slides or Blocks",
            "Staff induction forms",
            "Templates",
            "TriStar IHC Projects",
            "Old Sheets",
            "Back ups",
            "BACKUPS",
            "AZ follow up data docs",
            "collection",
            "Book",
            "CRC Slides to Scan",
            "searches",
            "Numbers.xlsx",
            "Numbers (1).xlsx",
            "svs",
            "search",
            "summary",
            "provisional",
            "4D scanning cohorts",
        ]
    ]

    # Define an empty list for the filtered files
    filtered_files = []

    # Define an empty list to store the table names
    onedrive_tbls = []

    # Loop through the excel files and filter out the unwanted files based on the exclusions
    for file in excel_files:
        # convert the file path to lower case
        file_path_str = str(file).lower()
        # exclude the unwanted files
        if any(excl in file_path_str for excl in exclusions):
            continue
        else:
            filtered_files.append(file)

    # Keep the files with the words "follow up data" or "prelim data" in their names
    filtered_files = [
        file
        for file in filtered_files
        if any(kw in str(file).lower() for kw in ["follow up data", "prelim data"])
    ]

    # Loop through the filtered files
    for file in filtered_files:
        # convert the file name to lower case
        lower_case_file_name = file.name.lower()
        keywords = ["leeds follow up data", "follow up data"]
        for kw in keywords:
            if kw in lower_case_file_name:
                # get the index of the keywords "leeds follow up data" or "follow up data" in order to slice the name
                kw_index = lower_case_file_name.find(kw)
                onedrive_tbl = file.name[:kw_index].strip()

                # for the normal tissue cohorts specifically, we need to add the word tissue to the name for it to appear in the cohort name - for consistency
                if (
                    "normal" in onedrive_tbl.lower()
                    and "tissue" not in onedrive_tbl.lower()
                ):
                    onedrive_tbl = onedrive_tbl.replace("Normal", "Normal tissue")
                # Add the one drive table name to the list of one drive tables
                if onedrive_tbl != "":
                    onedrive_tbls.append(onedrive_tbl)
                break
    return onedrive_tbls


# define a function to run the update to the cohort mapping
def update_cohort_mapping():
    """
    This function gets the one drive table names using the get_onedrive_tbl_names function and then cleans the
    table names using clean logic to get the cohort name. The function returns the mapping of the untidy cohort
    to the tidy version. Then, the function dumps the changes to the json file
    """

    # create a new temporary updated mapping dictionary
    cohort_mapping_updated = {}
    # get the onedrive tbl names by running the function
    onedrive_tbls = get_onedrive_tbl_names()
    # Loop through the cohort names in the patient list
    for tbl in onedrive_tbls:
        # set the cleaned name as the value for the original cohort name in the updated mapping dictionary
        cohort_mapping_updated[tbl] = clean_logic(tbl)
    # Dump the changes in the json file
    with open(cohort_names_mapping_path, "w") as f:
        json.dump(cohort_mapping_updated, f, indent=4)
