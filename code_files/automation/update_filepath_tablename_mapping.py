# Import required libraries
import json
import re
from pathlib import Path

from openpyxl import load_workbook

# Import the required functions
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

# Define the file paths required
box_followup = Path(
    "C:/Users/AmmashA.IA/Box/Follow up data (Leeds)/Updated follow up data sheets"
)
box_prelim = Path(
    "C:/Users/AmmashA.IA/Box/Preliminary data files (Leeds)/Updated prelim data sheets"
)
all_box_files = list(box_followup.glob("*.xlsx")) + list(box_prelim.glob("*.xlsx"))

# Load the json files
json_file_names = [
    "cohort_names_mapping",
    "cohort_schema_mapping",
    "cohortname_cohortID_mapping",
    "file_upload_settings",
    "filepath_tablename_mapping",
    "filepath_cohortname_mapping",
    "onedrive_cohort_schema_mapping",
]

# define the directory to the json files
json_files_dir = project_root / "json_files"

# define the path to the filepath table name mapping json file
filepath_tablename_mapping_path = json_files_dir / "filepath_tablename_mapping.json"

# Define a dictionary for the configuration information
config_data = {}

# Loop through the json files
for name in json_file_names:
    # Create the json path
    json_path = json_files_dir / f"{name}.json"
    # Load the json file configuration data
    with open(json_path, "r") as f:
        config_data[name] = json.load(f)


# Define a function to create and update the filepath (key) to sql table name mapping - 3rd to run
def update_filepath_tablename_mappings():
    """
    This function loops through the box files, extracts the file name and cleans it before retrieving the
    associated cohort name from the filepath to cohort name mapping dictionary. the function then checks for the
    dataset type and accordingly creates the file name by setting the keys to the mapped cohort name plus the correct suffix
    and the values to the cleaned file name.
    """
    # reset the JSON configuration
    config_data["filepath_tablename_mapping"] = {}

    # create a temporary lower case mapping of the cohort names dictionary
    lower_case_cohort_names_mapping = {
        k.lower().strip(): v for k, v in config_data["cohort_names_mapping"].items()
    }

    # loop through the workbooks in the box file
    for file in all_box_files:
        # load the workbook in read-only mode
        wb = load_workbook(file, read_only=True)
        # extract the file name by cutting out everything before the file name as well as the suffix
        pattern = r"(L00.*?)_(Follow|Prelim)"
        match = re.search(pattern, str(file))
        # if a match is found, extract the file name
        if match:
            file_name = match.group(1)
        # a regex expression to replaces occurences of "2 ()" with "2_()"
        file_name = re.sub(r"(2)\s*(\()", r"\1_\2", file_name)
        # replace double underscores
        file_name = file_name.replace("__", "_")
        # a regex expression to remove extra spaces after underscores
        file_name = re.sub(r"_\s+", "_", file_name, flags=re.UNICODE)
        # map the file_name to the cohort name using the mapping dictionary
        # should just use the file_name until the cohort name ends i.e.
        # use manifest number + tracking number + cohort name
        cohort_name = config_data["filepath_cohortname_mapping"].get(file_name)
        # retrieve the mapped cohort name from the dictionary
        mapped_cohort_name = lower_case_cohort_names_mapping.get(
            cohort_name.strip().lower()
        )
        # when the workbook has 1 sheet and since we are working with the follow up data folder,
        # it would automatically be a follow up sheet or a prelim sheet
        if mapped_cohort_name is not None:
            # if there is 1 sheet in the workbook
            if len(wb.sheetnames) == 1:
                # if the file is from the follow up folder, create the file name with the _FollowUp_Anon suffix
                if "Updated follow up data" in str(file):
                    config_data["filepath_tablename_mapping"][
                        mapped_cohort_name + "_FollowUp_Anon"
                    ] = file_name
                # if the file is from the follow up folder, create the file name with the _Prelim_Anon suffix
                elif "Updated prelim data" in str(file):
                    config_data["filepath_tablename_mapping"][
                        mapped_cohort_name + "_Prelim_Anon"
                    ] = file_name
            else:
                # Otherwise, loop through the sheets in the workbook
                for sheet in wb:
                    # create the file name with the correct suffix depending on the sheet title
                    if any(
                        kw in sheet.title.lower() for kw in ("follow", "anonymised")
                    ):
                        config_data["filepath_tablename_mapping"][
                            mapped_cohort_name + "_FollowUp_Anon"
                        ] = file_name
                    elif any(kw in sheet.title.lower() for kw in ("prelim", "basic")):
                        config_data["filepath_tablename_mapping"][
                            mapped_cohort_name + "_Prelim_Anon"
                        ] = file_name
                    else:
                        config_data["filepath_tablename_mapping"][
                            mapped_cohort_name + "_Supplementary_Anon"
                        ] = file_name

    # dump the changes in the json file
    with open(filepath_tablename_mapping_path, "w") as f:
        json.dump(config_data["filepath_tablename_mapping"], f, indent=4)
