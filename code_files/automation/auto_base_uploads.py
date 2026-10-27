# Import the required libraries
import json
import re
import sys
import warnings
import zipfile
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from sqlalchemy import text
from sqlalchemy.exc import SQLAlchemyError

# disable the warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Import the required functions
from code_files.querying.fetch_information import query_patientList
from code_files.utilities.config.mappings import create_column_map
from code_files.utilities.io.readers import find_headerNumber
from code_files.utilities.io.setup_db_cnxn import setup_cnxn
from code_files.utilities.pipelines.upload_pipeline import import_worksheet

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

# define the directory to the json files
json_files_dir = project_root / "json_files"

# define the filepath to table name mapping
filepath_tablename_mapping_path = json_files_dir / "filepath_table_name_mapping.json"

# Define the file paths required
box_followup = Path(
    "C:/Users/AmmashA.IA/Box/Follow up data (Leeds)/Updated follow up data sheets"
)
box_prelim = Path(
    "C:/Users/AmmashA.IA/Box/Preliminary data files (Leeds)/Updated prelim data sheets"
)
nhs_onedrive = Path("D:/AmmashA.IA/OneDrive - NHS/TrisStar metadata")
all_box_files = list(box_followup.glob("*.xlsx")) + list(box_prelim.glob("*.xlsx"))

# Define the directories to the required json files
cohort_names_mapping_path = json_files_dir / "cohort_names_mapping.json"
cohort_schema_mapping_path = json_files_dir / "cohort_schema_mapping.json"
cohortname_cohortID_mapping_path = json_files_dir / "cohortname_cohortID_mapping.json"
file_upload_settings_path = json_files_dir / "json_files" / "file_upload_settings.json"
filepath_tablename_mapping_path = json_files_dir / "filepath_tablename_mapping.json"
filepath_cohortname_mapping_path = json_files_dir / "filepath_cohortname_mapping.json"

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

# Define a dictionary for the configuration information
config_data = {}

# Loop through the json files
for name in json_file_names:
    # Create the json path
    json_path = json_files_dir / f"{name}.json"
    # Load the json file configuration data
    with open(json_path, "r") as f:
        config_data[name] = json.load(f)

# get the transferred cohorts from tristar pathway as well as the patient list
transferred_tristar_pathway = pd.DataFrame(columns=["New_Name"])
valid_patientList = pd.DataFrame()

try:
    query_tristar_pathway = text(
        """
    SELECT * FROM "imp_documents"."tristar_pathway"
    """
    )

    # run the query and load the patient list into a dataframe
    transferred_tristar_pathway = pd.read_sql(query_tristar_pathway, engine)
    valid_patientList = query_patientList(
        where_clause="[FileName] IS NOT NULL AND [DonorID] IS NOT NULL"
    )
except SQLAlchemyError:
    print(
        "The script is not running from the NHS remote PC and cannot access the SQL database."
    )

    if "New_Name" not in transferred_tristar_pathway.columns:
        transferred_tristar_pathway["New_Name"] = pd.Series(dtype="object")

    # strip the file names in the tristar pathway dataset
    transferred_tristar_pathway["New_Name"] = transferred_tristar_pathway["New_Name"].apply(
        lambda x: x.strip(" \n\\n") if isinstance(x, str) else x
    )
    transferred_tristar_pathway["New_Name"] = transferred_tristar_pathway[
        "New_Name"
    ].replace("\\n", "", regex=False)


# Define a function to loop through the box folder and upload any datasets that aren't on the SQL database
def auto_upload():
    """
     The function ensures that everything that is on Box is uploaded to the database. It loops through
     the anonymised files on Box Drive, determines its dataset type and header row, assigns the suffix,
    creates the file name, fetches the cohort and schema names, and uploads the dataset to the database.
    """

    for file in all_box_files:
        # load each workbook in the folder
        wb = load_workbook(file, read_only=True)
        id_column = "DonorID"

        # loop through the sheets in each workbook
        for sheet in wb.sheetnames:
            # Reset the dataset type and dataset suffix variables
            dataset_type = None
            suffix = None

            # if the workbook consists of 1 sheet, then it is a follow up or prelim sheet with no supplementary data
            # set the suffix variable accordingly
            if len(wb.sheetnames) == 1:
                if "Follow up data (Leeds)" in str(file):
                    dataset_type = "anonymised follow up"
                    suffix = "_FollowUp_Anon"
                elif "Preliminary data files" in str(file):
                    dataset_type = "anonymised prelim"
                    suffix = "_Prelim_Anon"
            # if the workbook consists of 2 or more sheets, one of them will be a follow up or prelim sheet and the rest will be
            # supplementary data - set the suffix variable accordingly
            elif len(wb.sheetnames) >= 2:
                if any(
                    kw in sheet.lower() for kw in ("follow up", "anonymised", "anon")
                ):
                    dataset_type = "anonymised follow up"
                    suffix = "_FollowUp_Anon"
                elif "prelim" in sheet.lower():
                    dataset_type = "anonymised prelim"
                    suffix = "_Prelim_Anon"
                elif any(kw in sheet.lower() for kw in ("supplement", "supp")):
                    dataset_type = "anonymised supplementary"
                    suffix = "_Supplementary_Anon"
                elif any(kw in sheet.lower() for kw in ("stain", "met")):
                    dataset_type = "anonymised supplementary"
                    suffix = "_OtherData_Anon"
                # in the rare cases where the workbook is a prelim workbook but has more than one sheet
                else:
                    dataset_type = "anonymised prelim"
                    suffix = "_Prelim_Anon"

            # find the headerNumber of each sheet
            header_number = find_headerNumber(
                filename_main=file, selected_sheet=sheet, dataset_type=dataset_type
            )

            # assign a name to the sheet by first extracting the file name from the file path
            file_name = re.split(r"[/\\/]", str(file))[-1][:-5]
            # define a pattern for the file name - prelim or follow up data files
            pattern = r"(L00.*?)_(Follow|Prelim)"
            # search for matches to the pattern
            match = re.search(pattern, str(file))
            if not match:
                print(f"ℹ️ Skipping the file {file}. Match not found")
                continue
            # extract the file name from the match without the _FollowUp or _Prelim suffix
            file_name = match.group(1)
            # Clean up the file name by replacing any double underscores or extra spaces
            file_name = file_name.replace("__", "_")
            file_name = re.sub(r"_\s+", "_", file_name)
            # get the cohort name from the filepath to cohort name mapping json file
            mapped_cohort_name = config_data["filepath_cohortname_mapping"].get(
                file_name
            )
            # get the schema name from the cohort to schema mapping json file
            schema_name = config_data["cohort_schema_mapping"].get(mapped_cohort_name)

            # create the name for the dataframe
            try:
                df_name_main = mapped_cohort_name + suffix

            except (AttributeError, TypeError):
                sys.stderr.write(
                    "\n===================================================================================================================\n"
                )
                sys.stderr.write(
                    f"Error in Mapping: File is {file_name} Mapped Cohort Name is {mapped_cohort_name} and Suffix is {suffix}. cohort_name is {mapped_cohort_name}. File name is {file_name}"
                )
                sys.stderr.write(
                    "\n===================================================================================================================\n"
                )
                sys.exit(1)  # exit with an error code
                continue

            # if the dataset exists in the json file, don't run the auto base upload
            if df_name_main in config_data["file_upload_settings"]:
                print(f"ℹ️ Skipping {df_name_main} - already exists in the database.")
                continue
            else:
                print(f"Table {df_name_main} does not exist in the database.")
                # run the import worksheet function
                import_worksheet(
                    filename_main=str(file),
                    selected_sheet=sheet,
                    header_row_main=header_number,
                    df_name_main=df_name_main,
                    engine=engine,
                    column_name_map=create_column_map(),
                    dataset_type=dataset_type,
                    id_column=id_column,
                    auto_mode=True,
                    cohort_name=mapped_cohort_name,
                    schema_name=schema_name,
                )
                print(f"Uploaded the data for the cohort {mapped_cohort_name}")


# In[]:
def auto_upload_onedrive():
    """
    This function ensures that all files on the NHS OneDrive are uploaded to the database. It looks at
    all the excel files on OneDrive with 'follow up data' in their names while excluding the specified files.
    The function then extracts the table name from the file name, maps it to the cohort name, gets the schema
    name, and then sets the dataset type as well as suffix variables. The function then creates the sheet name
    and finds the header number before assigning the id_column variable and uploading the dataset to the database.
    """

    # get the list of paths to the excel workbooks on the NHS onedrive
    excel_files = list(nhs_onedrive.rglob("*.xls*"))

    # define the files to be excluded
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
            "Miscellaneous 1",
            "Nivo prelim",
            "Pembro Nivo 2019",
            "pembro prelim",
        ]
    ]
    # filter the files to keep out the ones to be excluded
    filtered_files = [
        file
        for file in excel_files
        if not any(excl in str(file).lower() for excl in exclusions)
        and any(kw in str(file).lower() for kw in ["follow up data", "prelim data"])
    ]

    # loop through the filtered files
    for file in filtered_files:
        # Use pd.ExcelFile to open the file object once in memory for all sheets
        try:
            xl_file = pd.ExcelFile(file, engine="openpyxl")
        except (OSError, ValueError, TypeError, InvalidFileException, zipfile.BadZipFile) as e:
            print(f"ℹ️ Skipping unreadable file {file.name}: {e}")
            continue
        # convert the file name to lower case
        lower_case_file_name = file.name.lower().replace(
            "leeds follow up data", "follow up data"
        )

        # Skip the files which do not contain the word 'follow up data'
        if "follow up data" not in lower_case_file_name:
            continue

        # determine the index of the words 'follow up data'
        kw_index = lower_case_file_name.find("follow up data")
        # extract the table name by splitting at the index of the words 'follow up data' and removing extra spaces
        onedrive_tbl = file.name[:kw_index].strip()
        # if the table is a normal tissue table without the word "tissue", add it
        if "Normal" in onedrive_tbl and "tissue" not in onedrive_tbl.lower():
            onedrive_tbl = onedrive_tbl.replace("Normal", "Normal tissue")
        # get the cohort name by mapping to the onedrive_tbl name
        cohort_name = config_data["cohort_names_mapping"].get(onedrive_tbl)

        # When no cohort name is found in the mapping
        if not cohort_name:
            print(f"⚠️ Missing cohort mapping for: {onedrive_tbl}")
            continue

        # Determine the schema name using the mapping
        schema_name = config_data["cohort_schema_mapping"].get(cohort_name)

        # Loop through the pre-cached sheet names
        for sheet in xl_file.sheet_names:
            # convert each sheet name to lower case
            sheet_lower = sheet.lower()

            # Map dataset type using memory-cached names
            if any(
                kw in sheet_lower for kw in ["anonymized", "follow up", "anonymised"]
            ):
                dataset_type = "unanonymised follow up"
                suffix = "_FollowUp_Unanon"
            elif "prelim" in sheet_lower:
                dataset_type = "unanonymised prelim"
                suffix = "_Prelim_Unanon"
            elif any(
                kw in sheet_lower
                for kw in ["supplement", "mets", "staining", "supplementary"]
            ):
                dataset_type = "unanonymised supplementary"
                suffix = "_Supplementary_Unanon"
            else:
                continue

            # create the sheet name
            sheet_name = cohort_name + suffix

            # Skip checking early if it already exists in the DB
            if (
                sheet_name in config_data["file_upload_settings"]
                or "ovarian500" in sheet_name
            ):
                print(f"ℹ️ Skipping {sheet_name} - already exists in the database.")
                continue

            # Update find_headerNumber to use the opened xl_file object if you can modify it,
            # otherwise it reads the path again.
            header_number = find_headerNumber(
                filename_main=file, selected_sheet=sheet, dataset_type=dataset_type
            )

            # Read in the df
            df = pd.read_excel(xl_file, sheet_name=sheet, header=header_number)
            # assign the id_column variable
            if "donor id" in df.columns.str.lower():
                id_column = "DonorID"
            else:
                id_column = "NHS_Number"

            if df.empty:
                continue

            # Run the import
            print(f"🚀 Processing table: {file.name} and sheet: {sheet_name}")
            print(f"Cohort Name: {cohort_name}")
            print(f"Schema Name: {schema_name}")
            import_worksheet(
                filename_main=str(file),
                selected_sheet=sheet,
                header_row_main=header_number,
                df_name_main=sheet_name,
                engine=engine,
                column_name_map=create_column_map(),
                dataset_type=dataset_type,
                id_column=id_column,
                auto_mode=True,
                cohort_name=cohort_name,
                schema_name=schema_name,
            )
            print("==================================================================")


def start_auto_base_uploads():
    auto_upload()
    auto_upload_onedrive()
