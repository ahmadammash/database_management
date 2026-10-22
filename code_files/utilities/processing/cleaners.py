# Import the required libraries
import re

import numpy as np
import openpyxl
import pandas as pd


def trim_column_names(main_workbook):
    """
    This functions trims extra spaces from the column names in a dataset
    """
    main_workbook.columns = [str(col).strip() for col in main_workbook.columns]
    return main_workbook


# In[15]: Function to trim extra spaces from all row values


def trim_spaces(main_workbook):
    """
    This function trims extra spaces from all rows in a dataset
    """
    main_workbook.columns = [" ".join(col.split()) for col in main_workbook.columns]
    return main_workbook


# ----------------------------------------------------------------------------------------------------------------------------------------

# In[16]: Function to convert the dataset's column names to title case


def title_case(main_workbook):
    """
    This function converts the column names to title case
    """
    main_workbook.columns = [col.title() for col in main_workbook.columns]
    return main_workbook


# ----------------------------------------------------------------------------------------------------------------------------------------

# In[17]: Function to Replace "z" with "s" (standardise how words like Anonymise are spelled)


def replace_z_s(main_workbook):
    """ "
    This function replaces the z with an s in the column names for consistency (e.g., Anonymised vs Anonymized)
    """
    main_workbook.columns = [col.replace("z", "s") for col in main_workbook.columns]
    return main_workbook


# ----------------------------------------------------------------------------------------------------------------------------------------


def rename_columns(main_workbook, column_name_map):
    """
    This function renames the columns using the column name map provided as an input argument
    """
    main_workbook.rename(columns={"Stage At Dx (Ajcc)": "Stage At Dx"})
    main_workbook = main_workbook.rename(columns=column_name_map)
    return main_workbook


# ----------------------------------------------------------------------------------------------------------------------------------------

# In[19]: Function to remove empty rows and complete the dataframe (scan logs)


def remove_empty_and_complete(main_workbook):
    """
    This function removes empty rows from scan logs and completes the rows with a blank value
    """
    # drop the completely empty rows
    main_workbook.dropna(how="all", inplace=True)

    # drop the rows where the LH, Part Number, and Anonymised Code columns are all NaN
    main_workbook.dropna(
        subset=["LH_Number", "Part_Number", "Anonymised_Code"], how="all", inplace=True
    )

    # reset the dataframe index
    main_workbook.reset_index(drop=True, inplace=True)

    # Loop through the dataframe rows to complete the LH_Number column values
    for index, row in main_workbook.iterrows():
        # Check if the Part_Number is not NaN and the LH_Number is NaN
        # replace any empty LH_Number with the LH_Number of the row above it
        if pd.isna(row["LH_Number"]):
            main_workbook.at[index, "LH_Number"] = main_workbook.at[
                index - 1, "LH_Number"
            ]

    # Loop through the dataframe rows and capture the values that do not belong under the LH_Number column
    for index, row in main_workbook.iterrows():
        # Check if the LH_Number value does not start with LH and is not NaN
        if pd.notna(row["LH_Number"]) and not str(row["LH_Number"]).startswith("LH"):
            # move the value to a new column called Indication
            main_workbook.at[index, "Indication"] = row["LH_Number"]
            # Set the LH Number value to NaN
            main_workbook.at[index, "LH_Number"] = pd.NA
            # Copy the Indication value to the row below until a new LH_Number is found
            for i in range(index + 1, len(main_workbook)):
                if str(main_workbook.at[i, "LH_Number"]).startswith("LH") or pd.isna(
                    main_workbook.at[i, "LH_Number"]
                ):
                    main_workbook.at[i, "Indication"] = main_workbook.at[
                        index, "Indication"
                    ]
                else:
                    break

    # Make sure no empty rows are left in the dataframe
    main_workbook.dropna(how="all", inplace=True)

    # Make sure no rows with an empty LH, Part Number, and Anonymised Code are left
    main_workbook.dropna(
        subset=["LH_Number", "Part_Number", "Anonymised_Code"], how="all", inplace=True
    )

    # reset the dataframe index
    main_workbook.reset_index(drop=True, inplace=True)

    return main_workbook


# ----------------------------------------------------------------------------------------------------------------------------------------

# In[20]: Function to analyse the cell colors in the scan logs data files


def analyse_cell_colors(filename_main, selected_sheet, header_row):
    """
    This function looks at the colors of cells and converts the color to a value which is added
    under a newly introduced column to replace the color on the database
    """
    # Load the workbook using openpyxl
    wb = openpyxl.load_workbook(filename_main, data_only=True)
    sheet = wb[selected_sheet]

    # remove empty rows from the excel sheet so the indices match up with main_workbook when analysing cell colors
    header_original_excel = 8
    data_start_original_excel = 8

    # Map the column names from row 8 to their 1-based index numbers
    col_map = {}
    for col_idx in range(1, sheet.max_column + 1):
        cell_val = sheet.cell(row=header_original_excel, column=col_idx).value
        if cell_val:
            col_map[str(cell_val).strip()] = col_idx

    # Grab the column positions
    col_lh = col_map.get("LH Number")

    # read in the dataframe
    main_workbook = pd.read_excel(
        filename_main, sheet_name=selected_sheet, header=header_row
    )

    # trim the spaces from the column names
    main_workbook.columns = [col.strip() for col in main_workbook.columns]

    # Pre-initialize the columns in the DataFrame to prevent fragmenting/assignment issues
    flag_cols = [
        "Recovery",
        "Large_Slide",
        "Scan_Only",
        "Removed",
        "Normal",
        "Non_Block_Case",
    ]
    for col in flag_cols:
        main_workbook[col] = "N/A"

    print(main_workbook.columns)

    # Map the dataframe columns to excel column indices (1-based)
    # Use the original names as they appear in excel before the columns are renamed
    try:
        col_lh = main_workbook.columns.get_loc("LH Number") + 1
        print(f"col_lh: {col_lh}")
        col_part = main_workbook.columns.get_loc("Part Number") + 1
        print(f"col_part: {col_part}")
        col_anon = main_workbook.columns.get_loc("Anonymised Code") + 1
        print(f"col_anon: {col_anon}")
        col_svs = main_workbook.columns.get_loc("Svs Number to transfer") + 1
        print(f"col_svs: {col_svs}")
    except KeyError:
        # Fallback if the column names were not recognised
        return main_workbook

    # Loop through the rows of the dataframe
    for index, row in main_workbook.iterrows():
        # Get the corresponding excel row (1-based)
        excel_row = index + header_row + 2

        # get the anonymised code value in the current row
        anon_code = row["Anonymised Code"]

        # Get the cell fills
        lh_cell = sheet.cell(row=excel_row, column=col_lh)
        print(f"row number: {row}")
        print(f"lh_cell: {lh_cell}")
        part_cell = sheet.cell(row=excel_row, column=col_part)
        print(f"row number: {row}")
        print(f"part_cell: {part_cell}")
        anon_cell = sheet.cell(row=excel_row, column=col_anon)
        print(f"row number: {row}")
        print(f"anon_cell: {anon_cell}")
        svs_cell = sheet.cell(row=excel_row, column=col_svs)
        print(f"row number: {row}")
        print(f"svs_cell: {svs_cell}")

        # Initialize flags
        recovery_flag = large_slide_flag = scan_only_flag = removed_flag = (
            normal_flag
        ) = non_block_case_flag = False

        # Check colors
        for cell in [lh_cell, part_cell, anon_cell, svs_cell]:
            color = cell.fill.start_color.rgb
            print(f"cell {cell} color: {color}")
            if color == "#FFFF00":
                recovery_flag = True
            elif color == "#FFC000":
                large_slide_flag = True
            elif color == "#00B0F0":
                scan_only_flag = True
            elif color == "#FF0000":
                removed_flag = True
            elif color == "#7030A0":
                normal_flag = True
            elif color == "#00B050":
                non_block_case_flag = True

        print(
            "----------------------------------------------------------------------------------"
        )

        # Set the new columns based on flags
        main_workbook.at[index, "Recovery"] = "Yes" if recovery_flag else "N/A"
        main_workbook.at[index, "Large_Slide"] = "Yes" if large_slide_flag else "N/A"
        main_workbook.at[index, "Scan_Only"] = "Yes" if scan_only_flag else "N/A"
        main_workbook.at[index, "Removed"] = "Yes" if removed_flag else "N/A"
        main_workbook.at[index, "Normal"] = "Yes" if normal_flag else "N/A"
        main_workbook.at[index, "Non_Block_Case"] = (
            "Yes" if non_block_case_flag else "N/A"
        )

    return main_workbook


# ----------------------------------------------------------------------------------------------------------------------------------------


def standardise_cohortName_patientList(main_workbook):
    """
    This function filters the patient list so that it only includes the actual cancer cohorts
    """
    # Define the cohort names to be excluded
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

    # Filter the dataset for the rows with the included cohort names
    main_workbook = main_workbook[~main_workbook["Cohort"].isin(cohort_names_to_remove)]

    return main_workbook


def rename_PatientInventory_Columns(main_workbook):
    """
    This function renames the patient list columns using a defined column name mapping
    """
    # Map the current names of the columns in the Patient List to the standardised names
    # Adjust the column names of the Patient Inventory Dataset
    column_name_map_PatientInventory = {
        "Donor Id": "DonorID",
        "Specimen Number": "SpecimenNumber",
        "Nhs Number": "NHS_Number",
        "Sent To Tristar": "SentToTristar",
        "No. Tumour Blocks Sent To Tristar": "NumTumourBlocksSentToTriStar",
        "No. Normal Tissue Blocks Sent To Tristar": "NumNormalBlocksSnetToTriStar",
        "Total Number Of Blocks Sent To Tristar": "TotalNumBlocksSentToTriStar",
        "Date Blocks Sent": "DateBlocksSent",
        "Manifest Id": "ManifestID",
        "Cohort": "CohortName",
        "Cohort Id": "CohortID",
        "File Name": "FileName",
        "All Blocks Returned From Tristar?": "AllBlocksReturnedFromTriStar",
        "No. Blocks Returned From Tristar": "NumBlocksReturnedFromTriStar",
        "Date Blocks Returned From Tristar": "DateBlocksReturnedFromTriStar",
        "Tracking Number": "TrackingNumber",
        "Block Return Comments": "BlockReturnComments",
        "Patient Number": "PatientNumber",
        "Disease Episode": "DiseaseEpisode",
    }

    # Rename the columns in the Patient Inventory dataset
    main_workbook = main_workbook.rename(columns=column_name_map_PatientInventory)
    return main_workbook


# ----------------------------------------------------------------------------------------------------------------------------------------


# define a function to get the cleaned cohort name from the onedrive tbl name
def clean_logic(name):
    """
    This function accepts the name of a cohort and cleans it by removing symbols and spaces, correcting
    typos, and converting "Ship 1" to "S1" - for example - for standardisation.
    """
    # e.g. convert "Ship 1" to "S1"
    name = re.sub(r"\(?[Ss]hip[\s_\-]*(\d+)\)?", r"S\1", name)
    # replace all "&" with "And"
    name = name.replace(" & ", "").replace("&", "")
    # replace all "/" with ""
    name = name.replace("/", "")
    # replace all "\" with ""
    name = name.replace("\\", "")
    name = re.sub(r"[\(\)]", "", name)
    # make replacements for special characters and spaces
    name = (
        name.replace("_", "")
        .replace("-", "")
        .replace(",", "")
        .replace(" ", "")
        .replace("__", "")
    )
    # replace any "efgr" with "egfr"
    name = name.replace("EFGR", "EGFR")
    # replace the repeated Bladder words
    name = name.replace("Bladder1Bladder", "Bladder1")
    # replace spelling errors in cholangiocarcinoma
    name = name.replace("Cholangiocaricinomas", "Cholangiocarcinoma")
    name = name.replace("Cholangiocaricinoma", "Cholangiocarcinoma")
    # convert to lower case
    name = name.lower()
    # return the stripped version of the cohort name
    return name.strip()


# ----------------------------------------------------------------------------------------------------------------------------------------


# define a function to convert column names to title case whilst preserving acronyms
def title_case_preserve_acronyms(text: str) -> str:
    # Match all words and non-words
    tokens = re.findall(r"\S+|\s+", text)
    formatted_tokens = []

    for token in tokens:
        # check how many capital letters are in the word
        cap_count = sum(1 for char in token if char.isupper())

        # if the word has 2 or more letters, do not convert it to title case
        if cap_count >= 2:
            formatted_tokens.append(token)
        else:
            formatted_tokens.append(token.capitalize())

    return "".join(formatted_tokens)
