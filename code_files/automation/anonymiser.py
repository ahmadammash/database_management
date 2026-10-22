# Import the required libraries
import json
import os
from typing import Annotated

import numpy as np
import pandas as pd
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# Import the required functions
from code_files.utilities.processing.cleaners import title_case_preserve_acronyms

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

# Define the path to the JSON file for the cohort_mapping
json_path = project_root / "json_files" / "cohort_names_mapping.json"

# Load the cohort mapping json file
with open(json_path, "r") as f:
    cohort_mapping = json.load(f)

# query the master list from the database
query = """SELECT * FROM imp_documents.master_list_long"""
master_list = pd.read_sql(query, cnxn)

# Define a function to anonymise the data
def anonymiser(
    source_sheet_name: Annotated[
        str, "The name of the sheet to anonymise in NHS OneDrive Box Uploads file"
    ],
    df_onedrive: Annotated[
        pd.DataFrame, "The dataframe from the NHS OneDrive Box Uploads file"
    ],
    cohort_name: Annotated[
        str, "The cohort name associated with the dataframe to be anonymised"
    ],
    source_file_name: Annotated[
        str, "The name of the file from the NHS OneDrive Box Uploads file"
    ],
    testing_mode: Annotated[bool, "Set to true if the function is being tested"],
    skip_box_upload: Annotated[
        bool, "Set to true in order to skip the Box upload step"
    ],
):
    """
    This function anonymises datasets and makes specific corrections to the provided follow up, prelim, or
    supplementary data and standardises its format before anonymising and formatting it, adding the Donor IDs, and uploading it back to Box.
    """

    # assign a value for the dataset_type variable depending on the nature of the dataset used
    if "_Prelim_" in source_file_name and "prelim" in source_sheet_name.lower():
        dataset_type = "anonymised prelim"
    elif (
        any(x in source_file_name.lower() for x in ["_follow_up_", "follow up", "anon"])
        and "follow up" in source_sheet_name.lower()
    ):
        dataset_type = "anonymised follow up"
    else:
        dataset_type = "anonymised supplementary"

    # Create an independent copy of the dataframe by setting deep=True
    df_box = df_onedrive.copy(deep=True)

    # trim spaces of all columns in the onedrive sheet
    df_box = df_box.applymap(lambda x: x.strip() if isinstance(x, str) else x)

    # Convert the columns to the type general
    df_box = df_box.astype("object")

    # follow up and prelim data: rename the Stage At DX (AJCC) column to Stage At DX
    if dataset_type in ["anonymised prelim", "anonymised follow up"]:
        df_box.rename(columns={"Stage At DX (AJCC)": "Stage At DX"}, inplace=True)

    # Convert the header row of the OneDrive sheet to title case while preserving the acronyms
    df_box.columns = [title_case_preserve_acronyms(col) for col in df_box.columns]

    # Replace occurences of "Finish Date" in the OneDrive sheet with "End Date" - if there are any dates
    df_box.rename(columns=lambda x: x.replace("Finish Date", "End Date"), inplace=True)

    # Convert all the date columns back to date in the format d/m/yyyy - make sure to exclude any TTE values
    date_columns = [
        col
        for col in df_box.columns
        if "date" in col.lower() and "tte" not in col.lower()
    ]
    for col in date_columns:
        df_box[col] = pd.to_datetime(df_box[col], errors="coerce").dt.strftime(
            "%d/%m/%Y"
        )

    # Convert all TTE columns to numeric type with two decimal places
    tte_columns = [col for col in df_box.columns if "tte_" in col.lower()]
    for col in tte_columns:
        df_box[col] = pd.to_numeric(df_box[col], errors="coerce").round(2)

    # re-create a fresh version of the dataframe
    df_box = df_box.copy()

    # For follow up and prelim data - Add the Donor ID column - set all its values to blank initially
    if not any("donor" in col.lower() for col in df_box.columns):
        df_box["Donor ID"] = pd.NA

    # Move the Donor ID column to the front by adding the columns to a list, inserting the DonorID column, and then resetting the columns order to the new one
    if dataset_type in ["anonymised follow up", "anonymised prelim"]:
        cols = df_box.columns.tolist()
        cols.insert(0, cols.pop(cols.index("Donor ID")))
        df_box = df_box[cols]

    # populate the Donor ID column in the Box sheet by matching the cohort name and specimen number with the master list
    specimen_cols = [
        col
        for col in df_box.columns
        if any(
            x in col.lower()
            for x in [
                "selected specimen anonymised id",
                "id matched met",
                "id normal tissue",
            ]
        )
    ]
    # Iterate through the dataframe
    for index, row in df_box.iterrows():
        # Set the cohort name equal to the cohort name provided to the function as an input parameter
        cohort = cohort_name
        # Keep the specimen IDs which are not null
        valid_specimens = row[specimen_cols].dropna()
        # If there is at least one valid specimen, take the first one. Otherwise, set it to None
        if not valid_specimens.empty:
            specimen_number = valid_specimens.values[0]
        else:
            specimen_number = None
        # find the matching row in the master list
        matching_row = master_list[
            (master_list["CohortName"] == cohort)
            & (master_list["AnonymisedID"] == specimen_number)
        ]
        # if the matching row is found, populate the donor id column
        if not matching_row.empty:
            # get the matching donor id and add it to the dataframe row at the same index
            donor_id = matching_row["DonorID"].values[0]
            df_box.at[index, "Donor ID"] = donor_id

    # If the column is a date column, apply the formula to convert it to Quarter-Year format - if there is a date column in the supplementary data
    # keep N/A values as N/A
    for col in date_columns:
        df_box[col] = pd.to_datetime(df_box[col], errors="coerce", dayfirst=True)
        df_box[col] = df_box[col].map(
            lambda x: f"Q{int(x.quarter)} - {int(x.year)}" if pd.notnull(x) else "N/A"
        )

    # Remove the sensitive columns
    cols_to_drop = [
        col for col in df_box.columns if "nhs" in col.lower() or "dob" in col.lower()
    ]
    df_box.drop(columns=cols_to_drop, inplace=True)

    # Define the date columns that should be kept in the dataframe
    date_columns_to_keep = [
        "Specimen Surgery Date",
        "Primary Surgery Date",
        "Matched Met 1 Surgery Date",
        "Matched Met 2 Surgery Date",
        "Matched Met 3 Surgery Date",
        "Matched Met 4 Surgery Date",
        "Matched Met 5 Surgery Date",
        "Matched Normal Tissue Surgery Date",
        "Date Of DX",
        "TTE",
    ]

    # reset the columns to drop list
    cols_to_drop = []

    # Loop through the columns and pick out the ones that should be deleted
    for col in df_box.columns:
        col_upper = col.upper()
        # columns that fall under the Date category are those containing the word Date as well as the TTE columns
        is_date = "DATE" in col_upper and "TTE" not in col_upper
        keep = (
            any(
                date_to_keep.upper() in col_upper
                for date_to_keep in date_columns_to_keep
            )
            or "TTE" in col_upper
        )

        # Columns that are under the dates category and are not in the "to keep" list should be removed
        if is_date and not keep:
            cols_to_drop.append(col)

    # drop the columns to be deleted
    df_box.drop(columns=cols_to_drop, inplace=True)

    # replace all blank values with N/A
    df_box.fillna("N/A", inplace=True)

    # replace the z with s in the column names - where applicable in the supplementary data
    df_box.rename(
        columns=lambda x: x.replace("Anonymized", "Anonymised") if "z" in x else x,
        inplace=True,
    )

    # replace "Finish Date" with "End Date" in the column names - where applicable in the supplementary data
    df_box.rename(
        columns=lambda x: (
            x.replace("Finish Date", "End Date") if "Finish Date" in x else x
        ),
        inplace=True,
    )

    # replace "Dx" with "DX" in the column names
    df_box.rename(
        columns=lambda x: x.replace("Dx", "DX") if "Dx" in x else x, inplace=True
    )

    # replace the Id in ID columns with ID
    df_box.rename(
        columns=lambda x: x.replace("Id", "ID") if "Anonymised" in x else x,
        inplace=True,
    )

    # make replacements for the TNM columns
    df_box.rename(
        columns=lambda x: x.replace("Tnm", "TNM") if "Tnm" in x else x, inplace=True
    )
    df_box.rename(
        columns=lambda x: x.replace("Pt", "pT") if "Pt" in x else x, inplace=True
    )
    df_box.rename(
        columns=lambda x: x.replace("Pn", "pN") if "Pn" in x else x, inplace=True
    )
    df_box.rename(
        columns=lambda x: x.replace("Pm", "pM") if "Pm" in x else x, inplace=True
    )
    df_box.rename(
        columns=lambda x: x.replace("Cm", "cM") if "Cm" in x else x, inplace=True
    )
    df_box.rename(
        columns=lambda x: x.replace("Donor Id", "Donor ID") if "Donor" in x else x,
        inplace=True,
    )

    # return the appropriate dataframe by keeping the correct columns depending on the dataset type
    # for prelim files, remove all columns after the Stage At DX
    if dataset_type == "anonymised prelim":
        # drop all columns after the Stage at DX column
        df_box = df_box.loc[:, :"Stage At DX"]

    # Remove duplicated columns
    df_box = df_box.loc[:, ~df_box.columns.duplicated()]

    # When box upload is required, run the section below
    if not skip_box_upload:
        # If the function is being tested, the output path is the same and will be the testing folder on Box
        if testing_mode:
            output_path_box = f"C:/Users/AmmashA.IA/OneDrive - University of Leeds/UOL Data Manager/database_management/data_files/leeds_upload_testing/{source_file_name}"
            output_path_onedrivebox = f"C:/Users/AmmashA.IA/OneDrive - University of Leeds/UOL Data Manager/database_management/data_files/leeds_upload_testing/{source_file_name}"
            # create the directories
            os.makedirs(os.path.dirname(output_path_box), exist_ok=True)
            os.makedirs(os.path.dirname(output_path_onedrivebox), exist_ok=True)

        # Define the output path for anonymised follow up data
        elif dataset_type == "anonymised follow up" and not testing_mode:
            output_path_box = f"C:/Users/AmmashA.IA/Box/Follow up data (Leeds)/Updated follow up data sheets/{source_file_name}"
            output_path_onedrivebox = f"C:/Users/AmmashA.IA/OneDrive - NHS/TrisStar metadata/Box uploads/Updated template/Follow up/{source_file_name}"
            # create the directories for the follow up files
            os.makedirs(os.path.dirname(output_path_box), exist_ok=True)
            os.makedirs(os.path.dirname(output_path_onedrivebox), exist_ok=True)

        # Define the output path for anonymised prelim data
        elif dataset_type == "anonymised prelim" and not testing_mode:
            output_path_box = f"C:/Users/AmmashA.IA/Box/Preliminary data files (Leeds)/Updated prelim data sheets/{source_file_name}"
            output_path_onedrivebox = f"C:/Users/AmmashA.IA/OneDrive - NHS/TrisStar metadata/Box uploads/Updated template/Prelim/{source_file_name}"
            # create the directories for the prelim files
            os.makedirs(os.path.dirname(output_path_box), exist_ok=True)
            os.makedirs(os.path.dirname(output_path_onedrivebox), exist_ok=True)

        # Define the output path for anonymised supplementary data
        elif dataset_type == "anonymised supplementary" and not testing_mode:
            output_path_box = f"C:/Users/AmmashA.IA/Box/Follow up data (Leeds)/Updated follow up data sheets/{source_file_name}"
            output_path_onedrivebox = f"C:/Users/AmmashA.IA/OneDrive - NHS/TrisStar metadata/Box uploads/Updated template/Follow up/{source_file_name}"
            # create the directories for the supplementary up files
            os.makedirs(os.path.dirname(output_path_box), exist_ok=True)
            os.makedirs(os.path.dirname(output_path_onedrivebox), exist_ok=True)

        # Uploading the anonymised Supplementary data
        if dataset_type == "anonymised supplementary":
            with pd.ExcelWriter(
                output_path_box, engine="openpyxl", mode="a", if_sheet_exists="replace"
            ) as writer:
                # Upload the dataframe to Box
                df_box.to_excel(writer, sheet_name=source_sheet_name, index=False)

                # define the worksheet object from the writer
                worksheet = writer.sheets[source_sheet_name]

                # if a Supplementary table exists in the workbook, delete it so it is replaced with the updated data
                if "Supplementary" in worksheet.tables:
                    del worksheet.tables["Supplementary"]

                # define the range of the table and the column settings as the shape of the dataframe - starting from cell A1
                (max_row, max_col) = df_box.shape
                table_range = f"A1:{get_column_letter(max_col)}{max_row + 1}"

                # create the supplementary data table and style it
                tab = Table(displayName="Supplementary", ref=table_range)
                style = TableStyleInfo(
                    name="TableStyleLight1",
                    showFirstColumn=False,
                    showLastColumn=False,
                    showRowStripes=False,
                    showColumnStripes=False,
                )
                tab.tableStyleInfo = style
                worksheet.add_table(tab)

                # create the fill color for the header row
                grey_fill = PatternFill(
                    start_color="00E7E6E6", end_color="00E7E6E6", fill_type="solid"
                )

                # fill the column headers with the grey fill
                for col_idx in range(1, max_col + 1):
                    worksheet.cell(row=1, column=col_idx).fill = grey_fill

                # set the column widths
                for i in range(1, max_col + 1):
                    worksheet.column_dimensions[get_column_letter(i)].width = 25

            return df_box

        # Uploading follow up and prelim data
        # Convert the dataframe to an excel table range format
        for path in [output_path_box, output_path_onedrivebox]:
            # Re-upload the newly anonymised data to the Box folder
            with pd.ExcelWriter(
                path, engine="openpyxl", mode="a", if_sheet_exists="replace"
            ) as writer:
                # Write the dataframe to the specified sheet in the output Excel file
                df_box.to_excel(writer, sheet_name=source_sheet_name, index=False)

                # define the worksheet object from the writer
                worksheet = writer.sheets[source_sheet_name]

                # If a table named "FollowUp" already exists, delete it
                if "FollowUp" in worksheet.tables:
                    del worksheet.tables["FollowUp"]

                # define the range of the table and the column settings
                (max_row, max_col) = df_box.shape
                table_range = f"A1:{get_column_letter(max_col)}{max_row + 1}"
                tab = Table(
                    displayName="FollowUp"
                    if dataset_type == "anonymised follow up"
                    else "Prelim",
                    ref=table_range,
                )
                style = TableStyleInfo(
                    name="TableStyleLight1",
                    showFirstColumn=False,
                    showLastColumn=False,
                    showRowStripes=False,
                    showColumnStripes=False,
                )
                tab.tableStyleInfo = style
                worksheet.add_table(tab)

                # create the fill colors
                yellow_fill = PatternFill(
                    start_color="00FFFF00", end_color="00FFFF00", fill_type="solid"
                )
                grey_fill = PatternFill(
                    start_color="00E7E6E6", end_color="00E7E6E6", fill_type="solid"
                )
                blue_fill = PatternFill(
                    start_color="009BC2E6", end_color="009BC2E6", fill_type="solid"
                )
                light_red_fill = PatternFill(
                    start_color="00F8CBAD", end_color="00F8CBAD", fill_type="solid"
                )
                light_orange_fill = PatternFill(
                    start_color="00FFD966", end_color="00FFD966", fill_type="solid"
                )
                dark_orange_fill = PatternFill(
                    start_color="00FFC000", end_color="00FFC000", fill_type="solid"
                )
                dark_grey_fill = PatternFill(
                    start_color="00ACB9CA", end_color="00ACB9CA", fill_type="solid"
                )
                pink_fill = PatternFill(
                    start_color="00FF8AD8", end_color="00FF8AD8", fill_type="solid"
                )
                light_green_fill = PatternFill(
                    start_color="00A9D08E", end_color="00A9D08E", fill_type="solid"
                )

                # fill the column headers in the follow up data based on their names
                # for prelim data, the column headers should all have a yellow fill

                # all columns from Donor ID until before Any Neoadjuvant Therapy should be yellow
                start_column = "Donor ID"
                end_column = "Any Neoadjuvant Therapy"
                start_idx = None
                end_idx = None

                # work with the prelim data case first - where all columns are yellow
                if dataset_type == "anonymised prelim":
                    for col_idx in range(1, max_col + 1):
                        worksheet.cell(row=1, column=col_idx).fill = yellow_fill
                else:
                    # if the dataset is not of the prelim type, loop through the header row
                    # for each column name, check whether its "Donor ID" and "Any Neoadjuvant Therapy"
                    # assign the start and end indices accordingly
                    for col_idx in range(1, max_col + 1):
                        cell_value = worksheet.cell(row=1, column=col_idx).value

                        if cell_value == start_column:
                            start_idx = col_idx
                        if cell_value == end_column:
                            end_idx = col_idx - 1

                    # having assigned the start and end indices, apply the yellow fill to the header row columns between the start and end indices
                    if start_idx and end_idx:
                        first = min(start_idx, end_idx)
                        last = max(start_idx, end_idx)
                        for col_idx in range(first, last + 1):
                            worksheet.cell(row=1, column=col_idx).fill = yellow_fill

                    # all columns from Any Neoadjuvant Therapy until before Adjuvant Tx1 should be grey
                    start_column = "Any Neoadjuvant Therapy"
                    end_column = "Adjuvant Tx1"
                    start_idx = None
                    end_idx = None

                    # Loop through the header row and determine the indices of the start and end columns in this case
                    for col_idx in range(1, max_col + 1):
                        cell_value = worksheet.cell(row=1, column=col_idx).value

                        if cell_value == start_column:
                            start_idx = col_idx
                        if cell_value == end_column:
                            end_idx = col_idx - 1

                    # having determined the start and end indices, apply the grey fill to the columns in between those two
                    if start_idx and end_idx:
                        first = min(start_idx, end_idx)
                        last = max(start_idx, end_idx)
                        for col_idx in range(first, last + 1):
                            worksheet.cell(row=1, column=col_idx).fill = grey_fill

                    # all columns from Adjuvant Tx1 until before Sites/Recurrence Progression 1 should be blue
                    start_column = "Adjuvant Tx1"
                    end_column = "Sites Recurrence/ Progression 1"
                    start_idx = None
                    end_idx = None

                    # Loop through the header row and determine the indices of the start and end columns in this case
                    for col_idx in range(1, max_col + 1):
                        cell_value = worksheet.cell(row=1, column=col_idx).value

                        if cell_value == start_column:
                            start_idx = col_idx
                        if cell_value == end_column:
                            end_idx = col_idx - 1

                    # having determined the start and end indices, apply the grey fill to the columns in between those two
                    if start_idx and end_idx:
                        first = min(start_idx, end_idx)
                        last = max(start_idx, end_idx)
                        for col_idx in range(first, last + 1):
                            worksheet.cell(row=1, column=col_idx).fill = blue_fill

                    # all columns from Sites/Recurrence Progression 1 until before Subsequent Tx1 should be light red
                    start_column = "Sites Recurrence/ Progression 1"
                    end_column = "Subsequent Tx1"
                    start_idx = None
                    end_idx = None

                    # Loop through the header row and determine the indices of the start and end columns in this case
                    for col_idx in range(1, max_col + 1):
                        cell_value = worksheet.cell(row=1, column=col_idx).value

                        if cell_value == start_column:
                            start_idx = col_idx
                        if cell_value == end_column:
                            end_idx = col_idx - 1

                    # having determined the start and end indices, apply the grey fill to the columns in between those two
                    if start_idx and end_idx:
                        first = min(start_idx, end_idx)
                        last = max(start_idx, end_idx)
                        for col_idx in range(first, last + 1):
                            worksheet.cell(row=1, column=col_idx).fill = light_red_fill

                    # all columns from Subsequent Tx1 until before Current Status should be light orange
                    start_column = "Subsequent Tx1"
                    end_column = "Current Status"
                    start_idx = None
                    end_idx = None

                    # Loop through the header row and determine the indices of the start and end columns in this case
                    for col_idx in range(1, max_col + 1):
                        cell_value = worksheet.cell(row=1, column=col_idx).value

                        if cell_value == start_column:
                            start_idx = col_idx
                        if cell_value == end_column:
                            end_idx = col_idx - 1

                    # having determined the start and end indices, apply the grey fill to the columns in between those two
                    if start_idx and end_idx:
                        first = min(start_idx, end_idx)
                        last = max(start_idx, end_idx)
                        for col_idx in range(first, last + 1):
                            worksheet.cell(
                                row=1, column=col_idx
                            ).fill = light_orange_fill

                    # all columns from Current Status until before Comments From Clinical Team should be blue
                    start_column = "Current Status"
                    end_column = "Comments From Clinical Team"
                    start_idx = None
                    end_idx = None

                    # Loop through the header row and determine the indices of the start and end columns in this case
                    for col_idx in range(1, max_col + 1):
                        cell_value = worksheet.cell(row=1, column=col_idx).value

                        if cell_value == start_column:
                            start_idx = col_idx
                        if cell_value == end_column:
                            end_idx = col_idx - 1

                    # having determined the start and end indices, apply the grey fill to the columns in between those two
                    if start_idx and end_idx:
                        first = min(start_idx, end_idx)
                        last = max(start_idx, end_idx)
                        for col_idx in range(first, last + 1):
                            worksheet.cell(row=1, column=col_idx).fill = blue_fill

                    # the comments from clinical team column should be dark orange
                    for col_idx in range(1, max_col + 1):
                        cell_value = worksheet.cell(row=1, column=col_idx).value

                        if cell_value == "Comments From Clinical Team":
                            worksheet.cell(
                                row=1, column=col_idx
                            ).fill = dark_orange_fill

                    # all columns from TTE Neoadjuvant Tx1 Start Date until before TTE Overall Survival (OS) In Months should be dark grey
                    start_column = "TTE Neoadjuvant Tx1 Start Date"
                    end_column = "TTE Overall Survival (OS) In Months"
                    start_idx = None
                    end_idx = None

                    # Loop through the header row and determine the indices of the start and end columns in this case
                    for col_idx in range(1, max_col + 1):
                        cell_value = worksheet.cell(row=1, column=col_idx).value

                        if cell_value == start_column:
                            start_idx = col_idx
                        if cell_value == end_column:
                            end_idx = col_idx - 1

                    # having determined the start and end indices, apply the grey fill to the columns in between those two
                    if start_idx and end_idx:
                        first = min(start_idx, end_idx)
                        last = max(start_idx, end_idx)
                        for col_idx in range(first, last + 1):
                            worksheet.cell(row=1, column=col_idx).fill = dark_grey_fill

                    # columns TTE Overall Survival (OS) In Months and Event (Death; All Causes) should be pink
                    for col_idx in range(1, max_col + 1):
                        cell_value = worksheet.cell(row=1, column=col_idx).value

                        if cell_value in [
                            "TTE Overall Survival (OS) In Months",
                            "Event (Death; All Causes)",
                        ]:
                            worksheet.cell(row=1, column=col_idx).fill = pink_fill

                    # columns TTE Progression Free Survival (PFS) In Months and Event (Recurrence Or Cancer Death) should be light green
                    for col_idx in range(1, max_col + 1):
                        cell_value = worksheet.cell(row=1, column=col_idx).value

                        if cell_value in [
                            "TTE Progression Free Survival (PFS) In Months",
                            "Event (Recurrence Or Cancer Death)",
                        ]:
                            worksheet.cell(
                                row=1, column=col_idx
                            ).fill = light_green_fill

                    # columns Event (Recurrence/Progression) and Event (Cancer Death) should be grey
                    for col_idx in range(1, max_col + 1):
                        cell_value = worksheet.cell(row=1, column=col_idx).value

                        if cell_value in [
                            "Event (Recurrence/Progression)",
                            "Event (Cancer Death)",
                        ]:
                            worksheet.cell(row=1, column=col_idx).fill = grey_fill

                    # round down all TTE column values to 2 decimal places
                    tte_columns = [
                        col for col in df_box.columns if "tte" in col.lower()
                    ]
                    for col in tte_columns:
                        s = pd.to_numeric(df_box[col], errors="coerce")
                        df_box[col] = np.floor(s * 100) / 100

                # set the column widths
                for i in range(1, max_col + 1):
                    worksheet.column_dimensions[get_column_letter(i)].width = 25

    return df_box


# -----------------------------------------------------------------------------------------------------------------------------------------------

# df_prelim_test = pd.read_excel("C:/Users/AmmashA.IA/OneDrive - NHS/TrisStar metadata/CRC's/CRC 1/CRC 1 Leeds follow up data V2.xlsx", sheet_name="CRC 1 Follow up", header=2)
# df_followup_test = pd.read_excel("C:/Users/AmmashA.IA/OneDrive - NHS/TrisStar metadata/NSCLC/NSCLC 3/NSCLC 3 (ADC met) Leeds follow up data V3.xlsx", sheet_name="NSCLC ADC Met Follow Up", header=2)
# df_supp_test = pd.read_excel("C:/Users/AmmashA.IA/OneDrive - NHS/TrisStar metadata/NSCLC/NSCLC 3/NSCLC 3 (ADC met) Leeds follow up data V3.xlsx", sheet_name="NSCLC ADC Met Supplementary", header=0)


# prelim_anon_test = anonymiser(
#     source_sheet_name="test_prelim_sheet_name",
#     df_onedrive=df_prelim_test,
#     cohort_name="crc1",
#     source_file_name="test_Prelim_file.xlsx",
#     testing_mode=True

# )

# followup_anon_test = anonymiser(
#     source_sheet_name="Follow up",
#     df_onedrive=df_followup_test,
#     cohort_name="nsclc3adcmet",
#     source_file_name="test_Follow_Up_file.xlsx",
#     testing_mode=True
# )

# supplementary_anon_test = anonymiser(
#     source_sheet_name="Supplementary",
#     df_onedrive=df_supp_test,
#     cohort_name="nsclc3adcmet",
#     source_file_name="test_Follow_Up_file.xlsx",
#     testing_mode=True

# )
