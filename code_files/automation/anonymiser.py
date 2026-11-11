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
    This function anonymises datasets, standardises formats, matches Donor IDs,
    and writes the anonymised data to a NEW separate worksheet in the workbook.
    """

    # Assign dataset_type
    if "_Prelim_" in source_file_name and "prelim" in source_sheet_name.lower():
        dataset_type = "anonymised prelim"
    elif (
        any(x in source_file_name.lower() for x in ["_follow_up_", "follow up", "anon"])
        and "follow up" in source_sheet_name.lower()
    ):
        dataset_type = "anonymised follow up"
    else:
        dataset_type = "anonymised supplementary"

    # Define the new destination sheet name
    target_sheet_name = f"{source_sheet_name}_Anon"

    # Create an independent copy
    df_box = df_onedrive.copy(deep=True)

    # Trim spaces across all string columns
    df_box = df_box.map(lambda x: x.strip() if isinstance(x, str) else x)

    # Convert columns to object type
    df_box = df_box.astype("object")

    # Rename Stage At DX (AJCC)
    if dataset_type in ["anonymised prelim", "anonymised follow up"]:
        df_box.rename(columns={"Stage At DX (AJCC)": "Stage At DX"}, inplace=True)

    # Convert header row to title case preserving acronyms
    df_box.columns = [title_case_preserve_acronyms(col) for col in df_box.columns]

    # Replace Finish Date with End Date
    df_box.rename(columns=lambda x: x.replace("Finish Date", "End Date"), inplace=True)

    # Convert all date columns back to d/m/YYYY
    date_columns = [
        col
        for col in df_box.columns
        if "date" in col.lower() and "tte" not in col.lower()
    ]
    for col in date_columns:
        df_box[col] = pd.to_datetime(df_box[col], errors="coerce").dt.strftime(
            "%d/%m/%Y"
        )

    # Convert TTE columns to numeric rounded to 2 decimal places
    tte_columns = [col for col in df_box.columns if "tte_" in col.lower()]
    for col in tte_columns:
        df_box[col] = pd.to_numeric(df_box[col], errors="coerce").round(2)

    df_box = df_box.copy()

    # Add Donor ID column if missing
    if not any("donor" in col.lower() for col in df_box.columns):
        df_box["Donor ID"] = pd.NA

    # Move Donor ID to the front
    if dataset_type in ["anonymised follow up", "anonymised prelim"]:
        cols = df_box.columns.tolist()
        cols.insert(0, cols.pop(cols.index("Donor ID")))
        df_box = df_box[cols]

    # Populate Donor ID from master_list
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

    for index, row in df_box.iterrows():
        cohort = cohort_name
        valid_specimens = row[specimen_cols].dropna()
        specimen_number = valid_specimens.values[0] if not valid_specimens.empty else None

        matching_row = master_list[
            (master_list["CohortName"] == cohort)
            & (master_list["AnonymisedID"] == specimen_number)
        ]
        if not matching_row.empty:
            donor_id = matching_row["DonorID"].values[0]
            df_box.at[index, "Donor ID"] = donor_id

    # Format date columns into Quarter-Year
    for col in date_columns:
        dt_s = pd.to_datetime(df_box[col], errors="coerce", dayfirst=True)
        df_box[col] = dt_s.map(
            lambda x: f"Q{int(x.quarter)} {int(x.year)}" if pd.notnull(x) else "N/A"
        )

    # Drop direct identifier columns (NHS, DOB)
    cols_to_drop = [
        col for col in df_box.columns if "nhs" in col.lower() or "dob" in col.lower()
    ]
    df_box.drop(columns=cols_to_drop, inplace=True)

    # Filter date columns to keep
    date_columns_to_keep = [
        "Specimen Surgery Date",
        "Primary Surgery Date",
        "Matched Met 1 Surgery Date",
        "Matched Met 2 Surgery Date",
        "Matched Met 3 Surgery Date",
        "Matched Met 4 Surgery Date",
        "Matched Met 5 Surgery Date",
        "Matched Met 6 Surgery Date",
        "Matched Met 7 Surgery Date",
        "Matched Normal Tissue Surgery Date",
        "Date Of DX",
        "TTE",
    ]

    cols_to_drop = []
    for col in df_box.columns:
        col_upper = col.upper()
        is_date = "DATE" in col_upper and "TTE" not in col_upper
        keep = (
            any(date_to_keep.upper() in col_upper for date_to_keep in date_columns_to_keep)
            or "TTE" in col_upper
        )
        if is_date and not keep:
            cols_to_drop.append(col)

    df_box.drop(columns=cols_to_drop, inplace=True)
    df_box.fillna("N/A", inplace=True)

    # Standardise column naming conventions
    df_box.rename(
        columns=lambda x: x.replace("Anonymized", "Anonymised") if "z" in x else x,
        inplace=True,
    )
    df_box.rename(
        columns=lambda x: x.replace("Finish Date", "End Date") if "Finish Date" in x else x,
        inplace=True,
    )
    df_box.rename(
        columns=lambda x: x.replace("Dx", "DX") if "Dx" in x else x, inplace=True
    )
    df_box.rename(
        columns=lambda x: x.replace("Id", "ID") if "Anonymised" in x else x,
        inplace=True,
    )
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

    if dataset_type == "anonymised prelim":
        df_box = df_box.loc[:, :"Stage At DX"]

    # Remove duplicated columns
    df_box = df_box.loc[:, ~df_box.columns.duplicated()]

    # Round down TTE columns
    tte_cols = [col for col in df_box.columns if "tte" in col.lower()]
    for col in tte_cols:
        s = pd.to_numeric(df_box[col], errors="coerce")
        df_box[col] = np.floor(s * 100) / 100

    # Write to Excel if upload requested
    if not skip_box_upload:
        if testing_mode:
            paths = [
                f"D:/AmmashA.IA/database_management/data_files/leeds_upload_testing/{source_file_name}"
            ]
        elif dataset_type == "anonymised prelim":
            paths = [
                f"C:/Users/AmmashA.IA/Box/Preliminary data files (Leeds)/Updated prelim data sheets/{source_file_name}",
                f"D:/AmmashA.IA/OneDrive - NHS/TrisStar metadata/Box uploads/Updated template/Prelim/{source_file_name}",
            ]
        else:
            paths = [
                f"C:/Users/AmmashA.IA/Box/Follow up data (Leeds)/Updated follow up data sheets/{source_file_name}",
                f"D:/AmmashA.IA/OneDrive - NHS/TrisStar metadata/Box uploads/Updated template/Follow up/{source_file_name}",
            ]

        for p in paths:
            os.makedirs(os.path.dirname(p), exist_ok=True)

        # Fills
        yellow_fill = PatternFill(start_color="00FFFF00", end_color="00FFFF00", fill_type="solid")
        grey_fill = PatternFill(start_color="00E7E6E6", end_color="00E7E6E6", fill_type="solid")
        blue_fill = PatternFill(start_color="009BC2E6", end_color="009BC2E6", fill_type="solid")
        light_red_fill = PatternFill(start_color="00F8CBAD", end_color="00F8CBAD", fill_type="solid")
        light_orange_fill = PatternFill(start_color="00FFD966", end_color="00FFD966", fill_type="solid")
        dark_orange_fill = PatternFill(start_color="00FFC000", end_color="00FFC000", fill_type="solid")
        dark_grey_fill = PatternFill(start_color="00ACB9CA", end_color="00ACB9CA", fill_type="solid")
        pink_fill = PatternFill(start_color="00FF8AD8", end_color="00FF8AD8", fill_type="solid")
        light_green_fill = PatternFill(start_color="00A9D08E", end_color="00A9D08E", fill_type="solid")

        for path in paths:
            # Open existing workbook or create new if it does not exist
            writer_mode = "a" if os.path.exists(path) else "w"
            writer_kwargs = {"engine": "openpyxl", "mode": writer_mode}
            if writer_mode == "a":
                writer_kwargs["if_sheet_exists"] = "replace"

            with pd.ExcelWriter(path, **writer_kwargs) as writer:
                # Write to the NEW separate worksheet
                df_box.to_excel(writer, sheet_name=target_sheet_name, index=False)
                worksheet = writer.sheets[target_sheet_name]

                max_row, max_col = df_box.shape
                table_range = f"A1:{get_column_letter(max_col)}{max_row + 1}"

                # Unique table name based on dataset type to avoid Excel workbook-level name collisions
                table_display_name = (
                    "Table_Prelim_Anon"
                    if dataset_type == "anonymised prelim"
                    else "Table_FollowUp_Anon"
                    if dataset_type == "anonymised follow up"
                    else "Table_Supplementary_Anon"
                )

                if table_display_name in worksheet.tables:
                    del worksheet.tables[table_display_name]

                tab = Table(displayName=table_display_name, ref=table_range)
                tab.tableStyleInfo = TableStyleInfo(
                    name="TableStyleLight1",
                    showFirstColumn=False,
                    showLastColumn=False,
                    showRowStripes=False,
                    showColumnStripes=False,
                )
                worksheet.add_table(tab)

                # Format Supplementary headers
                if dataset_type == "anonymised supplementary":
                    for col_idx in range(1, max_col + 1):
                        worksheet.cell(row=1, column=col_idx).fill = grey_fill

                # Format Prelim headers
                elif dataset_type == "anonymised prelim":
                    for col_idx in range(1, max_col + 1):
                        worksheet.cell(row=1, column=col_idx).fill = yellow_fill

                # Format Follow Up headers by section
                else:
                    def fill_range(start_name, end_name, fill):
                        start_idx, end_idx = None, None
                        for c in range(1, max_col + 1):
                            val = worksheet.cell(row=1, column=c).value
                            if val == start_name:
                                start_idx = c
                            if val == end_name:
                                end_idx = c - 1
                        if start_idx and end_idx:
                            first, last = min(start_idx, end_idx), max(start_idx, end_idx)
                            for c in range(first, last + 1):
                                worksheet.cell(row=1, column=c).fill = fill

                    fill_range("Donor ID", "Any Neoadjuvant Therapy", yellow_fill)
                    fill_range("Any Neoadjuvant Therapy", "Adjuvant Tx1", grey_fill)
                    fill_range("Adjuvant Tx1", "Sites Recurrence/ Progression 1", blue_fill)
                    fill_range("Sites Recurrence/ Progression 1", "Subsequent Tx1", light_red_fill)
                    fill_range("Subsequent Tx1", "Current Status", light_orange_fill)
                    fill_range("Current Status", "Comments From Clinical Team", blue_fill)
                    fill_range("TTE Neoadjuvant Tx1 Start Date", "TTE Overall Survival (OS) In Months", dark_grey_fill)

                    for col_idx in range(1, max_col + 1):
                        val = worksheet.cell(row=1, column=col_idx).value
                        if val == "Comments From Clinical Team":
                            worksheet.cell(row=1, column=col_idx).fill = dark_orange_fill
                        elif val in ["TTE Overall Survival (OS) In Months", "Event (Death; All Causes)"]:
                            worksheet.cell(row=1, column=col_idx).fill = pink_fill
                        elif val in ["TTE Progression Free Survival (PFS) In Months", "Event (Recurrence Or Cancer Death)"]:
                            worksheet.cell(row=1, column=col_idx).fill = light_green_fill
                        elif val in ["Event (Recurrence/Progression)", "Event (Cancer Death)"]:
                            worksheet.cell(row=1, column=col_idx).fill = grey_fill

                # Set column widths
                for i in range(1, max_col + 1):
                    worksheet.column_dimensions[get_column_letter(i)].width = 25

    return df_box

# -----------------------------------------------------------------------------------------------------------------------------------------------

# df_prelim_test = pd.read_excel("D:/AmmashA.IA/OneDrive - NHS/TrisStar metadata/CRC's/CRC 1/CRC 1 Leeds follow up data V2.xlsx", sheet_name="CRC 1 Follow up", header=2)
# df_followup_test = pd.read_excel("D:/AmmashA.IA/OneDrive - NHS/TrisStar metadata/NSCLC/NSCLC 3/NSCLC 3 (ADC met) Leeds follow up data V3.xlsx", sheet_name="NSCLC ADC Met Follow Up", header=2)
# df_supp_test = pd.read_excel("D:/AmmashA.IA/OneDrive - NHS/TrisStar metadata/NSCLC/NSCLC 3/NSCLC 3 (ADC met) Leeds follow up data V3.xlsx", sheet_name="NSCLC ADC Met Supplementary", header=0)


# prelim_anon_test = anonymiser(
#     source_sheet_name="test_prelim_sheet_name",
#     df_onedrive=df_prelim_test,
#     cohort_name="crc1",
#     source_file_name="test_Prelim_file.xlsx",
#     testing_mode=True,
#     skip_box_upload=False

# )

# followup_anon_test = anonymiser(
#     source_sheet_name="Follow up",
#     df_onedrive=df_followup_test,
#     cohort_name="nsclc3adcmet",
#     source_file_name="test_Follow_Up_file.xlsx",
#     testing_mode=True,
#     skip_box_upload=False
# )

# supplementary_anon_test = anonymiser(
#     source_sheet_name="Supplementary",
#     df_onedrive=df_supp_test,
#     cohort_name="nsclc3adcmet",
#     source_file_name="test_Follow_Up_file.xlsx",
#     testing_mode=True,
#     skip_box_upload=False

# )

