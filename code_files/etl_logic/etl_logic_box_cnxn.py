# Import the required libraries
import os
import sys
import time
import warnings

import pandas as pd
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# Suppress the python warning regarding the downcasting behaviour in the replace function
pd.set_option("future.no_silent_downcasting", True)
warnings.filterwarnings(
    "ignore",
    category=FutureWarning,
    message="Downcasting behavior in `replace` is deprecated",
)

# disable warning errors
warnings.filterwarnings("ignore", category=UserWarning, message=".*")
warnings.filterwarnings("ignore", category=FutureWarning, message=".*")

# -----------------------------------------------------------------------------------------------------------------------------------------------------#

# Define the number of maximum retries for reading the file from the configuration JSON file
# Define the delay between retries
MAX_RETRIES = 5
RETRY_DELAY = 2  # seconds

# define the separator line
separator = "\n" + "=" * 156 + "\n"

# -----------------------------------------------------------------------------------------------------------------------------------------------------#


# Function to run the upload process
def run_upload_process_box_cnxn(cohort_config, cohort_name):
    """
    This function runs the upload process of the box connection.
    """

    # Get the Settings from the configuration file
    box_filepath = cohort_config["box_filepath"]
    onedrivebox_filepath = cohort_config["onedrivebox_filepath"]
    onedrivebox_filename = cohort_config["onedrivebox_filename"]
    onedrivebox_sheet_name = cohort_config["onedrivebox_sheet_name"]
    onedrivebox_header_number = cohort_config["onedrivebox_header_number"]
    dataset_type = cohort_config["dataset_type"]

    # Normalise the OneDrive file path
    onedrivebox_filepath = os.path.normpath(onedrivebox_filepath)

    # Initialise the box dataframe variable to None
    df_box = None

    # Retry the read-in of the file within the maximum retries range
    for attempt in range(MAX_RETRIES):
        # Extract and Process the data
        try:
            # Read in the updated worksheet
            df_box = pd.read_excel(
                str(onedrivebox_filepath),
                sheet_name=onedrivebox_sheet_name,
                header=onedrivebox_header_number,
                dtype=object,
            )
            # if read-in successful, break the loop and exit
            print(f"✅ File read successfully on attempt {attempt + 1}.")
            break
        except Exception as e:
            last_exception = e
            print(f"❌ Attempt {attempt + 1} failed to read the file")
            # if the last retry is reached and fails, exit and indicate that the read-in failed
            if attempt == MAX_RETRIES - 1:
                sys.stderr.write(separator)
                sys.stderr.write(
                    f"❌ CRITICAL ERROR: Max retries reached for file: {onedrivebox_filepath}\n"
                )
                sys.stderr.write(
                    f"🔎 REASON: {type(last_exception).__name__}: {last_exception!s}\n"
                )
                sys.stderr.write("📝 Max retries reached. File read failed")
                sys.stderr.write(separator)
                return False
        time.sleep(RETRY_DELAY)

    try:
        # Define the output path for the Box file - setting it based on the dataset type
        # output_path = f"D:/AmmashA.IA/OneDrive - NHS/Leeds MS SQL Server Database Management/{box_filename}_copy.xlsx"
        output_path = box_filepath
        os.makedirs(os.path.dirname(output_path), exist_ok=True)

        # If the file does not exist, create an empty one
        if not os.path.exists(output_path):
            with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                df_box.to_excel(writer, onedrivebox_sheet_name, index=False)

        # Convert the dataframe to an excel table range format
        with pd.ExcelWriter(
            output_path, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            # Write the dataframe to the specified sheet in the output Excel file
            df_box.to_excel(writer, sheet_name=onedrivebox_sheet_name, index=False)

            # define the workbook and worksheet objects from the writer
            workbook = writer.book
            worksheet = writer.sheets[onedrivebox_sheet_name]

            # delete specific tables which might already exist in the workbook, depending on the dataset type variable
            match dataset_type:
                case "anonymised supplementary":
                    if "Supplementary" in worksheet.tables:
                        del worksheet.tables["Supplementary"]
                case "anonymised prelim":
                    if "FollowUp" in worksheet.tables:
                        del worksheet.tables["FollowUp"]
                case "anonymised follow up":
                    if "FollowUp" in worksheet.tables:
                        del worksheet.tables["FollowUp"]

            # define the range of the table and the column settings
            (max_row, max_col) = df_box.shape
            table_range = f"A1:{get_column_letter(max_col)}{max_row + 1}"

            # define a temporary displayName variable
            match dataset_type:
                case "anonymised supplementary":
                    displayName = "Supplementary"
                case "anonymised follow up":
                    displayName = "FollowUp"
                case "anonymised prelim":
                    displayName = "FollowUp"

            # create the table range and assign a style to it
            tab = Table(displayName=displayName, ref=table_range)
            style = TableStyleInfo(
                name="TableStyleLight1",
                showFirstColumn=False,
                showLastColumn=False,
                showRowStripes=False,
                showColumnStripes=False,
            )
            tab.tableStyleInfo = style
            worksheet.add_table(tab)

            if dataset_type == "anonymised supplementary":
                # create the fill color for the header row when the dataset is supplementary
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

    except Exception as e:
        last_exception = e
        print("❌ Data upload failed.")
        sys.stderr.write(separator)
        sys.stderr.write(
            f"❌ Upload failed for the OneDrive Box Uploads file {onedrivebox_filename}"
        )
        sys.stderr.write(
            f"🔎 REASON: {type(last_exception).__name__}: {last_exception!s}\n"
        )
        sys.stderr.write(separator)
        return False
