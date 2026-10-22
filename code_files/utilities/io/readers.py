import pandas as pd


def find_headerNumber(filename_main, selected_sheet, dataset_type):

    # give a default value to the header number
    header_number = 0

    # read in a preview of the dataframe in order to determine the header row
    df_preview = pd.read_excel(
        filename_main, sheet_name=selected_sheet, header=None, nrows=100
    )

    # iterate through the rows to identify the header row based on the dataset type
    for idx, row in df_preview.iterrows():
        first_cell = str(row[0]).strip() if pd.notna(row[0]) else ""

        if dataset_type in ["unanonymised follow up", "unanonymised prelim"] and (
            first_cell.lower() == "nhs number"
            or first_cell.lower() == "nhs"
            or "qc" in first_cell.lower()
        ) or (
            dataset_type in ["anonymised follow up", "anonymised prelim"]
            and first_cell.lower() == "donor id"
        ) or dataset_type in [
            "anonymised supplementary",
            "unanonymised supplementary",
        ] and first_cell.lower() in ["nhs number", "donor id"] or dataset_type == "scan logs" and first_cell.lower() == "lh number":
            header_number = idx
    return header_number


def show_worksheet_names(filename_main):
    xl = pd.ExcelFile(filename_main)
    sheet_names_main = xl.sheet_names
    print("Below are the worksheets in the imported workbook:")
    for i, sheet in enumerate(sheet_names_main):
        print(f"{i + 1}: {sheet}")
    return sheet_names_main
