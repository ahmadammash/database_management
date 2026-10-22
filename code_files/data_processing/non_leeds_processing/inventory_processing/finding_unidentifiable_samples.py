# Import the required libraries
import openpyxl
import pandas as pd

from code_files.data_processing.utilities_processing.functions import clean_ids

# Import functions from the project directory
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

# get the data_processing file directory
data_processing_file = project_root / "code_files" / "data_processing"

# -------------------------------------------------------------------------------------------------------


# define the function to find unidentifiable samples
def findUnidentifiableSamples(source_file_path, id_columns_source: list[str], origin):
    # 1. Import the master inventory
    inventory = pd.read_excel(
        project_root
        / "data_files"
        / "non_leeds_data"
        / "inventory"
        / "combined_inventory.xlsx"
    )

    if origin == "Hannover":
        inventory = inventory.loc[
            inventory["Origin"].str.contains(origin)
            | inventory["Origin"].str.contains("Atreide")
        ]

    else:
        inventory = inventory.loc[inventory["Origin"].str.contains(origin)]

    # 2. Standardize source data format (Single vs Multi-sheet)
    wb = openpyxl.load_workbook(filename=source_file_path, read_only=True)
    nb_sheets = len(wb.sheetnames)
    wb.close()

    if nb_sheets == 1:
        single_df = pd.read_excel(source_file_path)
        source_sheets_dict = {"Sheet1": single_df}
    else:
        source_sheets_dict = pd.read_excel(source_file_path, sheet_name=None)

    # 3. Import the link file
    link_file_source_path = (
        project_root / "data_files" / "non_leeds_data" / "inventory" / "link_file.xlsx"
    )
    link_file = pd.read_excel(link_file_source_path)

    id_cols_inventory = ["Donor_ID", "Sample_ID", "Original_ID"]
    regex_pattern = r"[^A-Za-z0-9]"

    # 4. Clean IDs across all source sheets
    cols_to_clean = id_cols_inventory + [
        id for id in id_columns_source if id not in id_cols_inventory
    ]
    for sheet_name in source_sheets_dict:
        source_sheets_dict[sheet_name] = clean_ids(
            source_sheets_dict[sheet_name].copy(),
            cols_to_clean,
            regex_pattern=regex_pattern,
        )
        print(source_sheets_dict[sheet_name].head())

    # 5. Bring in longer Original_IDs from the link file
    inventory = inventory.merge(
        link_file[["Original_ID", "Sample_ID"]],
        on="Sample_ID",
        how="left",
        suffixes=("", "_link_file"),
    )
    inventory["Original_ID"] = inventory["Original_ID"].fillna(
        inventory["Original_ID_link_file"]
    )
    inventory = inventory.drop(columns=["Original_ID_link_file"])

    inventory = clean_ids(inventory, id_cols_inventory, regex_pattern=regex_pattern)

    # 6. Extract all unique cleaned/processed IDs from source sheets to build our search components
    cleaned_inventory_cols = [f"{col}_cleaned" for col in id_cols_inventory]
    inventory_id_pool = set(inventory[cleaned_inventory_cols].values.flatten())
    inventory_id_pool.discard("nan")
    inventory_id_pool.discard("None")

    source_id_pool = set()
    for data in source_sheets_dict.values():
        cleaned_source_cols = [
            f"{c}_cleaned" for c in id_columns_source if f"{c}_cleaned" in data.columns
        ]

        if cleaned_source_cols:
            flattened = data[cleaned_source_cols].values.flatten()
            source_id_pool.update(
                [
                    str(x).strip()
                    for x in flattened
                    if pd.notna(x) and str(x).strip() not in ("nan", "None", "")
                ]
            )

    # 7. Core Match Direction: Check if a source ID is a substring of an inventory ID
    def check_inventory_row(row):
        for inv_col in cleaned_inventory_cols:
            inv_val = str(row[inv_col])
            if pd.isna(row[inv_col]) or inv_val.strip() in ("nan", "None", ""):
                continue

            for src_id in source_id_pool:
                if src_id in inv_val:
                    return True
        return False

    print("Evaluating substring matches (Source ID inside Inventory ID)...")

    # Mark which rows in the inventory found a match inside the source documents
    inventory_matched_mask = inventory.apply(check_inventory_row, axis=1)

    # 8. Separate inventory into Identified and Unidentified chunks
    matched_inventory_samples = inventory[inventory_matched_mask]
    unidentified_inventory_samples = inventory[~inventory_matched_mask]

    print(f"Total matching inventory rows processed: {len(matched_inventory_samples)}")
    print(
        f"Total unidentified inventory rows remaining: {len(unidentified_inventory_samples)}"
    )

    # 9. Save out the unidentified samples using the INVENTORY format layout
    unidentified_output_path = (
        project_root
        / "data_files"
        / "non_leeds_data"
        / "inventory"
        / "unidentified_samples.xlsx"
    )

    unidentified_inventory_samples = unidentified_inventory_samples.drop_duplicates()

    inventory = unidentified_inventory_samples.copy()

    with pd.ExcelWriter(
        unidentified_output_path, mode="a", engine="openpyxl", if_sheet_exists="replace"
    ) as writer:
        unidentified_inventory_samples.to_excel(writer, sheet_name=origin, index=False)
    print(f"Unidentified inventory register written to: {unidentified_output_path}")

    # Return the clean, remaining inventory data alongside the unmatched inventory subset
    return inventory, unidentified_inventory_samples


# In[]:

paths = [
    project_root
    / "data_files"
    / "non_leeds_data"
    / "metadata"
    / "arteide_hannover_samples.xlsx",
    project_root
    / "data_files"
    / "non_leeds_data"
    / "metadata"
    / "chu_nice_inventory.xlsx",
    project_root
    / "data_files"
    / "non_leeds_data"
    / "metadata"
    / "hgugm_inventories.xlsx",
    project_root
    / "data_files"
    / "non_leeds_data"
    / "metadata"
    / "iom_inventories.xlsx",
]

ids = [
    ["Patient_ID", "Sample_ID"],
    ["Sample_ID", "Biobank_ID", "Pathology_ID"],
    ["Unique_ID", "Sample_ID"],
    ["Sample_ID", "Accession_Number", "Current_Label"],
]

origins = ["Hannover", "Nice", "HGUGM", "IOM"]

for path, id, origin in zip(paths, ids, origins):
    findUnidentifiableSamples(path, id, origin)


# %%
