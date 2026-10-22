# Import the required libraries
import sys
from pathlib import Path

import pandas as pd

current_dir = Path(__file__).resolve()
project_root = next(
    (p for p in Path(".").resolve().parents if (p / "pyproject.toml").exists()),
    Path(".").resolve(),
)
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

# get the code files direcrtory
code_files = project_root / "code_files"

# add the root path to the system
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

# -------------------------------------------------------------------------------------------------------

# tristar inventory
inventory = pd.read_excel(
    project_root
    / "data_files"
    / "non_leeds_data"
    / "inventory"
    / "combined_inventory.xlsx"
)

# iom source workbook
iom_source_path = (
    project_root / "data_files" / "non_leeds_data" / "metadata" / "iom_inventories.xlsx"
)

# hgugm source workbook
hgugm_source_path = (
    project_root
    / "data_files"
    / "non_leeds_data"
    / "metadata"
    / "hgugm_inventories.xlsx"
)

# -------------------------------------------------------------------------------------------------------
# Work with the Leeds samples

# get a unique list of the first 3 letters from the donor id, sample id, and original id columns
donorid_first3 = set(inventory["Donor_ID"].str[:3])
sampleid_first3 = set(inventory["Sample_ID"].str[:3])
originalid_first3 = set(inventory["Original_ID"].str[:3])

# filter out leeds samples
to_drop = (
    ((inventory["Donor_ID"].str.startswith("LH")) & (inventory["Origin"] == "Leeds"))
    | ((inventory["Sample_ID"].str.startswith("LH")) & (inventory["Origin"] == "Leeds"))
    | (
        (inventory["Original_ID"].str.startswith("LH"))
        & (inventory["Origin"] == "Leeds")
    )
    | inventory["Donor_ID"].str.startswith("LDS")
    | inventory["Sample_ID"].str.startswith("LDS")
    | inventory["Original_ID"].str.startswith("LDS")
    | inventory["Donor_ID"].str.startswith("Leeds")
    | inventory["Sample_ID"].str.startswith("Leeds")
    | inventory["Original_ID"].str.startswith("Leeds")
)

inventory = inventory.loc[~to_drop]

# check for remaining leeds samples based on the origin column
unidentified_leeds_samples = inventory.loc[inventory["Origin"] == "Leeds"]

# move the unidentified samples files to a csv
unidentified_leeds_samples.to_excel(
    project_root
    / "data_files"
    / "non_leeds_data"
    / "inventory"
    / "unidentified_leeds_samples.xlsx",
    sheet_name="unidentified_leeds",
    index=False,
)

# remove the remaining unidentified leeds samples
inventory = inventory.loc[inventory["Origin"] != "Leeds"]


# -------------------------------------------------------------------------------------------------------
# Work with the IOM samples
# check the IOM samples
iom_samples = inventory.loc[inventory["Origin"].str.contains("IOM")]

# load the iom source data workbook
iom_source_all_sheets = pd.read_excel(iom_source_path, sheet_name=None)

# loop through each sheet and print the first few rows
for sheet, data in iom_source_all_sheets.items():
    print(data.iloc[:3, :6])
    print("\n")
    print(
        "----------------------------------------------------------------------------"
    )

"""
The ID columns are:
Sample_ID, Original_ID, Subject_ID, Accession_Number, Current_Label
"""

# Configuration
id_cols_inventory = ["Donor_ID", "Sample_ID", "Original_ID"]
id_cols_source = [
    "Sample_ID",
    "Original_ID",
    "Subject_ID",
    "Accession_Number",
    "Current_Label",
]
regex_pattern = r"[^A-Za-z0-9]"


# Define a function to clean the ID columns
def clean_ids(df, columns):
    """Strip all non-alphanumeric characters from specified columns."""
    for col in columns:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(regex_pattern, "", regex=True)
    return df


# Clean the IOM Samples
iom_samples = clean_ids(
    iom_samples.copy(),
    id_cols_inventory + ["Subject_ID", "Accession_Number", "Current_Label"],
)

# Clean every sheet in the source workbook
for sheet in iom_source_all_sheets:
    iom_source_all_sheets[sheet] = clean_ids(
        iom_source_all_sheets[sheet], id_cols_source
    )

# Create Search Set from Inventory
# We flatten all IDs in the inventory to a single set for fast lookup
inventory_id_pool = set(iom_samples[id_cols_inventory].values.flatten())
inventory_id_pool.discard("nan")
inventory_id_pool.discard("None")

# Find Matches in Source Sheets
all_matches_list = []

for sheet_name, data in iom_source_all_sheets.items():
    # Identify which of our target ID columns actually exist in this sheet
    available_cols = [c for c in id_cols_source if c in data.columns]

    if available_cols:
        # Check if any value in the row's ID columns exists in the inventory pool
        mask = data[available_cols].isin(inventory_id_pool).any(axis=1)
        matches = data[mask].copy()

        if not matches.empty:
            matches["Source_Sheet"] = sheet_name
            all_matches_list.append(matches)

# Combine matches into a single DataFrame
if all_matches_list:
    ids_combined = pd.concat(all_matches_list, ignore_index=True)
    print(f"Total matching rows found in source: {len(ids_combined)}")
else:
    ids_combined = pd.DataFrame()
    print("No matches found.")

# Robust Exclusion Logic
if not ids_combined.empty:
    cols_in_matches = [c for c in id_cols_source if c in ids_combined.columns]
    excluded_ids = set(
        ids_combined[cols_in_matches].values.flatten()
    )  # ids that were matched
    excluded_ids.discard("nan")
    excluded_ids.discard("None")

    # Create a mask: True if ANY inventory ID column contains an excluded ID
    remove_mask = (
        iom_samples[id_cols_inventory].isin(excluded_ids).any(axis=1)
    )  # ids that were matched and to be removed
    matched_indices = iom_samples[remove_mask].index

    # update the iom_samples to keep the unmatched samples only
    unidentified_iom_samples = iom_samples.drop(index=matched_indices)

    # move the unidentified samples to an excel file
    unidentified_iom_samples.to_excel(
        project_root
        / "data_files"
        / "non_leeds_data"
        / "inventory"
        / "unidentified_iom_samples.xlsx"
    )

    # remove the same rows from the main inventory
    inventory = inventory.drop(index=matched_indices)

    print(f"Rows removed: {len(matched_indices)}")
else:
    print("Zero matches found.")

# -------------------------------------------------------------------------------------------------------
# Work with the HGUGM data
# In[]:
# define the hgugm samples dataframe from the inventory
hgugm_samples = inventory.loc[inventory["Origin"].str.contains("HGUGM")]

# get the hgugm source workbook
hgugm_source_all_sheets = pd.read_excel(
    project_root
    / "data_files"
    / "non_leeds_data"
    / "metadata"
    / "hgugm_inventories.xlsx",
    sheet_name=None,
)

# show the first few rows and columns in each source sheet
for sheet, data in hgugm_source_all_sheets.items():
    print(data.iloc[:3, :6])
    print("\n")
    print("------------------------------------------------------------------------")
    print("\n")

"""
The ID columns from the source files are Unique_ID and Sample_ID
"""


# %%
