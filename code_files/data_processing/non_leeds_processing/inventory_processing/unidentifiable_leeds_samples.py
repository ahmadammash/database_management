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
