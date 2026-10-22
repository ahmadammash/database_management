# Import the required libraries
import re
import warnings

import pandas as pd
from sqlalchemy import text

# Disable UserWarnings and FutureWarnings
warnings.filterwarnings("ignore", category=UserWarning)
warnings.filterwarnings("ignore", category=FutureWarning)

# Import the required functions from the files
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

# ----------------------------------------------------------------------------------------------------

# query the tristar inventory
fetch_inventory = """
SELECT * FROM "inventory"."tristar_inventory"
"""

# fetch the whole inventory
with engine.connect() as cnxn:
    inventory = pd.read_sql(text(fetch_inventory), cnxn)

inventory = inventory.drop_duplicates(subset=["Sample_ID", "Original_ID", "Donor_ID"])


# filter for the iom samples
with engine.connect() as cnxn:
    iom_inventory = pd.read_sql(text(fetch_inventory), cnxn)
iom_inventory = iom_inventory.loc[iom_inventory.Origin.str.contains("IOM")]
iom_inventory = iom_inventory.drop_duplicates(
    subset=["Sample_ID", "Original_ID", "Donor_ID"]
)

# filter for the chu nice samples
with engine.connect() as cnxn:
    chu_nice_inventory = pd.read_sql(text(fetch_inventory), cnxn)
chu_nice_inventory = chu_nice_inventory.loc[
    chu_nice_inventory.Origin.str.contains("Nice")
]
chu_nice_inventory = chu_nice_inventory.drop_duplicates(
    subset=["Sample_ID", "Original_ID", "Donor_ID"]
)

# filter for samples not from iom and chu nice
with engine.connect() as cnxn:
    other_samples = pd.read_sql(text(fetch_inventory), cnxn)
other_samples = other_samples.loc[
    ~(other_samples.Origin.str.contains("Nice"))
    & ~(other_samples.Origin.str.contains("IOM"))
    & ~(other_samples.Origin.str.contains("Leeds"))
]
other_samples = other_samples.drop_duplicates(
    subset=["Sample_ID", "Original_ID", "Donor_ID"]
)

# ----------------------------------------------------------------------------------------------------

# transform the inventory dfs to a long format and keep the IDs only
chu_nice_inventory_long = pd.melt(
    chu_nice_inventory,
    value_vars=["Donor_ID", "Sample_ID", "Original_ID"],
    var_name="ID_Type",
    value_name="ID",
)

iom_inventory_long = pd.melt(
    iom_inventory,
    value_vars=["Donor_ID", "Sample_ID", "Original_ID"],
    var_name="ID_Type",
    value_name="ID",
)

other_samples_long = pd.melt(
    other_samples,
    value_vars=["Donor_ID", "Sample_ID", "Original_ID"],
    var_name="ID_Type",
    value_name="ID",
)
# In[]:

# drop duplicates from the long format
chu_nice_inventory_long = chu_nice_inventory_long.drop_duplicates()
iom_inventory_long = iom_inventory_long.drop_duplicates()
other_samples_long = other_samples_long.drop_duplicates()

# process the ids in the chu_nice_inventory_long
chu_nice_inventory_long["ID"] = chu_nice_inventory_long["ID"].apply(
    lambda x: re.sub(r"[^A-Za-z0-9]", "", str(x))
)

# In[]:

# process the ids in the iom_inventory_long
iom_inventory_long["ID"] = iom_inventory_long["ID"].apply(
    lambda x: re.sub(r"[^A-Za-z0-9]", "", str(x))
)

# In[]:

# process the ids in the other_samples_long
other_samples_long["ID"] = other_samples_long["ID"].apply(
    lambda x: re.sub(r"[^A-Za-z0-9]", "", str(x))
)

# %%
