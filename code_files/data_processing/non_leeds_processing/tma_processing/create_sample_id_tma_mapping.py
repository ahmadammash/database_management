######################################################################################################
# General Setup and Imports
######################################################################################################
import pandas as pd
from sqlalchemy import text

# Import the required functions
from code_files.utilities.io.database import create_gcloud_engine
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

# Establish the connection to GCP
env_path = project_root / "code_files" / "utilities" / "io" / ".env"
engine = create_gcloud_engine(env_path)

######################################################################################################
# Collect the TMA table names from GCP and Create the Dictionary
######################################################################################################

query = text("""
SELECT table_name 
FROM information_schema.tables
WHERE table_schema = 'tma'
  AND table_type = 'BASE TABLE'
  AND table_name != 'sample_id_tma_num_mapping'  -- Exclude target mapping table
ORDER BY table_name;
""")

id_columns = ["Sample_ID", "Sample_ID_2", "DonorID"]
mappings = []  # Store tuples of (Sample_ID, TMA_Num)

with engine.connect() as cnxn:
    result = cnxn.execute(query)
    tables = [row[0] for row in result]

    for tbl in tables:
        # Check column names without fetching full table data
        peek_query = text(f'SELECT * FROM tma."{tbl}" LIMIT 1;')
        df_preview = pd.read_sql(peek_query, cnxn)

        for col in id_columns:
            if col in df_preview.columns:
                # Fetch distinct non-null IDs for this column
                id_query = text(
                    f'SELECT DISTINCT "{col}" FROM tma."{tbl}" WHERE "{col}" IS NOT NULL;'
                )
                ids_df = pd.read_sql(id_query, cnxn)

                # Append each ID and its corresponding table to our master list
                for sample_id in ids_df[col].tolist():
                    mappings.append((sample_id, tbl))

######################################################################################################
# Upload the mapping to GCP
######################################################################################################
# 1. Create DataFrame from all accumulated pairs
sample_id_tma_num_df = pd.DataFrame(mappings, columns=["Sample_ID", "TMA_Num"])

# 2. Remove duplicate pairs (e.g. if Sample_ID and Sample_ID_2 matched in the same table)
sample_id_tma_num_df = sample_id_tma_num_df.drop_duplicates()

# 3. Explicit type conversion (resolves np.int64 serialization errors)
sample_id_tma_num_df["Sample_ID"] = sample_id_tma_num_df["Sample_ID"].astype(str)
sample_id_tma_num_df["TMA_Num"] = sample_id_tma_num_df["TMA_Num"].astype(str)

# 4. Upload clean DataFrame
sample_id_tma_num_df.to_sql(
    "sample_id_tma_num_mapping",
    con=engine,
    if_exists="replace",
    schema="tma",
    index=False,
)
