# Import the required libraries
import pandas as pd
from sqlalchemy.types import BIGINT

from code_files.utilities.config.mappings import set_data_types

# Import the required functions
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# Establish a connection to the SQL Server database
project_root, cnxn, cursor, engine, current_file = setup_cnxn()

def clean_nhs_str(series: pd.Series) -> pd.Series:
    """Standardize NHS numbers to 10-digit string format without scientific notation or .0."""
    cleaned = (
        series.astype('string')
        .str.replace(r'\.0$', '', regex=True)
        .str.replace(r'\D', '', regex=True)
        .str.strip()
    )
    return cleaned.apply(lambda x: x.zfill(10) if pd.notna(x) and len(x) > 0 else pd.NA).astype('string')


def create_combined_supp_unanon():

    # Query all supplementary unanonymised tables
    all_supp_unanon_query = """
    SELECT TABLE_SCHEMA, TABLE_NAME 
    FROM INFORMATION_SCHEMA.TABLES
    WHERE TABLE_NAME LIKE '%Supplementary_Unanon%'
    """
    all_supp_unanon = pd.read_sql(all_supp_unanon_query, engine)

    supp_unanon_dfs_list = []
    
    # Query each supplementary table
    for idx, row in all_supp_unanon.iterrows():
        schema = row['TABLE_SCHEMA']
        table = row['TABLE_NAME']
        
        # Read table using engine to avoid connection warnings
        df_query = f"SELECT * FROM [{schema}].[{table}]"
        df = pd.read_sql(df_query, engine)

        # Standardize NHS_Number per table before concatenation
        if 'NHS_Number' in df.columns:
            df['NHS_Number'] = clean_nhs_str(df['NHS_Number'])
            
        if 'DonorID' in df.columns:
            df['DonorID'] = (
                df['DonorID']
                .astype('string')
                .str.replace(r'\.0$', '', regex=True)
                .str.strip()
                .replace('', pd.NA)
            )

        supp_unanon_dfs_list.append(df)

    if not supp_unanon_dfs_list:
        return

    # Concatenate the supplementary dfs
    supp_unanon_combined = pd.concat(supp_unanon_dfs_list, ignore_index=True)

    # create the data types dictionary
    # data_types=set_data_types(supp_unanon_combined, "unanonymised supplementary")
    data_types = {
        'NHS_Number': BIGINT
    }

    # Upload the combined dataframe with explicit SQL column types
    supp_unanon_combined.to_sql(
        "supp_unanon_combined",
        engine,
        if_exists="replace",
        schema="imp_documents",
        index=False,
        dtype=data_types
    )