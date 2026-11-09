# Import the required libraries
import pandas as pd

pd.set_option('display.max_rows', None)

# Import the required functions
from code_files.utilities.config.mappings import set_data_types
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# Establish a connection to the SQL Server database
project_root, cnxn, cursor, engine, current_file = setup_cnxn()

def add_indications_master_list():
    # Query datasets
    supp_unanon_combined = pd.read_sql(
        "SELECT * FROM imp_documents.supp_unanon_combined", cnxn
    )
    master_list_anon = pd.read_sql(
        "SELECT * FROM imp_documents.master_list_long", cnxn
    )
    master_list_long_unanon = pd.read_sql(
        "SELECT * FROM imp_documents.master_list_long_unanon", cnxn
    )

    # Work on copies to avoid modifying inputs in place
    master_dfs = {
        "anon": master_list_anon.copy(),
        "unanon": master_list_long_unanon.copy()
    }
    supp_df = supp_unanon_combined.copy()

    # Lookups
    nhs_lookup = (
        supp_df[supp_df['NHS_Number'].notna() & supp_df['Indication'].notna()]
        [['NHS_Number', 'CohortName', 'Indication']]
        .drop_duplicates(subset=['NHS_Number', 'CohortName'])
        .rename(columns={'Indication': 'Indication_nhs'})
    )

    donor_lookup = (
        supp_df[supp_df['DonorID'].notna() & supp_df['Indication'].notna()]
        [['DonorID', 'CohortName', 'Indication']]
        .drop_duplicates(subset=['DonorID', 'CohortName'])
        .rename(columns={'Indication': 'Indication_donor'})
    )

    # Successive LEFT merges
    for key, df in master_dfs.items():
        if 'DonorID' in df.columns:
            df = df.merge(
                nhs_lookup,
                on=['NHS_Number', 'CohortName'],
                how='left'
            ).merge(
                donor_lookup,
                on=['DonorID', 'CohortName'],
                how='left'
            )
            df['Indication'] = df['Indication_nhs'].combine_first(df['Indication_donor'])
            df = df.drop(columns=['Indication_nhs', 'Indication_donor'])
        else:
            df = df.merge(
                nhs_lookup,
                on=['NHS_Number', 'CohortName'],
                how='left'
            )
            df = df.rename(columns={'Indication_nhs': 'Indication'})

        # Assign back into the dictionary
        master_dfs[key] = df

    master_df = master_dfs["anon"]
    master_df_unanon = master_dfs["unanon"]

    # create the data types dictionary
    data_types_master_df = set_data_types(master_df, "anonymised follow up")
    data_types_master_df_unanon = set_data_types(master_df_unanon, "unanonymised follow up")

    # Upload to the database
    master_df.to_sql(
        name="master_list_long",
        con=engine,
        if_exists="replace",
        schema="imp_documents",
        index=False,
        dtype=data_types_master_df
    )

    master_df_unanon.to_sql(
        name="master_list_long_unanon",
        con=engine,
        if_exists="replace",
        schema="imp_documents",
        index=False,
        dtype=data_types_master_df_unanon
    )

    return master_df, master_df_unanon