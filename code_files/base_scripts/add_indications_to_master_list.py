# Import the required libraries
import pandas as pd

pd.set_option('display.max_rows', None)

# Import the required functions
from code_files.utilities.config.mappings import set_data_types
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# Establish a connection to the SQL Server database
project_root, cnxn, cursor, engine, current_file = setup_cnxn()

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

def add_indications_master_list(supp_unanon_combined, master_list_anon, master_list_long_unanon):
    # Work on copies to avoid modifying inputs in place
    master_df = master_list_anon.copy()
    master_df_unanon = master_list_long_unanon.copy()
    supp_df = supp_unanon_combined.copy()

    # Drop existing Indication column if present
    # for df in [master_df, master_df_unanon]:
    #     if 'Indication' in df.columns:
    #         df = df.drop(columns='Indication')

    # Lookup A: NHS Number + CohortName
    # Only keep rows where NHS_Number is valid and Indication is populated
    nhs_lookup = (
        supp_df[supp_df['NHS_Number'].notna() & supp_df['Indication'].notna()]
        [['NHS_Number', 'CohortName', 'Indication']]
        .drop_duplicates(subset=['NHS_Number', 'CohortName'])
        .rename(columns={'Indication': 'Indication_nhs'})
    )

    # Lookup B: DonorID + CohortName
    # Only keep rows where DonorID is valid and Indication is populated
    donor_lookup = (
        supp_df[supp_df['DonorID'].notna() & supp_df['Indication'].notna()]
        [['DonorID', 'CohortName', 'Indication']]
        .drop_duplicates(subset=['DonorID', 'CohortName'])
        .rename(columns={'Indication': 'Indication_donor'})
    )

    # Successive LEFT merges
    for df in [master_df, master_df_unanon]:
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

            # Waterfall resolution:
            # If DonorID was NULL on master_list, Indication_donor is NaN.
            # combine_first automatically resolves to Indication_nhs.            
            df['Indication'] = df['Indication_nhs'].combine_first(
            df['Indication_donor'])
            # Clean up auxiliary columns
            df = df.drop(columns=['Indication_nhs', 'Indication_donor'])

        else:
            df = df.merge(
                nhs_lookup,
                on=['NHS_Number', 'CohortName'],
                how='left'
            )            
            # Waterfall resolution:
            # If DonorID was NULL on master_list, Indication_donor is NaN.
            # combine_first automatically resolves to Indication_nhs.
            df = df.rename(columns={'Indication_nhs': 'Indication'})



    # create the data types dictionarya
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



