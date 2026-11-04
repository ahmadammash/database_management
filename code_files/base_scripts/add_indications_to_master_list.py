# Import the required libraries
import pandas as pd

# Import the required functions
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# Establish a connection to the SQL Server database
project_root, cnxn, cursor, engine, current_file = setup_cnxn()

def add_indications_master_list():

    # query the combined unanonymised supplementary data
    supp_unanon_combined = pd.read_sql(
        """
        SELECT * FROM imp_documents.supp_unanon_combined
        """,
        cnxn
    )

    # # query the combined anonymised supplementary data
    # supp_anon_combined = pd.read_sql(
    #     """
    #     SELECT * FROM imp_documents.supp_anon_combined
    #     """,
    #     cnxn
    # )

    # query the anonymised master list
    master_list_anon = pd.read_sql(
        """
        SELECT * FROM imp_documents.master_list_long
        """,
        cnxn
    )

    # query the unanonymised master list
    # master_list_unanon = pd.read_sql(
    #     """
    #     SELECT * FROM imp_documents.master_list_long_unanon
    #     """,
    #     cnxn
    # )

    # drop the indication column from the master list anon dataset
    if 'Indication' in master_list_anon.columns:
        master_list_anon = master_list_anon.drop(columns='Indication')

    # Merge the anonymised master list with the combined unanonymised supplementary dataset
    # Split supplementary data by presence of NHS_Number
    supp_by_nhs = supp_unanon_combined[supp_unanon_combined['NHS_Number'].notna()]
    supp_by_donor = supp_unanon_combined[supp_unanon_combined['NHS_Number'].isna()]

    # Merge each partition with the master list
    # If matching on NHS_Number, drop DonorID from supplementary to prevent column collision (_x, _y)
    supp_by_nhs_clean = supp_by_nhs.drop(columns=['DonorID'], errors='ignore')
    merged_nhs = pd.merge(
        master_list_anon,
        supp_by_nhs_clean[['NHS_Number', 'Indication', 'CohortName']],
        on=['NHS_Number', 'CohortName'],
        how='inner'
    )

    # If matching on DonorID, drop NHS_Number from supplementary
    supp_by_donor_clean = supp_by_donor.drop(columns=['NHS_Number'], errors='ignore')
    merged_donor = pd.merge(
        master_list_anon,
        supp_by_donor_clean[['DonorID', 'Indication', 'CohortName']],
        on=['DonorID', 'CohortName'],
        how='inner'
    )

    # Concatenate both results back together
    master_supp_merged = pd.concat([merged_nhs, merged_donor], ignore_index=True)

    # Upload the master list back to the database
    master_supp_merged.to_sql(
        name="master_list_long",
        con=engine,
        if_exists="replace",
        schema="imp_documents"
    )