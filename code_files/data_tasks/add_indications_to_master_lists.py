# Import the required libraries
import pandas as pd

# Import the required functions
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# Establish a connection to the SQL Server database
project_root, cnxn, cursor, engine, current_file = setup_cnxn()

# query the combined unanonymised supplementary data
supp_unanon_combined = pd.read_sql(
    """
    SELECT * FROM imp_documents.supp_unanon_combined
    """,
    cnxn
)

# query the combined anonymised supplementary data
supp_anon_combined = pd.read_sql(
    """
    SELECT * FROM imp_documents.supp_anon_combined
    """,
    cnxn
)

# query the anonymised master list
master_list_anon = pd.read_sql(
    """
    SELECT * FROM imp_documents.master_list_long
    """,
    cnxn
)

# query the unanonymised master list
master_list_unanon = pd.read_sql(
    """
    SELECT * FROM imp_documents.master_list_long_unanon
    """,
    cnxn
)

# Merge the anonymised master list with the combined unanonymised supplementary Indication column
# using the NHS number and the Cohort Name
master_list_anon = pd.merge(
                            master_list_anon, 
                            supp_unanon_combined[['CohortName', 'NHS_Number', 'Indication']],
                            on=['NHS_Number', 'CohortName'],
                            how='left')

# Upload the updated master list to the database
master_list_anon.to_sql(name = "master_list_long", con=engine, if_exists='replace', schema='imp_documents', index=False)

# Merge the unanonymised master list with the combined unanonymised supplementary Indication column
# Using the NHS number and the Cohort Name
# master_list_unanon = pd.merge(master_list_unanon,
#                               supp_unanon_combined[['CohortName', 'NHS_Number', 'Indication']],
#                               on=['NHS_Number', 'CohortName'],
#                               how='left'
#                               )

# # Upload the updated master list to the database
# master_list_unanon.to_sql(name = "master_list_long_unanon", con=engine, if_exists='replace', schema='imp_documents', index=False)