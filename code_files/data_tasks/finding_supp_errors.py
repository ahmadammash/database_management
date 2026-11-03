# ================================================================================================================
# IMPORTS AND SYSTEMS SETUP
# ================================================================================================================

# Import the required libraries
import warnings
from tkinter import *

import pandas as pd

# Suppress warnings
warnings.simplefilter("ignore", UserWarning)
warnings.simplefilter("ignore", FutureWarning)

# Import functions from the project directory
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# ================================================================================================================
# DATABASE CONNECTION & INITIALISATION
# ================================================================================================================

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()


query_anon_supp_tbls = """SELECT TABLE_NAME, TABLE_SCHEMA FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME LIKE '%Supplementary_Anon%'"""
all_anon_supp_tbls = pd.read_sql(query_anon_supp_tbls, cnxn)
query_unanon_supp_tbls = """SELECT TABLE_NAME, TABLE_SCHEMA FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME LIKE '%Supplementary_UnAnon%'"""
all_unanon_supp_tbls = pd.read_sql(query_unanon_supp_tbls, cnxn)




    


