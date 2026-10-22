# Import the required libraries
import json

import pandas as pd
from sqlalchemy.exc import SQLAlchemyError

# Import the required functions
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

# Define the path to the file upload settings JSON file
file_upload_settings_path = project_root / "json_files" / "file_upload_settings.json"

# Load the file upload settings file
with open(file_upload_settings_path, "r") as f:
    file_upload_settings = json.load(f)


def update_indications():
    """
    This function collects all the indications information from the unanonymised supplementary tables that have them
    and then transfers that information to the associated anonymised supplementary tables
    """
    # get all the unanonymised supplementary tables
    query = """
    SELECT table_schema, table_name
    FROM information_schema.tables
    WHERE table_name LIKE '%Supplementary_Unanon'
    """
    all_supplementary_tables = pd.read_sql(query, cnxn)

    # Loop through the schema and table names of the supplementary unanon tables
    for schema, tbl in zip(
        all_supplementary_tables["table_schema"], all_supplementary_tables["table_name"]
    ):
        # query each supplementary unanon table
        query_tbl_supp_unanon = f"""
        SELECT * FROM [{schema}].[{tbl}]
        """
        # read in each table as a dataframe
        df_unanon = pd.read_sql(query_tbl_supp_unanon, cnxn)
        # If the dataframe has the indication column in it, get the associated anonymised supplementary file
        if "Indication" in df_unanon.columns:
            # replace the Unanon with Anon in order to get the name of the corresponding anonymised supplementary dataset
            tbl_anon = tbl.replace("Unanon", "Anon")
            try:
                # query the anonymised supplementary dataset
                query_tbl_supp_anon = f"""
                SELECT * FROM [{schema}].[{tbl_anon}]
                """
                # read in each table as a dataframe
                df_anon = pd.read_sql(query_tbl_supp_anon, cnxn)
                # If the indication column is not already in the anonymised data,
                # merge the anonymised supplementary dataset with the Donor ID and Indication column from the unanonymised
                # dataset in order to add the Indication column
                if (
                    "Indication" not in df_anon.columns
                    and "DonorID" in df_unanon.columns
                    and df_unanon["DonorID"].notna().all()
                    and (df_unanon["DonorID"] != "N/A").all()
                ):
                    # if there are N/A or null Donor IDs, filter them out
                    df_unanon = df_unanon[
                        df_unanon["DonorID"].notnull()
                        & (df_unanon["DonorID"] != "N/A")
                    ]
                    df_anon = pd.merge(
                        df_anon,
                        df_unanon[["DonorID", "Indication"]],
                        how="inner",
                        on="DonorID",
                    )
                    # Upload the dataframe back to the database
                    df_anon.to_sql(
                        name=tbl_anon,
                        con=engine,
                        schema=schema,
                        if_exists="replace",
                    )
            except (
                SQLAlchemyError,
                TypeError,
                ValueError,
                KeyError,
                AttributeError,
            ) as exc:
                print(f"Skipped table {tbl_anon}: {exc}")
