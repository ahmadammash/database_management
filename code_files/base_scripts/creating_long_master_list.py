# In[]: GENERAL IMPORTS
# Import the required libraries
import re
import warnings

import numpy as np
import pandas as pd
from sqlalchemy.exc import SQLAlchemyError

# Disable UserWarnings and FutureWarnings
warnings.filterwarnings("ignore", category=UserWarning)
warnings.filterwarnings("ignore", category=FutureWarning)

# Import functions from the project directory
from code_files.querying.fetch_information import fetch_schema_names, query_patientList
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

schema_names = fetch_schema_names()  # get the schema names from the inventory
schemas_list = list(schema_names["schema_name"])  # convert the schema_names to a list

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()


# Define a function to create the long master list
def create_long_master_list():
    fetch_all_tables = """
    SELECT table_schema, table_name
    FROM information_schema.tables
    WHERE table_schema NOT IN ('information_schema', 'pg_catalog', 'guest')
                    AND table_schema NOT LIKE '%sys%'
                    AND (table_name LIKE '%_FollowUp_Anon%' or table_name LIKE '%_Prelim_Anon%');
    """
    # Run the query to fetch all user_defined tables
    all_tables = pd.read_sql(fetch_all_tables, cnxn)

    # Define a list to hold the dataframes generated from the queries
    dfs = []

    print(f"Starting to pull data from {len(all_tables)} tables...")

    # Define the target columns to be selected from each user-defined anonymised follow up data table
    target_columns = [
        "DonorID",
        "CohortName",
        "AnyImmunotherapy",
        "SpecSurgery",
        "PrimSurgeryDate",
        "PrimTumourSite",
        "SelSpecAnonymisedID",
        "AnonymisedID_MatchedMet1",
        "AnonymisedID_MatchedMet2",
        "AnonymisedID_MatchedMet3",
        "AnonymisedID_MatchedMet4",
        "AnonymisedID_MatchedMet5",
        "AnonymisedID_MatchedMet6",
        "AnonymisedID_MatchedMet7",
        "AnonymisedID_NormalTissue",
        "SpecTreatmentSubgrouping",
        "MatchedMet1_TreatmentSubgrouping",
        "MatchedMet2_TreatmentSubgrouping",
        "MatchedMet3_TreatmentSubgrouping",
        "MatchedMet4_TreatmentSubgrouping",
        "MatchedMet5_TreatmentSubgrouping",
        "MatchedMet6_TreatmentSubgrouping",
        "MatchedMet7_TreatmentSubgrouping",
        "MatchedNormal_TreatmentSubgrouping",
        "SpecTumourSite",
        "MatchedMet1_Site",
        "MatchedMet2_Site",
        "MatchedMet3_Site",
        "MatchedMet4_Site",
        "MatchedMet5_Site",
        "MatchedMet6_Site",
        "MatchedMet7_Site",
        "MatchedNormalTissue_Site",
        "SpecSurgeryDate",
        "MatchedMet1_SurgeryDate",
        "MatchedMet2_SurgeryDate",
        "MatchedMet3_SurgeryDate",
        "MatchedMet4_SurgeryDate",
        "MatchedMet5_SurgeryDate",
        "MatchedMet6_SurgeryDate",
        "MatchedMet7_SurgeryDate",
        "MatchedNormalTissue_SurgeryDate",
        "SpecBlockNum1",
        "SpecBlockNum2",
        "MatchedMet1_BlockNumber",
        "MatchedMet2_BlockNumber",
        "MatchedMet3_BlockNumber",
        "MatchedMet4_BlockNumber",
        "MatchedMet5_BlockNumber",
        "MatchedMet6_BlockNumber",
        "MatchedMet7_BlockNumber",
        "MatchedNormalTissue_BlockNumber",
        "SpecIsRecurrence",
        "Gender",
        "AgeDX",
        "SpecHistologicalDiagnosis",
        "ResectionMarginsClear",
        "StagingGradingInfoProvidedFor",
        "pT",
        "pN",
        "pM",
        "cM",
        "TNM_Stage",
        "Grade",
        "StagingCriteria",
        "StageAtDX",
        "DateDX",
        "AnyNeoadjuvantTherapy",
        "ResponseToNeoTherapy_TRScore",
        "NeoadjuvantTx1",
        "NeoadjuvantTx1_Agent",
        "TumourResponse_NeoadjuvantTx1",
        "NeoadjuvantTx2",
        "NeoadjuvantTx2_Agent",
        "TumourResponse_NeoadjuvantTx2",
        "NeoadjuvantTx3",
        "NeoadjuvantTx3_Agent",
        "TumourResponse_NeoadjuvantTx3",
        "NeoadjuvantTx4",
        "NeoadjuvantTx4_Agent",
        "TumourResponse_NeoadjuvantTx4",
        "NeoadjuvantTx5",
        "NeoadjuvantTx5_Agent",
        "TumourResponse_NeoadjuvantTx5",
        "NeoadjuvantTx6",
        "NeoadjuvantTx6_Agent",
        "TumourResponse_NeoadjuvantTx6",
        "NeoadjuvantTx7",
        "NeoadjuvantTx7_Agent",
        "TumourResponse_NeoadjuvantTx7",
        "NeoadjuvantTx8",
        "NeoadjuvantTx8_Agent",
        "TumourResponse_NeoadjuvantTx8",
        "NeoadjuvantTx9",
        "NeoadjuvantTx9_Agent",
        "TumourResponse_NeoadjuvantTx9",
        "NeoadjuvantTx10",
        "NeoadjuvantTx10_Agent",
        "TumourResponse_NeoadjuvantTx10",
        "AdjuvantTx1",
        "AdjuvantTx1_Agent",
        "TumourResponse_AdjuvantTx1",
        "AdjuvantTx2",
        "AdjuvantTx2_Agent",
        "TumourResponse_AdjuvantTx2",
        "AdjuvantTx3",
        "AdjuvantTx3_Agent",
        "TumourResponse_AdjuvantTx3",
        "AdjuvantTx4",
        "AdjuvantTx4_Agent",
        "TumourResponse_AdjuvantTx4",
        "AdjuvantTx5",
        "AdjuvantTx5_Agent",
        "TumourResponse_AdjuvantTx5",
        "AdjuvantTx6",
        "AdjuvantTx6_Agent",
        "TumourResponse_AdjuvantTx6",
        "AdjuvantTx7",
        "AdjuvantTx7_Agent",
        "TumourResponse_AdjuvantTx7",
        "AdjuvantTx8",
        "AdjuvantTx8_Agent",
        "TumourResponse_AdjuvantTx8",
        "AdjuvantTx9",
        "AdjuvantTx9_Agent",
        "TumourResponse_AdjuvantTx9",
        "AdjuvantTx10",
        "AdjuvantTx10_Agent",
        "TumourResponse_AdjuvantTx10",
        "AdjuvantTx11",
        "AdjuvantTx11_Agent",
        "TumourResponse_AdjuvantTx11",
        "AdjuvantTx12",
        "AdjuvantTx12_Agent",
        "TumourResponse_AdjuvantTx12",
        "AdjuvantTx13",
        "AdjuvantTx13_Agent",
        "TumourResponse_AdjuvantTx13",
        "AdjuvantTx14",
        "AdjuvantTx14_Agent",
        "TumourResponse_AdjuvantTx14",
        "AdjuvantTx15",
        "AdjuvantTx15_Agent",
        "TumourResponse_AdjuvantTx15",
        "Sites_RecurrenceProgression1",
        "Sites_RecurrenceProgression2",
        "Sites_RecurrenceProgression3",
        "Sites_RecurrenceProgression4",
        "Sites_RecurrenceProgression5",
        "Sites_RecurrenceProgression6",
        "Sites_RecurrenceProgression7",
        "Sites_RecurrenceProgression8",
        "Sites_RecurrenceProgression9",
        "Sites_RecurrenceProgression10",
        "Sites_RecurrenceProgression11",
        "Sites_RecurrenceProgression12",
        "Sites_RecurrenceProgression13",
        "Sites_RecurrenceProgression14",
        "Sites_RecurrenceProgression15",
        "SubsequentTx1",
        "SubsequentTx1_Agent",
        "TumourResponse_SubsequentTx1",
        "SubsequentTx2",
        "SubsequentTx2_Agent",
        "TumourResponse_SubsequentTx2",
        "SubsequentTx3",
        "SubsequentTx3_Agent",
        "TumourResponse_SubsequentTx3",
        "SubsequentTx4",
        "SubsequentTx4_Agent",
        "TumourResponse_SubsequentTx4",
        "SubsequentTx5",
        "SubsequentTx5_Agent",
        "TumourResponse_SubsequentTx5",
        "SubsequentTx6",
        "SubsequentTx6_Agent",
        "TumourResponse_SubsequentTx6",
        "SubsequentTx7",
        "SubsequentTx7_Agent",
        "TumourResponse_SubsequentTx7",
        "SubsequentTx8",
        "SubsequentTx8_Agent",
        "TumourResponse_SubsequentTx8",
        "SubsequentTx9",
        "SubsequentTx9_Agent",
        "TumourResponse_SubsequentTx9",
        "SubsequentTx10",
        "SubsequentTx10_Agent",
        "TumourResponse_SubsequentTx10",
        "SubsequentTx11",
        "SubsequentTx11_Agent",
        "TumourResponse_SubsequentTx11",
        "SubsequentTx12",
        "SubsequentTx12_Agent",
        "TumourResponse_SubsequentTx12",
        "SubsequentTx13",
        "SubsequentTx13_Agent",
        "TumourResponse_SubsequentTx13",
        "SubsequentTx14",
        "SubsequentTx14_Agent",
        "TumourResponse_SubsequentTx14",
        "SubsequentTx15",
        "SubsequentTx15_Agent",
        "TumourResponse_SubsequentTx15",
        "CurrentStatus",
        "DeathTumourRelated",
        "CauseOfDeath",
        "ConcurrentMedicalConditions",
        "Comments_ClinicalTeam",
        "TTE_NeoadjuvantTx1StartDate",
        "TTE_NeoadjuvantTx1EndDate",
        "TTE_NeoadjuvantTx2StartDate",
        "TTE_NeoadjuvantTx2EndDate",
        "TTE_NeoadjuvantTx3StartDate",
        "TTE_NeoadjuvantTx3EndDate",
        "TTE_NeoadjuvantTx4StartDate",
        "TTE_NeoadjuvantTx4EndDate",
        "TTE_NeoadjuvantTx5StartDate",
        "TTE_NeoadjuvantTx5EndDate",
        "TTE_NeoadjuvantTx6StartDate",
        "TTE_NeoadjuvantTx6EndDate",
        "TTE_NeoadjuvantTx7StartDate",
        "TTE_NeoadjuvantTx7EndDate",
        "TTE_NeoadjuvantTx8StartDate",
        "TTE_NeoadjuvantTx8EndDate",
        "TTE_NeoadjuvantTx9StartDate",
        "TTE_NeoadjuvantTx9EndDate",
        "TTE_NeoadjuvantTx10StartDate",
        "TTE_NeoadjuvantTx10EndDate",
        "TTE_SpecimenSurgeryDate",
        "TTE_PrimarySurgeryDate",
        "TTE_MatchedMet1_SurgeryDate",
        "TTE_MatchedMet2_SurgeryDate",
        "TTE_MatchedMet3_SurgeryDate",
        "TTE_MatchedMet4_SurgeryDate",
        "TTE_MatchedMet5_SurgeryDate",
        "TTE_MatchedNormalTissue_SurgeryDate",
        "TTE_AdjuvantTx1StartDate",
        "TTE_AdjuvantTx1EndDate",
        "TTE_AdjuvantTx2StartDate",
        "TTE_AdjuvantTx2EndDate",
        "TTE_AdjuvantTx3StartDate",
        "TTE_AdjuvantTx3EndDate",
        "TTE_AdjuvantTx4StartDate",
        "TTE_AdjuvantTx4EndDate",
        "TTE_AdjuvantTx5StartDate",
        "TTE_AdjuvantTx5EndDate",
        "TTE_AdjuvantTx6StartDate",
        "TTE_AdjuvantTx6EndDate",
        "TTE_AdjuvantTx7StartDate",
        "TTE_AdjuvantTx7EndDate",
        "TTE_AdjuvantTx8StartDate",
        "TTE_AdjuvantTx8EndDate",
        "TTE_AdjuvantTx9StartDate",
        "TTE_AdjuvantTx9EndDate",
        "TTE_AdjuvantTx10StartDate",
        "TTE_AdjuvantTx10EndDate",
        "TTE_AdjuvantTx11StartDate",
        "TTE_AdjuvantTx11EndDate",
        "TTE_AdjuvantTx12StartDate",
        "TTE_AdjuvantTx12EndDate",
        "TTE_AdjuvantTx13StartDate",
        "TTE_AdjuvantTx13EndDate",
        "TTE_AdjuvantTx14StartDate",
        "TTE_AdjuvantTx14EndDate",
        "TTE_AdjuvantTx15StartDate",
        "TTE_AdjuvantTx15EndDate",
        "TTE_DateRecurrenceProgression1",
        "TTE_DateRecurrenceProgression2",
        "TTE_DateRecurrenceProgression3",
        "TTE_DateRecurrenceProgression4",
        "TTE_DateRecurrenceProgression5",
        "TTE_DateRecurrenceProgression6",
        "TTE_DateRecurrenceProgression7",
        "TTE_DateRecurrenceProgression8",
        "TTE_DateRecurrenceProgression9",
        "TTE_DateRecurrenceProgression10",
        "TTE_DateRecurrenceProgression11",
        "TTE_DateRecurrenceProgression12",
        "TTE_DateRecurrenceProgression13",
        "TTE_DateRecurrenceProgression14",
        "TTE_DateRecurrenceProgression15",
        "TTE_SubsequentTx1StartDate",
        "TTE_SubsequentTx1EndDate",
        "TTE_SubsequentTx2StartDate",
        "TTE_SubsequentTx2EndDate",
        "TTE_SubsequentTx3StartDate",
        "TTE_SubsequentTx3EndDate",
        "TTE_SubsequentTx4StartDate",
        "TTE_SubsequentTx4EndDate",
        "TTE_SubsequentTx5StartDate",
        "TTE_SubsequentTx5EndDate",
        "TTE_SubsequentTx6StartDate",
        "TTE_SubsequentTx6EndDate",
        "TTE_SubsequentTx7StartDate",
        "TTE_SubsequentTx7EndDate",
        "TTE_SubsequentTx8StartDate",
        "TTE_SubsequentTx8EndDate",
        "TTE_SubsequentTx9StartDate",
        "TTE_SubsequentTx9EndDate",
        "TTE_SubsequentTx10StartDate",
        "TTE_SubsequentTx10EndDate",
        "TTE_SubsequentTx11StartDate",
        "TTE_SubsequentTx11EndDate",
        "TTE_SubsequentTx12StartDate",
        "TTE_SubsequentTx12EndDate",
        "TTE_SubsequentTx13StartDate",
        "TTE_SubsequentTx13EndDate",
        "TTE_SubsequentTx14StartDate",
        "TTE_SubsequentTx14EndDate",
        "TTE_SubsequentTx15StartDate",
        "TTE_SubsequentTx15EndDate",
        "TTE_Recurrence_PFS",
        "TTE_Death_OS_PFS",
        "TTE_Date_LastFollowUp",
        "TTE_DateLastSeen_OS",
        "TTE_OverallSurvival_Months",
        "TTE_PFS_Months",
        "Event_RecurrenceProgression",
        "Event_CancerDeath",
        "Event_DeathAllCauses",
        "Event_RecurrenceCancerDeath",
    ]

    # Query each table one by one
    for idx, row in all_tables.iterrows():
        schema = row["table_schema"]
        table = row["table_name"]

        try:
            # get the columns from each tables to remove any columns that do not exist
            peek_query = f"SELECT TOP 1 * FROM [{schema}].[{table}]"
            peek_df = pd.read_sql(peek_query, cnxn)
            existing_columns = peek_df.columns.to_list()

            # Filter the desired columns to just those that exist
            valid_columns = [col for col in target_columns if col in existing_columns]

            # Skip the table if none of the target columns exists
            if not valid_columns:
                print(
                    f"Skipping the table [{schema}].[{table}] - none of the target columns was found."
                )
                continue

            # format the columns with brackets for SQL querying
            columns_sql_str = ", ".join([f"[{col}]" for col in valid_columns])

            # Build and run the dynamic query to query each dataset
            single_query = f"""
            SELECT {columns_sql_str}
            FROM [{schema}].[{table}]
            """

            # generate the dataframe from the query output
            table_df = pd.read_sql(single_query, cnxn)

            # Add the schema and source table to the generated dataframe
            table_df["source_schema"] = schema
            table_df["source_table"] = table

            # Append the dataframe to the list of dataframes
            dfs.append(table_df)

        except (SQLAlchemyError, pd.errors.DatabaseError, ValueError, TypeError) as exc:
            print(f"Query failed for table [{schema}].[{table}]: {exc}")

    # Concatenate all the dataframes
    if dfs:
        compiled_data = pd.concat(dfs, ignore_index=True)
        print(f"Successfully combined {len(dfs)} datasets.")

        # Build the rename mapping dictionary first without mutating the DataFrame
        renamed_columns = {}

        for col in compiled_data.columns:
            # Primary specimen baseline mappings
            if col == "SelSpecAnonymisedID":
                renamed_columns[col] = "AnonymisedID_0"
            elif col == "SpecTreatmentSubgrouping":
                renamed_columns[col] = "TreatmentSubgrouping_0"
            elif col == "SpecTumourSite":
                renamed_columns[col] = "TumourSite_0"
            elif col == "SpecSurgeryDate":
                renamed_columns[col] = "SurgeryDate_0"
            elif col == "SpecBlockNum1":
                renamed_columns[col] = "BlockNum_0"
            elif col == "SpecBlockNum2":
                renamed_columns[col] = "BlockNum_1"
            elif col == "TTE_SpecimenSurgeryDate":
                renamed_columns[col] = "TTE_SurgeryDate_0"

            # Metastatic & Normal specimen mappings
            elif "AnonymisedID_" in col:
                match = re.search(r"AnonymisedID_([A-Za-z]+)(\d*)", col)
                if match:
                    num = match.group(2)
                    renamed_columns[col] = (
                        f"AnonymisedID_{int(num) + 1}" if num else "AnonymisedID_9"
                    )

            elif "_TreatmentSubgrouping" in col:
                match = re.search(r"([A-Za-z]+)(\d*)_TreatmentSubgrouping", col)
                if match:
                    num = match.group(2)
                    renamed_columns[col] = (
                        f"TreatmentSubgrouping_{int(num) + 1}"
                        if num
                        else "TreatmentSubgrouping_9"
                    )

            elif "_Site" in col:
                match = re.search(r"([A-Za-z]+)(\d*)_Site", col)
                if match:
                    num = match.group(2)
                    renamed_columns[col] = (
                        f"TumourSite_{int(num) + 1}" if num else "TumourSite_9"
                    )

            elif "_SurgeryDate" in col and "TTE" not in col:
                match = re.search(r"([A-Za-z]+)(\d*)_SurgeryDate", col)
                if match:
                    num = match.group(2)
                    renamed_columns[col] = (
                        f"SurgeryDate_{int(num) + 1}" if num else "SurgeryDate_9"
                    )

            elif "_BlockNumber" in col:
                match = re.search(r"([A-Za-z]+)(\d*)_BlockNumber", col)
                if match:
                    num = match.group(2)
                    renamed_columns[col] = (
                        f"BlockNum_{int(num) + 1}" if num else "BlockNum_9"
                    )

            elif (
                "_SurgeryDate" in col
                and "TTE" in col
                and ("MatchedMet" in col or "MatchedNormalTissue" in col)
            ):
                match = re.search(
                    r"(TTE_MatchedMet|TTE_MatchedNormalTissue)(\d*)_SurgeryDate", col
                )
                if match:
                    num = match.group(2)
                    renamed_columns[col] = (
                        f"TTE_SurgeryDate_{int(num) + 1}"
                        if num
                        else "TTE_SurgeryDate_9"
                    )

        # Apply column renaming once outside the loop
        compiled_data = compiled_data.rename(columns=renamed_columns)

        # Populate index 1 metadata from primary specimen (index 0)
        for stub in [
            "AnonymisedID",
            "TreatmentSubgrouping",
            "TumourSite",
            "SurgeryDate",
            "TTE_SurgeryDate",
        ]:
            col_0 = f"{stub}_0"
            col_1 = f"{stub}_1"
            if col_0 in compiled_data.columns and col_1 not in compiled_data.columns:
                compiled_data[col_1] = compiled_data[col_0]

        # Ensure column names in compiled_data are strictly unique
        compiled_data = compiled_data.loc[:, ~compiled_data.columns.duplicated()].copy()

        # Add a guaranteed unique row identifier
        compiled_data["row_id"] = np.arange(len(compiled_data))

        # Define stub names
        stub_names = [
            "AnonymisedID",
            "TreatmentSubgrouping",
            "TumourSite",
            "SurgeryDate",
            "TTE_SurgeryDate",
            "BlockNum",
        ]
        existing_stubs = [
            stub
            for stub in stub_names
            if any(col.startswith(f"{stub}_") for col in compiled_data.columns)
        ]

        # Melt dataframe to long format using 'row_id' as the explicit unique identifier
        compiled_data_long = pd.wide_to_long(
            compiled_data,
            stubnames=existing_stubs,
            i="row_id",
            j="Specimen_Index",
            sep="_",
            suffix=r"\d+",
        ).reset_index()

        # Assign Block_Type dynamically based on Specimen_Index
        def assign_block_type(idx):
            """
            This function assigns the block type to a sample based on the Specimen Index column values/
            """
            if idx in [0, 1]:
                return "Primary"
            elif idx == 9:
                return "Normal"
            else:
                return "Metastatic"

        compiled_data_long["Block_Type"] = compiled_data_long["Specimen_Index"].apply(
            assign_block_type
        )

        # replace [L] and (L) with L in the Block_Num column to avoid having to escape the brackets when querying data
        # and remove extra spaces
        compiled_data_long["BlockNum"] = (
            compiled_data_long["BlockNum"]
            .str.replace(r"\[L\]|\(L\)", "L", regex=True)
            .replace(r"\s+", "", regex=True)
            .str.strip()
        )

        # Add the Combined_Name column safely convert to string
        compiled_data_long["Combined_Name"] = (
            compiled_data_long["AnonymisedID"]
            .fillna("")
            .astype(str)
            .str.replace(" ", "", regex=False)
            + "-"
            + compiled_data_long["BlockNum"].astype(str)
        ).where(compiled_data_long["BlockNum"].notna(), pd.NA)

        # Clean up missing records
        compiled_data_long = compiled_data_long.dropna(subset=["AnonymisedID"])

        # trim extra spaces from all rows
        def trim_spaces(val):
            """
            This function trims spaces from all the rows of a dataframe
            """
            if isinstance(val, str):
                return re.sub(r"\s+", " ", val).strip()
            return val

        compiled_data_long = compiled_data_long.map(trim_spaces)

        # reorder the specific columns to place them at the start
        cols_to_reorder = [
            "AnonymisedID",
            "Block_Type",
            "BlockNum",
            "Combined_Name",
            "row_id",
            "source_schema",
            "source_table",
            "Specimen_Index",
            "SurgeryDate",
            "TTE_SurgeryDate",
            "TreatmentSubgrouping",
            "TumourSite",
        ]
        remaining_cols = [
            col for col in compiled_data_long.columns if col not in cols_to_reorder
        ]
        new_column_order = remaining_cols[:3] + cols_to_reorder + remaining_cols[3:]
        compiled_data_long = compiled_data_long[new_column_order]

        # Order the columns
        col_order = [
            "source_schema",
            "source_table",
            "CohortName",
            "DonorID",
            "AnonymisedID",
            "Block_Type",
            "BlockNum",
            "Combined_Name",
            "SpecIsRecurrence",
            "DateDX",
            "SpecHistologicalDiagnosis",
            "SpecSurgery",
            "SurgeryDate",
            "TumourSite",
            "TTE_SurgeryDate",
            "PrimSurgeryDate",
            "PrimTumourSite",
            "AnyImmunotherapy",
            "TreatmentSubgrouping",
            "AgeDX",
            "Gender",
            "StagingCriteria",
            "StagingGradingInfoProvidedFor",
            "StageAtDX",
            "pT",
            "pN",
            "pM",
            "cM",
            "TNM_Stage",
            "Grade",
            "ConcurrentMedicalConditions",
            "Comments_ClinicalTeam",
            "ResectionMarginsClear",
            "AnyNeoadjuvantTherapy",
            "ResponseToNeoTherapy_TRScore",
            "NeoadjuvantTx1",
            "NeoadjuvantTx1_Agent",
            "TumourResponse_NeoadjuvantTx1",
            "NeoadjuvantTx2",
            "NeoadjuvantTx2_Agent",
            "TumourResponse_NeoadjuvantTx2",
            "NeoadjuvantTx3",
            "NeoadjuvantTx3_Agent",
            "TumourResponse_NeoadjuvantTx3",
            "NeoadjuvantTx4",
            "NeoadjuvantTx4_Agent",
            "TumourResponse_NeoadjuvantTx4",
            "NeoadjuvantTx5",
            "NeoadjuvantTx5_Agent",
            "TumourResponse_NeoadjuvantTx5",
            "NeoadjuvantTx6",
            "NeoadjuvantTx6_Agent",
            "TumourResponse_NeoadjuvantTx6",
            "NeoadjuvantTx7",
            "NeoadjuvantTx7_Agent",
            "TumourResponse_NeoadjuvantTx7",
            "NeoadjuvantTx8",
            "NeoadjuvantTx8_Agent",
            "TumourResponse_NeoadjuvantTx8",
            "NeoadjuvantTx9",
            "NeoadjuvantTx9_Agent",
            "TumourResponse_NeoadjuvantTx9",
            "NeoadjuvantTx10",
            "NeoadjuvantTx10_Agent",
            "TumourResponse_NeoadjuvantTx10",
            "AdjuvantTx1",
            "AdjuvantTx1_Agent",
            "TumourResponse_AdjuvantTx1",
            "AdjuvantTx2",
            "AdjuvantTx2_Agent",
            "TumourResponse_AdjuvantTx2",
            "AdjuvantTx3",
            "AdjuvantTx3_Agent",
            "TumourResponse_AdjuvantTx3",
            "AdjuvantTx4",
            "AdjuvantTx4_Agent",
            "TumourResponse_AdjuvantTx4",
            "AdjuvantTx5",
            "AdjuvantTx5_Agent",
            "TumourResponse_AdjuvantTx5",
            "AdjuvantTx6",
            "AdjuvantTx6_Agent",
            "TumourResponse_AdjuvantTx6",
            "AdjuvantTx7",
            "AdjuvantTx7_Agent",
            "TumourResponse_AdjuvantTx7",
            "AdjuvantTx8",
            "AdjuvantTx8_Agent",
            "TumourResponse_AdjuvantTx8",
            "AdjuvantTx9",
            "AdjuvantTx9_Agent",
            "TumourResponse_AdjuvantTx9",
            "AdjuvantTx10",
            "AdjuvantTx10_Agent",
            "TumourResponse_AdjuvantTx10",
            "AdjuvantTx11",
            "AdjuvantTx11_Agent",
            "TumourResponse_AdjuvantTx11",
            "AdjuvantTx12",
            "AdjuvantTx12_Agent",
            "TumourResponse_AdjuvantTx12",
            "AdjuvantTx13",
            "AdjuvantTx13_Agent",
            "TumourResponse_AdjuvantTx13",
            "AdjuvantTx14",
            "AdjuvantTx14_Agent",
            "TumourResponse_AdjuvantTx14",
            "AdjuvantTx15",
            "AdjuvantTx15_Agent",
            "TumourResponse_AdjuvantTx15",
            "Sites_RecurrenceProgression1",
            "Sites_RecurrenceProgression2",
            "Sites_RecurrenceProgression3",
            "Sites_RecurrenceProgression4",
            "Sites_RecurrenceProgression5",
            "Sites_RecurrenceProgression6",
            "Sites_RecurrenceProgression7",
            "Sites_RecurrenceProgression8",
            "Sites_RecurrenceProgression9",
            "Sites_RecurrenceProgression10",
            "Sites_RecurrenceProgression11",
            "Sites_RecurrenceProgression12",
            "Sites_RecurrenceProgression13",
            "Sites_RecurrenceProgression14",
            "Sites_RecurrenceProgression15",
            "SubsequentTx1",
            "SubsequentTx1_Agent",
            "TumourResponse_SubsequentTx1",
            "SubsequentTx2",
            "SubsequentTx2_Agent",
            "TumourResponse_SubsequentTx2",
            "SubsequentTx3",
            "SubsequentTx3_Agent",
            "TumourResponse_SubsequentTx3",
            "SubsequentTx4",
            "SubsequentTx4_Agent",
            "TumourResponse_SubsequentTx4",
            "SubsequentTx5",
            "SubsequentTx5_Agent",
            "TumourResponse_SubsequentTx5",
            "SubsequentTx6",
            "SubsequentTx6_Agent",
            "TumourResponse_SubsequentTx6",
            "SubsequentTx7",
            "SubsequentTx7_Agent",
            "TumourResponse_SubsequentTx7",
            "SubsequentTx8",
            "SubsequentTx8_Agent",
            "TumourResponse_SubsequentTx8",
            "SubsequentTx9",
            "SubsequentTx9_Agent",
            "TumourResponse_SubsequentTx9",
            "SubsequentTx10",
            "SubsequentTx10_Agent",
            "TumourResponse_SubsequentTx10",
            "SubsequentTx11",
            "SubsequentTx11_Agent",
            "TumourResponse_SubsequentTx11",
            "SubsequentTx12",
            "SubsequentTx12_Agent",
            "TumourResponse_SubsequentTx12",
            "SubsequentTx13",
            "SubsequentTx13_Agent",
            "TumourResponse_SubsequentTx13",
            "SubsequentTx14",
            "SubsequentTx14_Agent",
            "TumourResponse_SubsequentTx14",
            "SubsequentTx15",
            "SubsequentTx15_Agent",
            "TumourResponse_SubsequentTx15",
            "TTE_PrimarySurgeryDate",
            "TTE_NeoadjuvantTx1EndDate",
            "TTE_NeoadjuvantTx1StartDate",
            "TTE_NeoadjuvantTx2EndDate",
            "TTE_NeoadjuvantTx2StartDate",
            "TTE_NeoadjuvantTx3EndDate",
            "TTE_NeoadjuvantTx3StartDate",
            "TTE_NeoadjuvantTx4EndDate",
            "TTE_NeoadjuvantTx4StartDate",
            "TTE_NeoadjuvantTx5EndDate",
            "TTE_NeoadjuvantTx5StartDate",
            "TTE_NeoadjuvantTx6EndDate",
            "TTE_NeoadjuvantTx6StartDate",
            "TTE_NeoadjuvantTx7EndDate",
            "TTE_NeoadjuvantTx7StartDate",
            "TTE_NeoadjuvantTx8EndDate",
            "TTE_NeoadjuvantTx8StartDate",
            "TTE_NeoadjuvantTx9EndDate",
            "TTE_NeoadjuvantTx9StartDate",
            "TTE_NeoadjuvantTx10EndDate",
            "TTE_NeoadjuvantTx10StartDate",
            "TTE_AdjuvantTx1StartDate",
            "TTE_AdjuvantTx1EndDate",
            "TTE_AdjuvantTx2StartDate",
            "TTE_AdjuvantTx2EndDate",
            "TTE_AdjuvantTx3StartDate",
            "TTE_AdjuvantTx3EndDate",
            "TTE_AdjuvantTx4StartDate",
            "TTE_AdjuvantTx4EndDate",
            "TTE_AdjuvantTx5StartDate",
            "TTE_AdjuvantTx5EndDate",
            "TTE_AdjuvantTx6StartDate",
            "TTE_AdjuvantTx6EndDate",
            "TTE_AdjuvantTx7StartDate",
            "TTE_AdjuvantTx7EndDate",
            "TTE_AdjuvantTx8StartDate",
            "TTE_AdjuvantTx8EndDate",
            "TTE_AdjuvantTx9StartDate",
            "TTE_AdjuvantTx9EndDate",
            "TTE_AdjuvantTx10StartDate",
            "TTE_AdjuvantTx10EndDate",
            "TTE_AdjuvantTx11StartDate",
            "TTE_AdjuvantTx11EndDate",
            "TTE_AdjuvantTx12StartDate",
            "TTE_AdjuvantTx12EndDate",
            "TTE_AdjuvantTx13StartDate",
            "TTE_AdjuvantTx13EndDate",
            "TTE_AdjuvantTx14StartDate",
            "TTE_AdjuvantTx14EndDate",
            "TTE_AdjuvantTx15StartDate",
            "TTE_AdjuvantTx15EndDate",
            "TTE_DateRecurrenceProgression1",
            "TTE_DateRecurrenceProgression2",
            "TTE_DateRecurrenceProgression3",
            "TTE_DateRecurrenceProgression4",
            "TTE_DateRecurrenceProgression5",
            "TTE_DateRecurrenceProgression6",
            "TTE_DateRecurrenceProgression7",
            "TTE_DateRecurrenceProgression8",
            "TTE_DateRecurrenceProgression9",
            "TTE_DateRecurrenceProgression10",
            "TTE_DateRecurrenceProgression11",
            "TTE_DateRecurrenceProgression12",
            "TTE_DateRecurrenceProgression13",
            "TTE_DateRecurrenceProgression14",
            "TTE_DateRecurrenceProgression15",
            "TTE_SubsequentTx1StartDate",
            "TTE_SubsequentTx1EndDate",
            "TTE_SubsequentTx2StartDate",
            "TTE_SubsequentTx2EndDate",
            "TTE_SubsequentTx3StartDate",
            "TTE_SubsequentTx3EndDate",
            "TTE_SubsequentTx4StartDate",
            "TTE_SubsequentTx4EndDate",
            "TTE_SubsequentTx5StartDate",
            "TTE_SubsequentTx5EndDate",
            "TTE_SubsequentTx6StartDate",
            "TTE_SubsequentTx6EndDate",
            "TTE_SubsequentTx7StartDate",
            "TTE_SubsequentTx7EndDate",
            "TTE_SubsequentTx8StartDate",
            "TTE_SubsequentTx8EndDate",
            "TTE_SubsequentTx9StartDate",
            "TTE_SubsequentTx9EndDate",
            "TTE_SubsequentTx10StartDate",
            "TTE_SubsequentTx10EndDate",
            "TTE_SubsequentTx11StartDate",
            "TTE_SubsequentTx11EndDate",
            "TTE_SubsequentTx12StartDate",
            "TTE_SubsequentTx12EndDate",
            "TTE_SubsequentTx13StartDate",
            "TTE_SubsequentTx13EndDate",
            "TTE_SubsequentTx14StartDate",
            "TTE_SubsequentTx14EndDate",
            "TTE_SubsequentTx15StartDate",
            "TTE_SubsequentTx15EndDate",
            "TTE_Date_LastFollowUp",
            "TTE_Death_OS_PFS",
            "TTE_DateLastSeen_OS",
            "TTE_OverallSurvival_Months",
            "TTE_PFS_Months",
            "TTE_Recurrence_PFS",
            "CurrentStatus",
            "CauseOfDeath",
            "DeathTumourRelated",
            "Event_CancerDeath",
            "Event_DeathAllCauses",
            "Event_RecurrenceCancerDeath",
            "Event_RecurrenceProgression",
        ]
        compiled_data_long = compiled_data_long[col_order]

        # Filter out null block numbers
        compiled_data_long = compiled_data_long.loc[
            ~compiled_data_long["BlockNum"].isna()
        ]

        # Drop duplicate rows
        compiled_data_long = compiled_data_long.drop_duplicates()

        # Map each row to the NHS number using the patient list
        # run the query to retrieve the patient list
        patientList = query_patientList()
        patientList = patientList[["SpecimenNumber", "NHS_Number"]]
        patientList = patientList.loc[~patientList["NHS_Number"].isna()]
        patientList.drop_duplicates()
        compiled_data_long = compiled_data_long.merge(
            patientList, left_on="AnonymisedID", right_on="SpecimenNumber", how="left"
        )
        compiled_data_long = compiled_data_long.drop_duplicates()
        compiled_data_long.drop(columns="SpecimenNumber")

        # Metastatic counts
        met_counts = (
            compiled_data_long[compiled_data_long["Block_Type"] == "Metastatic"]
            .drop_duplicates("Combined_Name")
            .groupby("NHS_Number")
            .size()
        )

        # Primary counts
        prim_counts = (
            compiled_data_long[compiled_data_long["Block_Type"] == "Primary"]
            .drop_duplicates("Combined_Name")
            .groupby("NHS_Number")
            .size()
        )

        # Pre-SoC counts
        pre_soc_counts = (
            compiled_data_long[
                compiled_data_long["TreatmentSubgrouping"]
                .str.lower()
                .isin(["treatment naive", "treatment naïve"])
            ]
            .drop_duplicates("Combined_Name")
            .groupby("NHS_Number")
            .size()
        )

        # Post-SoC counts
        post_soc_counts = (
            compiled_data_long[
                ~compiled_data_long["TreatmentSubgrouping"]
                .str.lower()
                .isin(["treatment naive", "treatment naïve"])
            ]
            .drop_duplicates("Combined_Name")
            .groupby("NHS_Number")
            .size()
        )

        # Mapping back to main dataframe
        compiled_data_long["count_patient_matched_mets"] = (
            compiled_data_long["NHS_Number"].map(met_counts).fillna(0).astype(int)
        )
        compiled_data_long["count_patient_matched_primaries"] = (
            compiled_data_long["NHS_Number"].map(prim_counts).fillna(0).astype(int)
        )
        compiled_data_long["count_patient_matched_pre_soc"] = (
            compiled_data_long["NHS_Number"].map(pre_soc_counts).fillna(0).astype(int)
        )
        compiled_data_long["count_patient_matched_post_soc"] = (
            compiled_data_long["NHS_Number"].map(post_soc_counts).fillna(0).astype(int)
        )

        # Subtract one from the count_patient_matched_mets for Block_Type Metastatic
        compiled_data_long.loc[
            compiled_data_long["Block_Type"] == "Metastatic",
            "count_patient_matched_mets",
        ] -= 1

        # Subtract one from the count_patient_matched_primaries for Block_Type Primary
        compiled_data_long.loc[
            compiled_data_long["Block_Type"] == "Primary",
            "count_patient_matched_primaries",
        ] -= 1

        # Subtract one from the pre SOC count when the sample itself is pre SOC
        compiled_data_long.loc[
            compiled_data_long["TreatmentSubgrouping"]
            .str.lower()
            .isin(["treatment naive", "treatment naïve"]),
            "count_patient_matched_pre_soc",
        ] -= 1

        # Subtract one from the post SOC count when the sample itself is post SOC
        compiled_data_long.loc[
            ~compiled_data_long["TreatmentSubgrouping"]
            .str.lower()
            .isin(["treatment naive", "treatment naïve"]),
            "count_patient_matched_post_soc",
        ] -= 1

        # Upload to SQL Database
        print("Uploading long format dataset to SQL...")
        compiled_data_long.to_sql(
            "master_list_long",
            engine,
            if_exists="replace",
            schema="imp_documents",
            index=False,
            chunksize=1000,
        )
        print("Upload complete!")

