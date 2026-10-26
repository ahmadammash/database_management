# Import the required libraries

import pandas as pd
from sqlalchemy import text
from sqlalchemy.exc import SQLAlchemyError

# Import functions from the project directory
from code_files.utilities.io.database import create_sql_engine
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()


# Define a function to fetch the schema names from the database
def fetch_schema_names():
    """
    This function gets the names of all the schemas available in the database.
    """

    # get the schemas from the database
    query = text("""
    SELECT schema_name
    FROM information_schema.schemata
    WHERE schema_name NOT IN ('information_schema', 'pg_catalog', 'guest')
                AND schema_name NOT LIKE '%sys%';
    """)

    # run the query to retrieve the schema names
    schema_names = pd.read_sql(query, engine)

    return schema_names


# Define the function to fetch the table names
def get_table_names():
    """
    This function gets the names of all the tables available in the database.
    """
    # Fetch the schema names from the SQL Server and convert it to a list
    schema_names = fetch_schema_names()
    schema_names = list(schema_names["schema_name"])

    # define the query to fetch the tables
    fetch_tables_query = """
    SELECT table_schema, table_name
    FROM information_schema.tables
    WHERE table_schema NOT IN ('information_schema', 'pg_catalog', 'guest')
                    AND table_schema NOT LIKE '%sys%'
    """

    # run the query
    fetched_tables = pd.read_sql(fetch_tables_query, cnxn)

    return fetched_tables


# Define a function to query the patient list for specific ids and return the cohort names
def query_patientList(
    cols: str = "*", id_in_clause: str = "", in_clause: str = "", where_clause: str = ""
):
    """
    This function builds a query to retrieve data from the patient list stored in the database.
    """

    # Define the base query
    query = f"SELECT {cols} FROM imp_documents.PatientList"  # {cols} is given to the function as an input

    # Define the list of where conditions
    where_conditions = []

    # If there is an IN Clause and an ID provided
    if in_clause.strip() and id_in_clause.strip():
        # convert the string IN Clause into a list and remove any extra spaces
        ids = [
            line.strip().replace(" ", "")
            for line in in_clause.splitlines()
            if line.strip()
        ]
        # if the list of IDs is valid
        if ids:
            # convert the list back to a string separated by commas and with single quotes around each list element
            quoted_ids = ", ".join([f"'{idx}'" for idx in ids])
            # add the IN Clause to the where conditions list
            where_conditions.append(f"{id_in_clause} IN ({quoted_ids})")

    # if there are (other) WHERE conditions provided, add them to the where_conditions list
    if where_clause.strip():
        where_conditions.append(where_clause.strip())

    # if the WHERE clause is valid, add each WHERE condition to the query by joining the list items with the "AND"
    # and starting off with the "WHERE"
    if where_conditions:
        query += " WHERE " + " AND ".join(where_conditions)

    # Execute the query
    # Create the connection engine
    engine = create_sql_engine()
    # Run the query and load the output into a dataframe
    with engine.connect() as conn:
        df = pd.read_sql(text(query), conn)
    return df


# Define a helper function to move a block of columns to a position after an anchor
def move_block_after(lst, block, anchor):
    """
    This function helps move a set of columns to position them after an anchor.
    """
    # Convert the list of columns to move into a set
    block = set(block)

    remaining = [x for x in lst if x not in block]
    to_move = [x for x in lst if x in block]
    result = []
    for x in remaining:
        result.append(x)
        if x == anchor:
            result.extend(to_move)
    return result


# Define a function to re-order follow up data columns from the query output
def reorder_query_output_FU(df):
    """
    This function reorders the columns output from the database query when the dataset is a follow up one. this ensures that the dataset
    columns follow the natural order they would usually have in any dataset.
    """
    cols = """DonorID
    CohortName
    AnyImmunotherapy
    SpecTreatmentSubgrouping
    SpecIsRecurrence
    SelSpecAnonymisedID
    SpecBlockNum1
    SpecBlockNum2
    SpecSurgeryDate
    SpecSurgery
    SpecTumourSite
    PrimSurgeryDate
    PrimTumourSite
    AnonymisedID_MatchedMet1
    MatchedMet1_BlockNumber
    MatchedMet1_Site
    MatchedMet1_SurgeryDate
    MatchedMet1_TreatmentSubgrouping
    AnonymisedID_MatchedMet2
    MatchedMet2_BlockNumber
    MatchedMet2_Site
    MatchedMet2_SurgeryDate
    MatchedMet2_TreatmentSubgrouping
    AnonymisedID_MatchedMet3
    MatchedMet3_BlockNumber
    MatchedMet3_Site
    MatchedMet3_SurgeryDate
    MatchedMet3_TreatmentSubgrouping
    AnonymisedID_MatchedMet4
    MatchedMet4_BlockNumber
    MatchedMet4_Site
    MatchedMet4_SurgeryDate
    MatchedMet4_TreatmentSubgrouping
    AnonymisedID_MatchedMet5
    MatchedMet5_BlockNumber
    MatchedMet5_Site
    MatchedMet5_SurgeryDate
    MatchedMet5_TreatmentSubgrouping
    AnonymisedID_MatchedMet6
    MatchedMet6_BlockNumber
    MatchedMet6_Site
    MatchedMet6_SurgeryDate
    MatchedMet6_TreatmentSubgrouping
    AnonymisedID_MatchedMet7
    MatchedMet7_BlockNumber
    MatchedMet7_Site
    MatchedMet7_SurgeryDate
    MatchedMet7_TreatmentSubgrouping
    Gender
    DateDX
    AgeDX
    SpecHistologicalDiagnosis
    StagingGradingInfoProvidedFor
    ResectionMarginsClear
    pT
    pN
    pM
    cM
    TNM_Stage
    Grade
    StagingCriteria
    StageAtDX
    AnyNeoadjuvantTherapy
    ResponseToNeoTherapy_TRScore
    NeoadjuvantTx1
    NeoadjuvantTx1_Agent
    TumourResponse_NeoadjuvantTx1
    NeoadjuvantTx2
    NeoadjuvantTx2_Agent
    TumourResponse_NeoadjuvantTx2
    NeoadjuvantTx3
    NeoadjuvantTx3_Agent
    TumourResponse_NeoadjuvantTx3
    NeoadjuvantTx4
    NeoadjuvantTx4_Agent
    TumourResponse_NeoadjuvantTx4
    NeoadjuvantTx5
    NeoadjuvantTx5_Agent
    TumourResponse_NeoadjuvantTx5
    NeoadjuvantTx6
    NeoadjuvantTx6_Agent
    TumourResponse_NeoadjuvantTx6
    NeoadjuvantTx7
    NeoadjuvantTx7_Agent
    TumourResponse_NeoadjuvantTx7
    NeoadjuvantTx8
    NeoadjuvantTx8_Agent
    TumourResponse_NeoadjuvantTx8
    NeoadjuvantTx9
    NeoadjuvantTx9_Agent
    TumourResponse_NeoadjuvantTx9
    NeoadjuvantTx10
    NeoadjuvantTx10_Agent
    TumourResponse_NeoadjuvantTx10
    AdjuvantTx1
    AdjuvantTx1_Agent
    TumourResponse_AdjuvantTx1
    AdjuvantTx2
    AdjuvantTx2_Agent
    TumourResponse_AdjuvantTx2
    AdjuvantTx3
    AdjuvantTx3_Agent
    TumourResponse_AdjuvantTx3
    AdjuvantTx4
    AdjuvantTx4_Agent
    TumourResponse_AdjuvantTx4
    AdjuvantTx5
    AdjuvantTx5_Agent
    TumourResponse_AdjuvantTx5
    AdjuvantTx6
    AdjuvantTx6_Agent
    TumourResponse_AdjuvantTx6
    AdjuvantTx7
    AdjuvantTx7_Agent
    TumourResponse_AdjuvantTx7
    AdjuvantTx8
    AdjuvantTx8_Agent
    TumourResponse_AdjuvantTx8
    AdjuvantTx9
    AdjuvantTx9_Agent
    TumourResponse_AdjuvantTx9
    AdjuvantTx10
    AdjuvantTx10_Agent
    TumourResponse_AdjuvantTx10
    AdjuvantTx11
    AdjuvantTx11_Agent
    TumourResponse_AdjuvantTx11
    AdjuvantTx12
    AdjuvantTx12_Agent
    TumourResponse_AdjuvantTx12
    AdjuvantTx13
    AdjuvantTx13_Agent
    TumourResponse_AdjuvantTx13
    AdjuvantTx14
    AdjuvantTx14_Agent
    TumourResponse_AdjuvantTx14
    AdjuvantTx15
    AdjuvantTx15_Agent
    TumourResponse_AdjuvantTx15
    Sites_RecurrenceProgression1
    SubsequentTx1
    SubsequentTx1_Agent
    TumourResponse_SubsequentTx1
    SubsequentTx2
    SubsequentTx2_Agent
    TumourResponse_SubsequentTx2
    SubsequentTx3
    SubsequentTx3_Agent
    TumourResponse_SubsequentTx3
    SubsequentTx4
    SubsequentTx4_Agent
    TumourResponse_SubsequentTx4
    SubsequentTx5
    SubsequentTx5_Agent
    TumourResponse_SubsequentTx5
    SubsequentTx6
    SubsequentTx6_Agent
    TumourResponse_SubsequentTx6
    SubsequentTx7
    SubsequentTx7_Agent
    TumourResponse_SubsequentTx7
    SubsequentTx8
    SubsequentTx8_Agent
    TumourResponse_SubsequentTx8
    SubsequentTx9
    SubsequentTx9_Agent
    TumourResponse_SubsequentTx9
    SubsequentTx10
    SubsequentTx10_Agent
    TumourResponse_SubsequentTx10
    SubsequentTx11
    SubsequentTx11_Agent
    TumourResponse_SubsequentTx11
    SubsequentTx12
    SubsequentTx12_Agent
    TumourResponse_SubsequentTx12
    SubsequentTx13
    SubsequentTx13_Agent
    TumourResponse_SubsequentTx13
    SubsequentTx14
    SubsequentTx14_Agent
    TumourResponse_SubsequentTx14
    SubsequentTx15
    SubsequentTx15_Agent
    TumourResponse_SubsequentTx15
    CurrentStatus
    DeathTumourRelated
    CauseOfDeath
    Comments_ClinicalTeam
    TTE_NeoadjuvantTx1StartDate
    TTE_NeoadjuvantTx1EndDate
    TTE_SpecimenSurgeryDate
    TTE_PrimarySurgeryDate
    TTE_MatchedMet1_SurgeryDate
    TTE_MatchedMet2_SurgeryDate
    TTE_MatchedMet3_SurgeryDate
    TTE_MatchedMet4_SurgeryDate
    TTE_MatchedMet5_SurgeryDate
    TTE_MatchedMet6_SurgeryDate
    TTE_MatchedMet7_SurgeryDate
    TTE_AdjuvantTx1StartDate
    TTE_AdjuvantTx1EndDate
    TTE_DateRecurrenceProgression1
    TTE_SubsequentTx1StartDate
    TTE_SubsequentTx1EndDate
    TTE_Recurrence_PFS
    TTE_Death_OS_PFS
    TTE_Date_LastFollowUp
    TTE_DateLastSeen_OS
    TTE_OverallSurvival_Months
    TTE_PFS_Months
    Event_RecurrenceProgression
    Event_CancerDeath
    Event_DeathAllCauses
    Event_RecurrenceCancerDeath
    AnonymisedID_NormalTissue
    MatchedNormalTissue_BlockNumber
    MatchedNormalTissue_Site
    MatchedNormalTissue_SurgeryDate
    MatchedNormal_TreatmentSubgrouping
    Sites_RecurrenceProgression2
    Sites_RecurrenceProgression3
    Sites_RecurrenceProgression4
    Sites_RecurrenceProgression5
    Sites_RecurrenceProgression6
    Sites_RecurrenceProgression7
    Sites_RecurrenceProgression8
    Sites_RecurrenceProgression9
    Sites_RecurrenceProgression10
    Sites_RecurrenceProgression11
    Sites_RecurrenceProgression12
    Sites_RecurrenceProgression13
    Sites_RecurrenceProgression14
    Sites_RecurrenceProgression15
    ConcurrentMedicalConditions
    TTE_NeoadjuvantTx2StartDate
    TTE_NeoadjuvantTx2EndDate
    TTE_NeoadjuvantTx3StartDate
    TTE_NeoadjuvantTx3EndDate
    TTE_NeoadjuvantTx4StartDate
    TTE_NeoadjuvantTx4EndDate
    TTE_NeoadjuvantTx5StartDate
    TTE_NeoadjuvantTx5EndDate
    TTE_NeoadjuvantTx6StartDate
    TTE_NeoadjuvantTx6EndDate
    TTE_NeoadjuvantTx7StartDate
    TTE_NeoadjuvantTx7EndDate
    TTE_NeoadjuvantTx8StartDate
    TTE_NeoadjuvantTx8EndDate
    TTE_NeoadjuvantTx9StartDate
    TTE_NeoadjuvantTx9EndDate
    TTE_NeoadjuvantTx10StartDate
    TTE_NeoadjuvantTx10EndDate
    TTE_MatchedNormalTissue_SurgeryDate
    TTE_AdjuvantTx2StartDate
    TTE_AdjuvantTx2EndDate
    TTE_AdjuvantTx3StartDate
    TTE_AdjuvantTx3EndDate
    TTE_AdjuvantTx4StartDate
    TTE_AdjuvantTx4EndDate
    TTE_AdjuvantTx5StartDate
    TTE_AdjuvantTx5EndDate
    TTE_AdjuvantTx6StartDate
    TTE_AdjuvantTx6EndDate
    TTE_AdjuvantTx7StartDate
    TTE_AdjuvantTx7EndDate
    TTE_AdjuvantTx8StartDate
    TTE_AdjuvantTx8EndDate
    TTE_AdjuvantTx9StartDate
    TTE_AdjuvantTx9EndDate
    TTE_AdjuvantTx10StartDate
    TTE_AdjuvantTx10EndDate
    TTE_AdjuvantTx11StartDate
    TTE_AdjuvantTx11EndDate
    TTE_AdjuvantTx12StartDate
    TTE_AdjuvantTx12EndDate
    TTE_AdjuvantTx13StartDate
    TTE_AdjuvantTx13EndDate
    TTE_AdjuvantTx14StartDate
    TTE_AdjuvantTx14EndDate
    TTE_AdjuvantTx15StartDate
    TTE_AdjuvantTx15EndDate
    TTE_DateRecurrenceProgression2
    TTE_DateRecurrenceProgression3
    TTE_DateRecurrenceProgression4
    TTE_DateRecurrenceProgression5
    TTE_DateRecurrenceProgression6
    TTE_DateRecurrenceProgression7
    TTE_DateRecurrenceProgression8
    TTE_DateRecurrenceProgression9
    TTE_DateRecurrenceProgression10
    TTE_DateRecurrenceProgression11
    TTE_DateRecurrenceProgression12
    TTE_DateRecurrenceProgression13
    TTE_DateRecurrenceProgression14
    TTE_DateRecurrenceProgression15
    TTE_SubsequentTx2StartDate
    TTE_SubsequentTx2EndDate
    TTE_SubsequentTx3StartDate
    TTE_SubsequentTx3EndDate
    TTE_SubsequentTx4StartDate
    TTE_SubsequentTx4EndDate
    TTE_SubsequentTx5StartDate
    TTE_SubsequentTx5EndDate
    TTE_SubsequentTx6StartDate
    TTE_SubsequentTx6EndDate
    TTE_SubsequentTx7StartDate
    TTE_SubsequentTx7EndDate
    TTE_SubsequentTx8StartDate
    TTE_SubsequentTx8EndDate
    TTE_SubsequentTx9StartDate
    TTE_SubsequentTx9EndDate
    TTE_SubsequentTx10StartDate
    TTE_SubsequentTx10EndDate
    TTE_SubsequentTx11StartDate
    TTE_SubsequentTx11EndDate
    TTE_SubsequentTx12StartDate
    TTE_SubsequentTx12EndDate
    TTE_SubsequentTx13StartDate
    TTE_SubsequentTx13EndDate
    TTE_SubsequentTx14StartDate
    TTE_SubsequentTx14EndDate
    TTE_SubsequentTx15StartDate
    TTE_SubsequentTx15EndDate
    """
    cols = [col.strip() for col in cols.splitlines()]
    cols = [col for col in cols if col != ""]
    extra_matched_met_cols = [
        "AnonymisedID_MatchedMet6",
        "MatchedMet6_BlockNumber",
        "MatchedMet6_Site",
        "MatchedMet6_SurgeryDate",
        "MatchedMet6_TreatmentSubgrouping",
        "AnonymisedID_MatchedMet7",
        "MatchedMet7_BlockNumber",
        "MatchedMet7_Site",
        "MatchedMet7_SurgeryDate",
        "MatchedMet7_TreatmentSubgrouping",
    ]
    anchor_extra_matched_met_cols = "MatchedMet5_TreatmentSubgrouping"

    extra_tte_matched_met_cols = [
        "TTE_MatchedMet6_SurgeryDate",
        "TTE_MatchedMet7_SurgeryDate",
    ]
    anchor_extra_tte_matched_met_cols = "TTE_MatchedMet5_SurgeryDate"

    normal_specimen_cols = [
        "AnonymisedID_NormalTissue",
        "MatchedNormalTissue_BlockNumber",
        "MatchedNormalTissue_Site",
        "MatchedNormalTissue_SurgeryDate",
        "MatchedNormal_TreatmentSubgrouping",
    ]
    anchor_normal_specimen_cols = "MatchedMet7_TreatmentSubgrouping"

    sites_recurr_prog_cols = [
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
    ]
    anchor_recurr_prog_cols = "Sites_RecurrenceProgression1"

    conc_med_cond_cols = ["ConcurrentMedicalConditions"]
    anchor_conc_med_cond_cols = "CauseOfDeath"

    tte_neoadjuvant_cols = [
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
    ]
    anchor_tte_neoadjuvant_cols = "TTE_NeoadjuvantTx1StartDate"

    tte_adjuvant_cols = [
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
    ]
    anchor_tte_adjuvant_cols = "TTE_AdjuvantTx1StartDate"

    tte_recurr_prog_cols = [
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
    ]
    anchor_tte_recurr_prog_cols = "TTE_DateRecurrenceProgression1"

    tte_subsequent_cols = [
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
    ]
    anchor_tte_subsequent_cols = "TTE_SubsequentTx1StartDate"

    tte_normal_cols = ["TTE_MatchedNormalTissue_SurgeryDate"]
    anchor_tte_normal_cols = "TTE_MatchedMet7_SurgeryDate"

    cols_to_move = [
        extra_matched_met_cols,
        extra_tte_matched_met_cols,
        sites_recurr_prog_cols,
        conc_med_cond_cols,
        tte_neoadjuvant_cols,
        tte_adjuvant_cols,
        tte_recurr_prog_cols,
        tte_subsequent_cols,
        normal_specimen_cols,
        tte_normal_cols,
    ]
    anchors = [
        anchor_extra_matched_met_cols,
        anchor_extra_tte_matched_met_cols,
        anchor_recurr_prog_cols,
        anchor_conc_med_cond_cols,
        anchor_tte_neoadjuvant_cols,
        anchor_tte_adjuvant_cols,
        anchor_tte_recurr_prog_cols,
        anchor_tte_subsequent_cols,
        anchor_normal_specimen_cols,
        anchor_tte_normal_cols,
    ]
    for block, anchor in zip(cols_to_move, anchors):
        cols = move_block_after(lst=cols, block=block, anchor=anchor)

    existing_cols = [col for col in cols if col in df.columns]

    # Add any columns fetched from the database that weren't in the master list
    extra_df_cols = [col for col in df.columns if col not in existing_cols]
    final_col_order = existing_cols + extra_df_cols

    # apply the reordering
    df = df[final_col_order]
    return df


# Define a function to query tables from the database
def query_data(
    tables: str,  # the tables to be queried
    columns: str | None = "*",  # the columns to be retrieved from the selected tables
    conditions: list[str] | None = None,  # the conditions to be applied
    id_in_clause: str | None = None,  # the ID column used in the IN clause
    in_clause: str | None = None,  # the list of ID values which the ID column values can be
    id_exclusions_clause: str | None = None, # the ID column used in the exclusions IN clause
    exclusions_clause: str | None = None, # the list of ID values to exclude
    order_by: str | None = None,  # the column name to order the query output by
    duplicate_results=True,  # whether to duplicate the results when duplicate ID numbers are provided in the IN clause
    isSupplementary=False,  # whether the tables being queried are supplementary data tables
    isMasterList=False,  # whether the table being queried is the master list
    top_rows: int | None = None,  # the number of rows to show in the query output
    applyReplacements=False,  # whether to apply replacements to the boolean columns (i.e., replace 1 with Yes)
    return_firstBlockOnly=False,  # whether to return the first primary or met block only from the master list
    return_FollowUpOnly=False,  # whether to return the follow up data only from the master list
    return_firstAvailableRecord=False,
) -> tuple[pd.DataFrame, str]:
    # when querying the master list
    """This function enables querying from the database by providing a number of arguments used to build and run the SQL query."""

    # Create a connection engine to the SQL database
    engine = create_sql_engine()

    # Define a list to store the dataframes which will be created from the tables to be queried
    dataframes_list = []

    # Define an empty string to store the current query statement
    current_statement = ""

    # Define a list to store the list of IDs that will form the IN Clause and the exclusion IN clause
    ids_list = []
    exclusion_ids_list = []

    # If an IN Clause and an ID for it are provided
    if in_clause and id_in_clause:
        ids_list = [
            line.strip()
            .strip("'")
            .strip('"')  # remove extra spaces and quotation marks from the IDs
            for line in in_clause.splitlines()  # use splitlines to convert the text which is across multiple lines into a list of IDs
            if line.strip()
        ]

    # If an Exclusion IN clause and ID are provided
    if exclusions_clause and id_exclusions_clause:
        exclusion_ids_list = [
            line.strip()
            .strip("'")
            .strip('"')
            for line in exclusions_clause.splitlines()
            if line.strip()
        ]

    # Create the set of unique IDs for the IN clause and the exlcusion IN clause
    unique_ids = list(set(ids_list)) if ids_list else []
    unique_exclusion_ids = list(set(exclusion_ids_list)) if exclusion_ids_list else []
    
    # Split using the '|' delimiter to protect table string structures
    table_array = tables.split("|") if "|" in tables else tables.split()

    with engine.connect() as conn:
        # Loop through the selected tables
        for tbl in table_array:
            # Remove existing brackets and trim whitespaces
            tbl = tbl.replace("]", "").replace("[", "").strip()

            # If the tbl variable is undefined or null, skip the current loop iteration
            if not tbl:
                continue

            # Format schema and table names with MSSQL square brackets []
            if "." in tbl:
                schema, table_name = tbl.split(".")
                table_ref = f"[{schema}].[{table_name}]"

            # If there is no ".", the schema name is not included - what is left would be the table name
            else:
                schema = None
                table_name = tbl
                table_ref = f"[{tbl}]"

            # Parse requested columns safely using comma distribution
            if columns != "*":
                requested_cols = [c.strip() for c in columns.split(",") if c.strip()]

                # Filter for the appropriate schema name
                schema_filter = f"AND TABLE_SCHEMA = '{schema}'" if schema else ""

                # Define a query statement to check whether a selected column exists in a table - while also including the filter for the schema name
                check_cols_stmt = f"""
                    SELECT COLUMN_NAME 
                    FROM INFORMATION_SCHEMA.COLUMNS 
                    WHERE TABLE_NAME = '{table_name}' {schema_filter}
                """

                # try to run the query to check which columns exist in the table
                try:
                    existing_cols_df = pd.read_sql(text(check_cols_stmt), conn)
                    # get the names of the unique and existing columns from the queried table
                    # remember the main for loop is looping through each table in the selected tables array
                    existing_cols = set(existing_cols_df["COLUMN_NAME"].tolist())

                    # Create list of valid columns wrapped in brackets - the columns which are both requested and exist in the table
                    valid_cols = [
                        f"[{c}]" for c in requested_cols if c in existing_cols
                    ]

                    # Ensure ID column for IN clause is included if present
                    if (
                        id_in_clause
                        and id_in_clause in existing_cols
                        and f"[{id_in_clause}]" not in valid_cols
                    ):
                        valid_cols.append(f"[{id_in_clause}]")

                    # Ensure ID column for exclusion IN clause is included if present
                    if (
                        id_exclusions_clause
                        and id_exclusions_clause in existing_cols
                        and f"[{id_exclusions_clause}]" not in valid_cols
                    ):
                        valid_cols.append(f"[{id_exclusions_clause}]")

                    # Skip the table if no matching columns exists
                    if not valid_cols:
                        print(
                            f"Skipping table {tbl}: None of the requested columns exist here."
                        )
                        continue

                    # create the column selection by joining the valid columns using a comma
                    col_selection = ", ".join(valid_cols)

                # If running the query to check the validity of columns fails, return an error
                except (SQLAlchemyError, ValueError) as ex:
                    print(f"Could not verify columns for {tbl}. Error: {ex}")
                    col_selection = ", ".join([f"[{c}]" for c in requested_cols])

            # When no specific columns are provided, query all the columns
            else:
                col_selection = "*"

            # Build the SELECT statement
            # Define the top clause using the top_rows variable's value
            top_clause = f"TOP ({top_rows})" if top_rows else ""

            # Build the first part of the current statement by combining the TOP clause with the column names to be selected
            # and separated by commas - alongside the table_schema.table_name
            # remember this is building a current statement for each table in the array of selected tables
            current_statement = (
                f"SELECT {top_clause} {col_selection} FROM {table_ref}".replace(
                    "  ", " "
                )
            )

            # Define a list to store the filters which will be applied to the selected tables
            filters = []

            # Add the IN clause condition
            if unique_ids and id_in_clause:
                formatted_in_clause = ", ".join([f"'{idx}'" for idx in unique_ids])
                filters.append(f"{table_ref}.[{id_in_clause}] IN ({formatted_in_clause})")

            # Add the exclusion IN clause condition
            if unique_exclusion_ids and id_exclusions_clause:
                formatted_exclusion_in_clause = ", ".join([f"'{idx}'" for idx in unique_exclusion_ids])
                filters.append(
                    f"{table_ref}.[{id_exclusions_clause}] NOT IN ({formatted_exclusion_in_clause})"
                )
            
            # Add user-defined WHERE conditions if any are provided
            if conditions:
                # use .extend to add the conditions to the end of the list
                filters.extend(conditions)

            # Combine filters with SQL 'AND'
            if filters:
                current_statement += " WHERE " + " AND ".join(filters)

            # Add the ORDER BY clause
            if order_by:
                clean_order_col = order_by.split([0]).replace("[", "").replace("]", "").strip()
                direction = "DESC" if "DESC" in order_by.upper() else "ASC"

                # check column existence if columns were validated
                if columns == "*" or clean_order_col in existing_cols:
                    current_statement = f"ORDER BY [{clean_order_col}] {direction}"

            # Execute the query for the current table
            try:
                # capture the query output from the current statement in df_temp
                df_temp = pd.read_sql(text(current_statement), conn)
                # if the dataframe is not empty, append it to the list of dataframes
                if not df_temp.empty:
                    dataframes_list.append(df_temp)
            # if the query fails, output an error
            except (SQLAlchemyError, ValueError) as e:
                print(f"Error querying table {tbl}: {e}")
                print(f"Failed SQL: {current_statement}")

    # If the list of dataframes is not empty, create the combined df which concatenates all the dataframes together
    if dataframes_list:
        combined_df = pd.concat(dataframes_list, axis=0, ignore_index=True)
        final_df = combined_df

        # Pad missing columns with None if explicit columns requested
        if columns != "*":
            # create the list of cleaned column names by splitting by commas
            clean_requested = [c.strip() for c in columns.split(",") if c.strip()]
            # loop through the list of clean column names
            for col in clean_requested:
                # if the column is not in the combined dataframe, add it to the dataframe and set all its values to None (null)
                if col not in combined_df.columns:
                    combined_df[col] = None

        # Drop duplicates based on the primary ID column if available
        if id_in_clause in combined_df.columns:
            combined_df = combined_df.drop_duplicates()

        # Handle exact filtering and duplication from driver ID list
        # If an ID and IN Clause are provided
        if ids_list and id_in_clause in combined_df.columns:
            # convert the ID column to the string type and strip extra spaces from it
            combined_df[id_in_clause] = (
                combined_df[id_in_clause].astype(str).str.strip()
            )
            # create the list of clean IDs
            clean_ids = [str(x).strip() for x in ids_list]
            # create the driver_df from the dictionary which maps the IN Clause ID to the list of clean IDs provided
            # the resulting dataframe is a single column named as the id_in_clause with values being all the ids in ids_list
            driver_df = pd.DataFrame({id_in_clause: clean_ids})

            # if the user requested keeping the duplicate results, create the final df by merging the driver_df (containing
            # all the id values in ids_list) with the combined df so that each id, regardless whether its duplicated or not,
            # gets mapped to the correct data in combined_df
            if duplicate_results:
                final_df = pd.merge(
                    driver_df, combined_df, on=id_in_clause, how="inner"
                )
            # otherwise, the final_df would be the combined_df (no duplicates) filtered for the IDs which are in the clean_ids list
            else:
                final_df = combined_df[combined_df[id_in_clause].isin(clean_ids)]

        # Handle sorting of follow up or prelim data (not supplementary data and not the master list)
        # if the data queried is not supplementary and is not the master list
        if not isSupplementary and not any(
            col in final_df.columns
            for col in ["source_schema", "source_table", "Block_Type", "Combined_Name"]
        ):
            # if the order_by clause is provided and the ID column is in the final merged dataframe
            if order_by and id_in_clause in final_df.columns:
                # Define the sort_col variable as the column to order by and strip the brackets of it
                # order_by.split()[0] splits by space and takes the first word so that ASC/DESC are left out
                sort_col = order_by.split()[0].replace("[", "").replace("]", "")
                # set the order direction variable depending on what is entered - ASC by default
                ascending = not "DESC" in order_by.upper()
                # make sure the column to sort by is in the final dataframe
                if sort_col in final_df.columns:
                    # sort the final dataframe by the requested column
                    final_df = final_df.sort_values(by=sort_col, ascending=ascending)
            # if no specific column is provided for ordering and when DonorID is present in the final dataframe, order the output by it (ASC)
            elif "DonorID" in final_df.columns:
                final_df = final_df.sort_values(by="DonorID")

            # the final step is to strip empty or extra spaces from the final dataframe columns
            final_df.columns = final_df.columns.str.strip()

            # Since the data is not supplementary and is not the master list, reorder it using the follow up data order
            final_df = reorder_query_output_FU(final_df)

        # Handle sorting for supplementary data
        elif isSupplementary:
            # if the order by clause is provided and the ID column for the IN clause is in the final dataframe
            if order_by and id_in_clause in final_df.columns:
                # create the sort_col by extracting the first word from the order by clause and stripping the brackets from them
                sort_col = order_by.split()[0].replace("[", "").replace("]", "")
                # assign the sort order using the value given in the order by input box
                ascending = not "DESC" in order_by.upper()
                # sort the dataframe by the column provided
                if sort_col in final_df.columns:
                    final_df = final_df.sort_values(by=sort_col, ascending=ascending)
            # if the order by clause is not provided and the DonorID column is in the final dataframe, sort by it
            elif "DonorID" in final_df.columns:
                final_df = final_df.sort_values(by="DonorID")

            # strip extra spaces from the final dataframe column names
            final_df.columns = final_df.columns.str.strip()

            # reorder the columns in the output using the column order for supplementary data
            final_df = reorder_query_output_Supp(final_df)

        # Handle sorting for querying the master list
        if isMasterList:
            master_col_order = [
                "source_schema",
                "source_table",
                "DonorID",
                "CohortName",
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
            ]

            # Get existing columns in desired order, then append remaining columns
            existing_master_cols = [
                c for c in master_col_order if c in final_df.columns
            ]
            other_cols = [c for c in final_df.columns if c not in existing_master_cols]
            final_df = final_df[existing_master_cols + other_cols]

        # Apply boolean value replacements if requested
        if applyReplacements:
            final_df = apply_bool_replacements(final_df)

        # Sort the final dataframe so FollowUp_Anon records are prioritised over Prelim_Aon records
        if "source_table" in final_df.columns:
            final_df["_table_priority"] = final_df["source_table"].apply(
                lambda x: 1 if "FollowUp_Anon" in str(x) else (2 if "Prelim_Anon" in str(x) else 3)
            )
            final_df = final_df.sort_values(by=["_table_priority"]).drop(columns=["_table_priority"])        
        


        # if the first available records (follow up or prelim records) is required - can still return 2 records if the
        # specimen has 2 block numbers associated with it
        if return_firstAvailableRecord:
            # Identify all Combined_Names that have at least one FollowUp_Anon record
            has_followup_mask = final_df["source_table"].str.contains("FollowUp_Anon", na=False)
            combined_names_with_followup = set(final_df.loc[has_followup_mask, "Combined_Name"].dropna())

            # Drop Prelim_Anon records ONLY if a FollowUp_Anon record exists for that Combined_Name
            is_prelim = final_df["source_table"].str.contains("Prelim_Anon", na=False)
            has_followup = final_df["Combined_Name"].isin(combined_names_with_followup)

            final_df = final_df.loc[~(is_prelim & has_followup)]


        # if only the follow up data tables data should be returned from the master list (excluding the prelim only rows)
        if return_FollowUpOnly:
            final_df = final_df.loc[
                final_df["source_table"].str.contains("FollowUp_Anon")
            ]        
        
        # if only one block of a specific type is required, filter the final dataframe for the first primary or met block
        # applicable to master list querying only
        if return_firstBlockOnly:
            final_df = final_df.drop_duplicates(
                subset=["DonorID", "Block_Type"], keep="first"
            )
        
        # Apply user ordering as the final step before returning
        if order_by:
            sort_col = order_by.split()[0].replace("[", "").replace("]", "").strip()
            ascending = "DESC" not in order_by.upper()

            if sort_col in final_df.columns:
                final_df = final_df.sort_values(by=sort_col, ascending=ascending)
        elif "DonorID" in final_df.columns:
            final_df = final_df.sort_values(by="DonorID", ascending=True)

        return final_df.reset_index(drop=True), current_statement

    else:
        print("No data was found or all tables returned empty results.")
        # Return empty DataFrame along with the last statement context to maintain 2-tuple contract
        return pd.DataFrame(), current_statement


# Define a function to re-order supplementary data columns
def reorder_query_output_Supp(df):
    """
    This function reorders the output of the database query function when the data is of the supplementary type - to ensure
    the order of columns makes sense.
    """
    # get all supplementary data columns to build
    query = """
    SELECT table_schema, table_name
    FROM information_schema.tables
    WHERE table_schema NOT IN ('information_schema', 'pg_catalog', 'guest')
                    AND table_schema NOT LIKE '%sys%'
                    AND (table_name LIKE '%_Supplementary_Anon%');
    """
    all_supplementary_tbl_names = pd.read_sql(query, cnxn)
    all_supplementary_cols = []

    # Loop through the supplementary data tables safely without mutating input frames
    for idx, row in all_supplementary_tbl_names.iterrows():
        tbl = row["table_name"]
        schema = row["table_schema"]
        query_meta = f"""
        SELECT TOP 0 * FROM [{schema}].[{tbl}]
        """
        # FIX 1: Use df_meta so we don't destroy your filtered input frame
        # FIX 2: Added "TOP 0" so it pulls ONLY the structural headers without reading millions of rows
        df_meta = pd.read_sql(query_meta, cnxn)
        all_supplementary_cols.append(list(df_meta.columns))

    all_supplementary_cols = [
        col for sublist in all_supplementary_cols for col in sublist
    ]
    all_supplementary_cols = sorted(set(all_supplementary_cols))

    # Define the logical grouped order
    supp_global_col_order = [
        # --- Group 1: Demographics & Master IDs ---
        "CohortName",
        "DonorID",
        "AnonymisedID_Data_Taken_From",
        "SelSpecAnonymisedID",
        "Details_Provided_For",
        "Ethnicity",
        # --- Group 2: Specimen / Blocks Metadata ---
        "Block",
        "Spec_Number_Used_For_Testing",
        "SpecNumber_Used_For_MolecularTesting_IfDifferent",
        "Block_Used_For_Screening",
        "Secondary_Block_Used_For_Screening",
        "Third_Block_Used_For_Screening",
        "Block_Used_For_Testing",
        "Block_Used_For_Testing_P16",
        "Block_Used_For_LynchScreening_MMR_IHC",
        "Block_Used_For_MMR_IHC",
        "Spec_Used_For_Screening_IfDifferent",
        "Spec_Used_For_MMR_Screening_IfDifferent",
        "Spec_Used_For_P53_Testing_IfDifferent",
        # --- Group 3: Core Pathology & Staging ---
        "Tumour_Site_Area",
        "Side_L_R",
        "Tumour_Type",
        "Differentiation",
        "Tumour_Necrosis",
        "Mitotic_Count",
        "Ki67",
        "Stage_AJCC7",
        "Stage_AJCC8",
        "pN_Stage_AJCC7",
        "pN_Stage_AJCC8",
        "FIGO_Stage",
        "WHO_Grade",
        "Num_Pos_LymphNodes",
        "Nottingham_Prognostic_Index",
        "Leibovich_Score",
        # --- Group 4: Core Molecular / NGS Testing Context ---
        "Detection_Platform",
        "Genes_Tested",
        "Pathogenic_Germline_Gene_Variants_Identified",
        "Any_Other_Variants_Detected",
        "Other_Genetic_Lesions",
        # --- Group 5: Common Somatic Mutations (Pan-Cancer) ---
        "BRAF_Mutation",
        "BRAF_Mutation_Shortened",
        "EGFR_Mutation",
        "EGFR_Mutation_Shortened",
        "EGFR_Amplification",
        "KRAS_Mutation",
        "KRAS_Mutation_Shortened",
        "NRAS_Mutation",
        "NRAS_Mutation_Shortened",
        "ALK_Rearrangement_Fusion",
        # --- Group 6: Mismatch Repair (MMR) & Lynch Syndrome ---
        "Lynch_Syndrome",
        "MMR_Result",
        "MMR_Result_Normal_Abnormal",
        "Inferred_MSI_Status",
        "MSI_Status",
        "MSIH_MSS",
        "MLH1_IHC_Pos_Neg_ClonalNeg",
        "MSH2_IHC_Pos_Neg_ClonalNeg",
        "MSH6_IHC_Pos_Neg_ClonalNeg",
        "PMS2_IHC_Pos_Neg_ClonalNeg",
        "MLH1_Promoter_Hypermethylation_DetectionResult",
        # --- Group 7: Breast / Ovarian Related Biomarkers ---
        "Spec_Receptor_Status",
        "Spec_Is_Hormone_Receptor_Pos_And_HER2_Neg",
        "ER_Status",
        "PR_Status",
        "HER2_IHC_Score",
        "HER2_FISH_Score",
        "Oncotype_DX_Score_DoneOnResection",
        "BRCA1_Mutation",
        "BRCA2_Mutation",
        "Homologous_Recombination_Repair_Deficiency_Test_HRD",
        "Platinum_Sensitive_Resistant_Refractory",
        # --- Group 8: Prostate Specific Data (PSA & ADT) ---
        "Psa",
        "PSA_At_Diagnosis_ugPerLitre",
        "PSA_Nadir_ugPerLitre_DetectionLimitsVary",
        "PSA_At_Last_FollowUp_ugPerLitre_DetectionLimitsVary",
        "Gleason_Score",
        "ISUP_Grade",
        "Androgen_Deprivation_Therapy",
        "Castrate_Resistant_Prostate_Cancer",
        "Biochemical_Relapse_During_ADT",
        "Clinical_Relapse_During_ADT",
        "Details_Of_Biochemical_Recurrence",
        "TTE_CRPC_Biochemical_From_ADT_StartDate",
        "TTE_CRPC_Biochemical_From_DateDiagnosis",
        "TTE_CRPC_Clinical_From_ADT_StartDate",
        "TTE_CRPC_Clinical_From_DateDiagnosis",
        # --- Group 9: Brain / Glioma / Neuroendocrine Markers ---
        "P16_PosNeg",
        "P53_Status_IHC",
        "Inferred_P53_Mutation_Status",
        "CDKN2A_Loss",
        "MGMT_Promoter_Methylation",
        "IHC_IDH1_Mutation",
        "NGS_IDH1_Mutation",
        "IDH2_Mutation",
        "Histone_H3F3A_Mutation",
        "Histone_HIST1H3B_Mutation",
        "Chromosome_1P36_Loss",
        "Chromosome_19Q13_Loss",
        "TERT_Promoter_Mutation",
        "Cd56",
        "Chromogranina",
        "Synaptophysin",
        # --- Group 10: Immuno-oncology (PD-L1) Markers ---
        "PDL1_Antibody",
        "PDL1_PosNeg",
        "Percent_Cells_Pos",
        "PDL1_Combined_Pos_Score_CPS_DAKO_PDL1_IHC_Pharmdx_Test",
        "PDL1_Tumour_Pos_Score_TPS_DAKO_PDL1_28_8_Pharmdx_Test",
        "PDL1_Proportion_Tumour_Infiltrating_ImmuneCells_Ventana_PDL1_SP142_Assay",
        "PDL1_Score_For_Atesolisumab",
        "PDL1_Score_For_Nivolumab",
        "PDL1_Score_For_Pembrolisumab",
        "Secondary_PDL1_Antibody",
        "Secondary_PDL1_PosNeg",
        "Secondary_PDL1_Percent_Cells_Pos",
        "Third_PDL1_Antibody",
        "Third_PDL1_PosNeg",
        "Third_PDL1_Percent_Cells_Pos",
        # --- Group 11: Longitudinal Treatment & Outcomes ---
        "Macroscopic_Surgical_Tumour_Debulking",
        "Last_Treatment_Prior_To_Spec1_Surgery",
        "Last_Treatment_Prior_To_Spec2_Surgery",
        "Last_Treatment_Prior_To_Spec3_Surgery",
        "Post_Treatment_Spec1_Anonymised_ID",
        "Post_Treatment_Spec1_PathNumber",
        "Post_Treatment_Spec2_Anonymised_ID",
        "Post_Treatment_Spec3_Anonymised_ID",
        "Notes",
    ]

    existing_columns = [col for col in supp_global_col_order if col in df.columns]
    remaining_columns = [col for col in df.columns if col not in supp_global_col_order]

    # Restructures the clean, filtered input data
    ordered_df = df[existing_columns + remaining_columns]
    return ordered_df


# define a function to fetch the unique column names from a set of tables (for populating the GUI list box)
def fetch_col_names(selected_tables):
    """
    This functions gets the unique column names from the tables selected in the database querying user interface.
    """
    # create the empty list of columns
    unique_cols_list = []
    # Loop through the selected tables and query the column names
    for tbl in selected_tables:
        query = f"SELECT name FROM sys.columns WHERE object_id = OBJECT_ID('{tbl}')"
        # run the query
        columns_list = pd.read_sql(query, cnxn)
        unique_cols_list.append(columns_list["name"])

    # flatten the list of unique columns
    flat_unique_cols_list = [
        col for nested_list in unique_cols_list for col in nested_list
    ]

    # convert the list to a set
    flat_unique_cols_list = list(set(flat_unique_cols_list))

    return flat_unique_cols_list


# define a function to replace the Boolean column values with "yes" and "no"
def apply_bool_replacements(df):
    """This function applies replacements to Boolean columns (0s and 1s) in the query output - it replaces the 0 with a No and the 1 with
    a Yes."""
    df = df.replace(False, "No")
    df = df.replace(True, "Yes")
    return df


# ----------------------------------------------------------------------------------------------------------

# query_data(

#     tables="""
#     crc.crc1_FollowUp_Anon
#     crc
#     """,
#     columns="*"
# )
