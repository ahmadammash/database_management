# Import the required libraries
import pandas as pd


# Import the required functions
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

project_root, cnxn, cursor, engine, current_file = setup_cnxn()

chu_nice_consolidated = pd.read_sql(
    "SELECT * FROM chu_nice.chu_nice_consolidated", engine
)

col_list = [col for col in chu_nice_consolidated.columns]

col_order = [

# Identification
'Sample_ID',
'Biobank_ID',
'Pathology_ID',

# Demographics and Lifestyle
'Gender',
'Age',
'Age_At_DX',
'Ethnicity',

# Sample and Tumour Information
'Collection_Date',
'Date_Of_Surgery',
'Histological_Diagnosis',
'pTNM_Stage',
'Stage',
'Surgery_Site',
'Specimen_1_Collection_Date',
'Specimen1_Type',
'Diagnosis',
'Date_Of_Clinical_DX',
'Date_Of_Histological_Diagnosis',
'Site_Of_Diagnosis',

# Treatment Information
'Date_Of_Met_Disease',
'Met_Disease',
'Neoadjuvant_Treatment_Type',
'Neoadjuvant_Treatment_StartDate',
'Neoadjuvant_Treatment_EndDate',
'Date_Of_Relapse',
'Other_Relapse_1',
'Other_Relapse_1_Sites',
'Other_Relapse_2',
'Other_Relapse_2_Sites',
'Other_Relapse_3',
'Other_Relapse_3_Sites',
'Other_Relapse_4',
'Other_Relapse_4_Sites',
'Radiation_1_Type',
'Radiation_1_Start_Date',
'Radiation_1_End_Date',
'Radiation_2_Type',
'Radiation_2_Start_Date',
'Radiation_2_End_Date',
'Adjuvant_ChemoTx_1',
'Adjuvant_ChemoTx_1_Start_Date',
'Adjuvant_ChemoTx_1_End_Date',
'Immuno_Tx',
'Immuno_Tx_Start_Date',
'Immuno_Tx_End_Date',
'Metastatic_Chemo_Targeted_Tx_1',
'Metastatic_Chemo_Targeted_Tx_1_Start_Date',
'Metastatic_Chemo_Targeted_Tx_1_End_Date',
'Metastatic_Chemo_Targeted_Tx_2',
'Metastatic_Chemo_Targeted_Tx_2_Start_Date',
'Metastatic_Chemo_Targeted_Tx_2_End_Date',
'Immuno_Met_Tx',
'Immuno_Met_Tx_Start_Date',
'Immuno_Met_Tx_End_Date',
'Other_Therapy',
'Other_Therapy_Start_Date',
'Other_Therapy_End_Date',

# Molecular data
'KRAS_Mutation',
'EGFR_Mutation',
'PDL1',
'PDL1_Percent',


# Census
'Current_Status',
'Death',
'Date_Last_Seen',
'Date_Last_FollowUp',
'Death_Date',
'Cause_Of_Death',
'Concurrent_Medical_Conditions',
'OS_Months',
'Comments_OS_PFS',
'PFS_Months',
'Responder_NonResponder',
'BestResponse_ResponseNonResponse'
 ]

chu_nice_consolidated = chu_nice_consolidated[[col for col in col_order if col in chu_nice_consolidated.columns]]

chu_nice_consolidated.to_sql(
    "chu_nice_consolidated", engine, schema='chu_nice', if_exists='replace', index=False
)