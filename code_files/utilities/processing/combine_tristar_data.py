# Import the required libraries
import sys

import pandas as pd

import os
import sys

from pathlib import Path

# Import functions from the project directory
from code_files.utilities.io.setup_db_cnxn import setup_cnxn
from code_files.utilities.io.database import create_gcloud_engine

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()
env_path = project_root / "code_files" / "utilities" / "io" / ".env"
engine = create_gcloud_engine(env_path)

misplaced_headers = [s.strip() for s in 
                     """
03/12/2025
8242-
HGUGM-8181-
mm

""".splitlines() if s.strip()]

def combine_TS_center_data(schema_name):
    all_tables_query = f"""
    SELECT table_name, table_schema
    FROM information_schema.tables
    WHERE table_schema = '{schema_name}'
    """
    all_tables = pd.read_sql(all_tables_query, engine)
    all_tables = all_tables.loc[~all_tables.table_name.str.contains('map')]

    dfs_list = []
    id_cols = [
        'Sample_ID',
        'SampleID',
        'Name',
        'Subject_ID',
        'Sample ID',
        'Block_Number',
        'DonorID',
        'Project_ID',
        'NO',
        'NO.',
    ]

    for idx, row in all_tables.iterrows():
        tbl = row['table_name']
        schema = row['table_schema']
        df_query = f'SELECT * FROM "{schema}"."{tbl}"'
        df = pd.read_sql(df_query, engine)

        # find out if the TMA has wrong column names
        if any(col in df.columns for col in misplaced_headers):
            print(f"TMA {tbl} has incorrect headers")


        # Find the first matching ID column in df
        found_id_col = next((col for col in id_cols if col in df.columns), None)

        if found_id_col is None:
            print(f'No ID column found in table {schema}.{tbl}. Skipping this table.')
            continue

        # Standardize column name and type
        df = df.rename(columns={found_id_col: 'Sample_ID'})
        df['Sample_ID'] = df['Sample_ID'].astype(str).str.strip()
        df['TMA_Number'] = tbl

        dfs_list.append(df)

    if not dfs_list:
        print('No valid tables with identifier columns were found.')
        return all_tables, pd.DataFrame()

    df_combined = pd.concat(dfs_list, ignore_index=True)

    # Group by the standardized column name
    harmonised_df_combined = df_combined.groupby(
        'Sample_ID', as_index=False
    ).first()

    return all_tables, harmonised_df_combined

#all_hgugm_tables, hgugm_combined = combine_TS_center_data("hgugm")
#all_iom_tables, iom_combined = combine_TS_center_data("iom_aviano")
# all_tmas, tmas_combined = combine_TS_center_data("tma")

# In[]:
ids_list = [s.strip() for s in """
TS-18 - 0380
TS-18 - 0383
TS-18 - 1965
TS-18 - 1970
TS-18 - 1993
TS-18 - 0281
TS-18 - 0361
""".splitlines() if s.strip()]

#In[]:
iom_tbls, iom_combined = combine_TS_center_data('iom_aviano')
hgugm_tbls, hgugm_combined = combine_TS_center_data('hgugm')

chu_nice_query = """
SELECT * FROM chu_nice.chu_nice_consolidated
"""
chu_nice_df = pd.read_sql(chu_nice_query, engine)

hannover_query = """
SELECT * FROM hannover.combined_hannover
"""
hannover_combined = pd.read_sql(hannover_query, engine)

nhsb_query = """
SELECT * FROM nhsb.nhsb_tri_samples
WHERE "Sample_ID" IN (
'TS-17 - 0900',
'TS-17 - 0918',
'TS-17 - 0937',
'TS-16 - 0019',
'TS-16 - 0534',
'TS-17 - 1100',
'TS-17 - 1085'
)
"""
nhsb_df = pd.read_sql(nhsb_query, engine)

tristar_dfs  = [iom_combined, hgugm_combined, hannover_combined, chu_nice_df, nhsb_df]

tristar_dfs_combined = pd.concat(tristar_dfs, ignore_index=True).groupby('Sample_ID', as_index=False).first()

# In[]:
canonical_ordered_columns = [
    # --- 1. Identifiers & Source Tracking ---
    "Patient_ID",
    "Sample_ID",
    "Original_ID",
    "Biobank_ID",
    "Pathology_ID",
    "Study_ID",
    "Project_Code",
    "Accession_Number",
    "Current_Label",
    "Unique_ID_Series",
    "Core_ID",
    "Core_Reference",
    "TMA_Number",
    "Spot_Number",
    "Block_ID",
    "Block_Number",
    "Block_A",
    "Block_B",
    "Block_C",
    "Block_D",
    "Block_E",
    "Block_F",
    "Block_G",
    "Block_H",
    "Met_Block_1",
    "Met_Block_2",
    "Met_Block_3",
    "Met_Block_4",
    # --- 2. Demographics & Lifestyle ---
    "Gender",
    "Age",
    "Age_At_DX",
    "Age_At_Specimen_Collection",
    "Race",
    "Ethnicity",
    "Smoking_History",
    "Alcohol_History",
    "Concurrent_Medical_Conditions",
    "Medications",
    # --- 3. Primary Diagnosis & Baseline Clinical ---
    "Diagnosis",
    "Diagnosis_Subject",
    "Diagnosis_Specimen",
    "Clinical_Diagnosis",
    "Pathological_Diagnosis",
    "Histological_Diagnosis",
    "Histopathologic_Subtype",
    "New_Diagnosis",
    "Date_Of_Diagnosis",
    "Date_Of_Clinical_DX",
    "Date_Of_Histological_Diagnosis",
    "Month_Year_Initial_Diagnosis_PrimaryCancer",
    "Primary_Site",
    "Surgery_Site",
    "Organ",
    "Originating_Site_Specimen",
    "TumourSite_Location",
    "Type",
    "Sub_Type",
    "CancerType_Indication",
    "General_Cancer_Subtypes",
    "Symptoms",
    "International_Prognostic_Index",
    "PET_At_Diagnosis",
    # --- 4. Histopathology, Grade & Staging ---
    "Tumor_Grade",
    "Histological_Grade",
    "Nuclear_Grade",
    "Histological_Tumour_Type",
    "Histological_Tumour_Type_Lauren",
    "Histological_Tumour_Type_WHO",
    "Gleason_Score",
    "TNM",
    "pTNM_Stage",
    "Stage",
    "Stage_At_Dx",
    "T",
    "N",
    "M",
    "AJCC_Stage",
    "AJCC_Stage_Ann_Arbor",
    "AJCC_TNM_Clinical",
    "AJCC_TNM_Pathological",
    "AJCC_pT",
    "AJCC_pN",
    "Regional_LymphNodes_Status",
    "Lymph_Nodes_Num_Assessed",
    "Lymph_Nodes_Num_Positive",
    "Lymph_Nodes_Localization",
    "Lymph_Nodes_Size",
    "Micro_Invasion_Other_Organs",
    "TumourInvasion_Extent",
    "Surgical_Margin_Status",
    "Gross_Description",
    "HistologicFeatures_MicroscopicDescription",
    # --- 5. Biospecimen, Tissue QC & Sample Attributes ---
    "Specimen_Type",
    "Sample_Type",
    "Specimen_Nature",
    "Tissue_Type",
    "Collection_Matrix",
    "Collection_Procedure",
    "Specimen_Collection_Method",
    "Specimen_Source",
    "Specimen_Source_Detail",
    "Specimen_Biopsy_Method",
    "Surgical_Specimen",
    "Date_Of_Surgery",
    "Date_Of_Collection",
    "Year_Of_Collection",
    "Month_Year_Surgery_SampleCollection",
    "Tumour_Size",
    "Tumour_Weight",
    "Volume",
    "Volume_UOM",
    "Paired_Normal",
    "Paired_Tumor",
    "Paired_Metastasis",
    "QC_Results_TumourTissue_Perc",
    "QC_Results_Necrotic_Tissue_Perc",
    "QC_Results_Other_Tissue_Perc",
    "QC_Results_Tumour_Size",
    # --- 6. Biomarkers, IHC & Molecular Profiling ---
    "ER_Status",
    "PR_Status",
    "HER2_Status",
    "HER2_IHC",
    "KI67_Percent",
    "PDL1_Status",
    "PDL1_Percent",
    "KRAS_Status",
    "KRAS_Mutation",
    "EGFR_Mutation",
    "Mutation_Data",
    "Molecular_Data",
    "RNAscope",
    "CD20",
    "CD79A",
    "BCL2",
    "CD10",
    "BCL6",
    "MUM1",
    "Other_Ancillary_Findings",
    # --- 7. Systemic Therapies (Neoadjuvant, Adjuvant, Lines) ---
    "Tx_Naive_Confirmation",
    "Previous_Treatment",
    "Previous_Treatment_ClinicalResponse",
    "Lines_Of_Therapy",
    "Treatment_Schedule",
    "N_Cycles",
    "Drug",
    # Neoadjuvant:
    "Neoadjuvant_Treatment_Type",
    "Neoadjuvant_Treatment_StartDate",
    "Neoadjuvant_Treatment_EndDate",
    "Neoadjuvant_Chemotherapy_Agent",
    # Adjuvant:
    "Adjuvant_ChemoTx_1",
    "Adjuvant_ChemoTx_1_Start_Date",
    "Adjuvant_ChemoTx_1_End_Date",
    "Adjuvant_ChemoTx_2",
    "Adjuvant_ChemoTx_2_Start_Date",
    "Adjuvant_ChemoTx_2_End_Date",
    "Adjuvant_ChemoTx_3",
    "Adjuvant_ChemoTx_3_Start_Date",
    "Adjuvant_ChemoTx_3_End_Date",
    "Adjuvant_ChemoTx_4",
    "Adjuvant_ChemoTx_4_Start_Date",
    "Adjuvant_ChemoTx_4_End_Date",
    # Lines of Therapy:
    "1st_Line_Chemotherapy",
    "1st_Line_Chemotherapy_Agent",
    "2nd_Line_Chemotherapy",
    "2nd_Line_Chemotherapy_Agent",
    "3rd_Line_Chemotherapy",
    "3rd_Line_Chemotherapy_Agent",
    # Metastatic & Targeted:
    "Metastatic_Chemo_Targeted_Tx_1",
    "Metastatic_Chemo_Targeted_Tx_1_Start_Date",
    "Metastatic_Chemo_Targeted_Tx_1_End_Date",
    "Metastatic_Chemo_Targeted_Tx_2",
    "Metastatic_Chemo_Targeted_Tx_2_Start_Date",
    "Metastatic_Chemo_Targeted_Tx_2_End_Date",
    "Metastatic_Chemo_Targeted_Tx_3",
    "Metastatic_Chemo_Targeted_Tx_3_Start_Date",
    "Metastatic_Chemo_Targeted_Tx_3_End_Date",
    # Biologics & Immunotherapy:
    "Rituximab",
    "Maintenance_Therapy_With_Rituximab",
    "Immuno_Tx",
    "Immuno_Tx_Start_Date",
    "Immuno_Tx_End_Date",
    "Immuno_Met_Tx",
    "Immuno_Met_Tx_Start_Date",
    "Immuno_Met_Tx_End_Date",
    # Palliative & Other:
    "Palliative_Chemotherapy",
    "Palliative_Chemotherapy_Agent",
    "Sensitizing_Chemotherapy",
    "Sensitizing_Chemotherapy_Agent",
    "Other_Therapy",
    "Other_Therapy_Start_Date",
    "Other_Therapy_End_Date",
    # Treatment Responses:
    "Response_To_Therapy",
    "Clinical_Response",
    "Best_Response",
    "Responder_Status",
    "Response_Therapy_1",
    "Response_Therapy_2",
    "Response_Therapy_3",
    "Response_Therapy_4",
    # --- 8. Radiation & Surgery Procedures ---
    "Radiotherapy",
    "Radiation_1_Type",
    "Radiation_1_Start_Date",
    "Radiation_1_End_Date",
    "Radiation_2_Type",
    "Radiation_2_Start_Date",
    "Radiation_2_End_Date",
    "Metastatic_Radiation_Targeted_Tx_1",
    "Metastatic_Radiation_Targeted_Tx_1_Start_Date",
    "Metastatic_Radiation_Targeted_Tx_1_End_Date",
    "Transplantation",
    # --- 9. Recurrence, Metastasis & Disease Progression ---
    "Date_Of_Relapse",
    "Date_Of_Disease_Recurrence",
    "Recurrent_Disease_Types",
    "Recurrent_Disease_Locations",
    "Other_Relapse_1",
    "Other_Relapse_1_Sites",
    "Other_Relapse_2",
    "Other_Relapse_2_Sites",
    "Other_Relapse_3",
    "Other_Relapse_3_Sites",
    "Other_Relapse_4",
    "Other_Relapse_4_Sites",
    "Date_Of_Met_Disease",
    "Met_Disease",
    "Mets_Site",
    "Metastatic_Site_Distant",
    "Metastatic_Site_Distant_Status",
    "Nodal_Mets",
    "Extranodal_Localization",
    "Bone_Marrow",
    "Bulky",
    # --- 10. Follow-Up, Survival & Vital Status ---
    "Current_Status",
    "Cancer_Status",
    "Viral_Status",
    "Death",
    "Date_Of_Death",
    "Cause_Of_Death",
    "Death_Due_To_Lymphoma",
    "Date_Last_Follow_Up",
    "Date_Last_Seen",
    "Date_Of_Last_Contact",
    "Follow_Up_Months",
    "Follow_Up_Result",
    "OS_Months",
    "DFS_Months",
    "PFS_Months",
    "Comments_OS_PFS",
    # --- 11. Logistics, Tracking & Shipping ---
    "TRK_Shipment",
    "TRK_Incoming",
    "Position_In_Box",
    "Line_Shipping_List",
    "Line_Sample_Manifest",
    "Purchase_Order",
    "Format",
    "Comments",
]


column_mapping = {
    # --- Demographics ---
    "AgeAtDX": "Age_At_DX",
    "Smoker": "Smoking_History",
    "Concurrent_Diseases": "Concurrent_Medical_Conditions",
    # --- Staging & Subtyping ---
    "Stage At Dx\xa0": "Stage_At_Dx",
    "StageAtDX": "Stage_At_Dx",
    "Histopathologic_Subtype": "Histological_Tumour_Type",
    "Isto.": "Histological_Diagnosis",
    "Anno.": "Year_Of_Collection",
    "Gleason": "Gleason_Score",
    "Histological_Grade_Nottingham_Histologic_Score": "Histological_Grade",
    "AJCC_Stage_Or_Other_Applicable_System": "AJCC_Stage",
    "AJCC_TNM_Or_Other_Applicable_System": "TNM",
    "DescriptionTumour_MacroscopicTumourType_GrossDescription": "Gross_Description",
    # --- Indications & Diagnostic Labels ---
    "CancerType_Indication_Diagnosis": "CancerType_Indication",
    "CancerType_Indication_C50_Malignant_Neoplasm_Of_Breast": "CancerType_Indication",
    "General_Subtypes_Malignant_Neoplasm_Breast_ICD_Code": "General_Cancer_Subtypes",
    "General_Subtypes_Malignant_Neoplasm_Gallbladder_ICD_Code": "General_Cancer_Subtypes",
    "General_Subtypes_Malignant_Neoplasm_Pancreas_ICD_Code": "General_Cancer_Subtypes",
    "General_Subtypes_Malignant_Neoplasm_MaleGenitalOrgans_ICD_Code": "General_Cancer_Subtypes",
    "General_Subtypes_Neoplasms_Primary_Lymphoid_Hematopoietic_And_R": "General_Cancer_Subtypes",
    "General_Subtypes_Malignant_Neoplasm_MesothelialAndSoftTissue_IC": "General_Cancer_Subtypes",
    "General_Subtypes_Malignant_Neoplasm_Liver_IntrahepticBileDucts_": "General_Cancer_Subtypes",
    # --- Specimen Attributes ---
    "Sample_Type": "Specimen_Type",
    "Specimen_1_Type": "Specimen_Type",
    "Specimen1_Type": "Specimen_Type",
    "Specimen_1_Collection_Date": "Date_Of_Collection",
    "Collection_Date": "Date_Of_Collection",
    # --- Biomarkers & IHC ---
    "ER": "ER_Status",
    "HormoneReceptorStatus_EstrogenReceptor": "ER_Status",
    "PR": "PR_Status",
    "Progestrone_Receptor": "PR_Status",
    "HER2": "HER2_Status",
    "H2": "HER2_Status",
    "HER2_Status_HumanEpidermalGrowthFactorReceptor2": "HER2_Status",
    "HER2_Status_Other_Ancillary_Findings": "HER2_Status",
    "KI67": "KI67_Percent",
    "Immunohistochemistry_KI67_Expression": "KI67_Percent",
    "Immunohistochemical_Markers": "Immunohistochemistry",
    "PDL1": "PDL1_Status",
    "PDL1_Percent": "PDL1_Percent",
    "Status_KRAS": "KRAS_Status",
    "Mutation_KRAS": "KRAS_Mutation",
    # --- Systemic & Adjuvant Therapies ---
    "Any_Neoadjuvant_Therapy": "Neoadjuvant_Treatment_Type",
    "Neoadjuvant_Chemotherapy": "Neoadjuvant_Treatment_Type",
    "Neoadjuvant_Therapy_Agent": "Neoadjuvant_Chemotherapy_Agent",
    "Adjuvant_ChemoTx_Regimen_1": "Adjuvant_ChemoTx_1",
    "Adjuvant_ChemoTx_Regimen_2": "Adjuvant_ChemoTx_2",
    "Radiation 1: Type": "Radiation_1_Type",
    "Metastatic_Radio_Targeted_Tx_1": "Metastatic_Radiation_Targeted_Tx_1",
    "Metastatic_Radio_Targeted_Tx_1_Start_Date": "Metastatic_Radiation_Targeted_Tx_1_Start_Date",
    "Metastatic_Radio_Targeted_Tx_1_End_Date": "Metastatic_Radiation_Targeted_Tx_1_End_Date",
    # --- Treatment Response Measures ---
    "Responder_NonResponder": "Responder_Status",
    "BestResponse_ResponseNonResponse": "Best_Response",
    "Response_To_Second_Line": "Response_Therapy_2",
    # --- Recurrence & Progression ---
    "Date_Of_Met_Disease_1": "Date_Of_Met_Disease",
    "Met_Disease_1": "Met_Disease",
    "Date_Of_Met_Disease_2": "Date_Of_Met_Disease",
    "Met_Disease_2": "Met_Disease",
    "Metastatic_Site": "Mets_Site",
    "Metastatic_Or_Primary": "Primary_Tumour_Site",
    # --- Survival, Mortality & Follow-up ---
    "\xa0Exitus": "Death",
    "Death_Date": "Date_Of_Death",
    "OS": "OS_Months",
    "DFS": "DFS_Months",
    "Survival": "Current_Status",
    "Census": "Current_Status",
    "Date_Last_FollowUp": "Date_Last_Follow_Up",
    # --- Shipping & Metadata Tracking ---
    "TRK_Spedizione": "TRK_Shipment",
    "Line_Sample_List": "Line_Shipping_List",
    "ID_Format": "Format",
}


def clean_clinical_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Strips hidden whitespace, coalesces duplicated columns,

    and orders fields by domain.
    """
    # Clean whitespace in column names
    df.columns = [c.strip() for c in df.columns]

    # Drop standard pandas indexing leftovers
    df = df.drop(columns=["level_0", "index"], errors="ignore")

    # Combine columns that share the same target canonical name
    for old_col, new_col in column_mapping.items():
        if old_col in df.columns:
            if new_col in df.columns and old_col != new_col:
                # Fill missing target values with the duplicate column's data
                df[new_col] = df[new_col].combine_first(df[old_col])
                df = df.drop(columns=[old_col])
            else:
                df = df.rename(columns={old_col: new_col})

    # Order dataframe by canonical list (keeping remaining unmapped at the end)
    existing_canonical = [c for c in canonical_ordered_columns if c in df.columns]
    unmapped_cols = [c for c in df.columns if c not in existing_canonical]

    return df[existing_canonical + unmapped_cols]

tristar_dfs_combined = clean_clinical_dataframe(tristar_dfs_combined)

# In[]:
ids_list = [s.strip().replace(' ', '') for s in 
            """
            TS-21-0002
            TS-21-0025
            TS-18-0090
            TS-18-0087
            IOM-708
            IOM-707
            IOM-692
            IOM-690
            IOM-687
            982-033
            982-040
            982-054
            TS-17 - 0900
            TS-17 - 0918
            TS-17 - 0937
            TS-16 - 0019
            TS-16 - 0534
            TS-17 - 1100
            TS-17 - 1085
            TS-18-1587
            TS-18-1593
            TS-18-1594
            TS-18-1678
            TS-18-1679
            H-BLC-003
            H-BLC-037
            H-BLC-045
            H-BLC-047
            H-BLC-058
            5271T1
            5905T2
            5938T1
            TS-18 - 2176
            TS-19 - 0058
            TS-17 - 1950
            TS-17 - 1934
            5902T2
            5901T2
            TS-17 - 1989
            7406T2
            """.splitlines() if s.strip()
            
            ]
iom_combined = clean_clinical_dataframe(iom_combined)
nhsb_df = clean_clinical_dataframe(nhsb_df)
chu_nice_df = clean_clinical_dataframe(chu_nice_df)
hannover_combined = clean_clinical_dataframe(hannover_combined)
hgugm_combined = clean_clinical_dataframe(hgugm_combined)
