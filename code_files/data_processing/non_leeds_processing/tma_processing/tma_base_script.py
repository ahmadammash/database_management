######################################################################################################
# General Setup and Imports
######################################################################################################
import pandas as pd
from openpyxl import load_workbook

######################################################################################################
# Import the TMA Workbook
######################################################################################################

# define the file path to the TMA workbook
filepath = "C:/Users/medaamm/OneDrive - University of Leeds/UOL Data Manager/database_management/data_files/non_leeds_data/tma_data/DATASHEET_TriStar.xlsx"

# define the TMA workbook
tma_wb = load_workbook(filename=filepath)

# Import the required functions
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

######################################################################################################
# Define the TMA Sheets
######################################################################################################

# define the workbook sheets (attribute, no parentheses)
tma_all_sheets = tma_wb.sheetnames

######################################################################################################
# Define the TMA data table ranges and extract the column names
######################################################################################################

# define the list of all columns across TMA sheets
all_tma_data_columns = []

# get all table ranges in the workbook
tma_table_ranges = {}

# Loop through sheets and record table metadata (sheet name & cell range)
for sheet_name in tma_all_sheets:
    sheet = tma_wb[sheet_name]
    for table_name in sheet.tables:
        table_obj = sheet.tables[table_name]
        tma_table_ranges[table_name] = {
            "sheet": sheet_name,
            "range": table_obj.ref,
        }

# Extract matching tables into Pandas DataFrames
tma_dfs = {}

for tbl_name, info in tma_table_ranges.items():
    if "data_" in tbl_name:
        sheet = tma_wb[info["sheet"]]

        # Get cell values from the worksheet range
        data = [[cell.value for cell in row] for row in sheet[info["range"]]]

        # Convert to DataFrame (Row 0 = headers, Rows 1+ = data)
        df = pd.DataFrame(data[1:], columns=data[0])

        # Clean header strings
        df.columns = df.columns.astype(str).str.strip()

        if "" in df.columns:
            print(tbl_name)

        # Append the columns to the list of columns
        all_tma_data_columns.append(df.columns)

        tma_dfs[tbl_name] = df


# flatten the columns list and remove duplicates
all_tma_data_columns = pd.Index(
    [col for cols in all_tma_data_columns for col in cols]
).unique()

######################################################################################################
# Rename and standardise the columns using fuzzy matching
######################################################################################################

column_name_map = {}

# Existing Loops (Fixed missing f-strings and syntax bugs)
for i in range(2, 13):
    column_name_map[f"follow up {i} (+ 3m)"] = f"FollowUp_{i}_Plus3Months"

for i in range(1, 16):
    column_name_map[f"TTE Subsequent Tx{i} Start Date"] = (
        f"TTE_Subsequent_Tx{i}_StartDate"
    )
    column_name_map[f"TTE Subsequent Tx{i} End Date"] = f"TTE_Subsequent_Tx{i}_EndDate"
    column_name_map[f"TTE Date Recurrence/ Progression {i}"] = (
        f"TTE_DateRecurrenceProgression{i}"
    )
    column_name_map[f"TTE Adjuvant Tx{i} Start Date"] = f"TTE_Adjuvant_Tx{i}_StartDate"
    column_name_map[f"TTE Adjuvant Tx{i} End Date"] = f"TTE_Adjuvant_Tx{i}_EndDate"
    column_name_map[f"Sites Recurrence/ Progression {i}"] = (
        f"Sites_RecurrenceProgression{i}"
    )
    column_name_map[f"Adjuvant Tx{i}"] = f"AdjuvantTx{i}"
    column_name_map[f"Adjuvant Tx{i} Agent"] = f"AdjuvantTx{i}_Agent"
    column_name_map[f"Tumour Response Adjuvant Tx{i}"] = f"TumourResponse_AdjuvantTx{i}"
    column_name_map[f"Subsequent Tx{i}"] = f"SubsequentTx{i}"
    column_name_map[f"Subsequent Tx{i} Agent"] = f"SubsequentTx{i}_Agent"
    column_name_map[f"Tumour Response Subsequent Tx{i}"] = (
        f"TumourResponse_SubsequentTx{i}"
    )

for i in range(1, 11):
    column_name_map[f"TTE Neoadjuvant Tx{i} Start Date"] = (
        f"TTE_Neoadjuvant_Tx{i}_StartDate"
    )
    column_name_map[f"TTE Neoadjuvant Tx{i} End Date"] = (
        f"TTE_Neoadjuvant_Tx{i}_EndDate"
    )
    column_name_map[f"Neoadjuvant Tx{i}"] = f"NeoadjuvantTx{i}"
    column_name_map[f"Neoadjuvant Tx{i} Agent"] = f"NeoadjuvantTx{i}_Agent"
    column_name_map[f"Tumour Response Neoadjuvant Tx{i}"] = (
        f"TumourResponse_NeoadjuvantTx{i}"
    )

for i in range(1, 6):
    column_name_map[f"Anonymized ID Matched Met {i}"] = f"AnonymisedID_MatchedMet{i}"
    column_name_map[f"Anonymized ID Matched met {i}"] = f"AnonymisedID_MatchedMet{i}"
    column_name_map[f"Matched Met {i} Block Number"] = f"MatchedMet{i}_BlockNumber"
    column_name_map[f"Matched met {i} Block Number"] = f"MatchedMet{i}_BlockNumber"
    column_name_map[f"Matched Met {i}"] = f"MatchedMet{i}_Site"
    column_name_map[f"Matched Met {i} Surgery Date"] = f"MatchedMet{i}_SurgeryDate"
    column_name_map[f"Matched Met {i} Treatment Subgrouping"] = (
        f"MatchedMet{i}_TreatmentSubgrouping"
    )
    column_name_map[f"TTE Matched Met {i} Surgery Date"] = (
        f"TTE_MatchedMet{i}_SurgeryDate"
    )
    column_name_map[f"Neoadjuvant treatment {i}"] = f"NeoadjuvantTx{i}"
    column_name_map[f"Matched met {i} Site"] = f"MatchedMet{i}_Site"
    column_name_map[f"Matched met {i} Surgery Date"] = f"MatchedMet{i}_SurgeryDate"
    column_name_map[f"Matched met {i} treatment subgrouping"] = (
        f"MatchedMet{i}_TreatmentSubgrouping"
    )

for i in range(1, 5):
    column_name_map[f"Adjuvant treatment {i}"] = f"AdjuvantTx{i}"

# Relapses / Relapse Sites
for i in range(1, 5):
    column_name_map[f"Other Relapse {i}"] = f"OtherRelapse_{i}"

# Radiation Cycles
for i in range(1, 3):
    column_name_map[f"Radiation {i}: Type"] = f"Radiation_{i}_Type"

# Subsequent treatments
for i in range(1, 4):
    column_name_map[f"Subsequent treatment {i}"] = f"SubsequentTx{i}"

# Sample Surgery/Description/Sites
for i in range(1, 3):
    column_name_map[f"Sample {i} site"] = f"Sample{i}_Site"
    column_name_map[f"Sample {i} description"] = f"Sample{i}_Description"
    column_name_map[f"Sample {i} surgery"] = f"Sample{i}_Surgery"

# Specimen Block Numbers
for i in range(1, 4):
    column_name_map[f"Specimen Block number {i}"] = f"SpecimenBlockNumber_{i}"
    column_name_map[f"Specimen Block Number {i}"] = f"SpecimenBlockNumber_{i}"


# Add unique/standalone columns to dictionary
column_name_map.update(
    {
        "Sample ID": "Sample_ID",
        "ID written on the  TL sample (NSCLC)": "Sample_ID",
        "Selected specimen Anonymised ID": "Sample_ID",
        "Name": "Sample_ID",
        "Ref#2": "Sample_ID",
        "SAMPLE ID#": "Sample_ID",
        "SAMPLE ID": "Sample_ID",
        "TMA Core Ref": "Core_ID",
        "SAMPLE ID#2": "Sample_ID_2",
        "Biopsy/resection3": "Biopsy_Resection",
        "Age at diagnosis": "AgeDX",
        "Tumour site": "SpecTumourSite",
        "Tumour type": "SpecTumourType",
        "Gleason score": "Gleason_Score",
        "Stage (either overall or pathological)": "StageAtDX",
        "ID Bl": "Sample_ID",
        "Age range": "Age_Range",
        "Clinical diagnosis": "Clinical_Diagnosis",
        "Tumor Grade": "Grade",
        "surgical specimen": "Surgical_Specimen",
        "MET (IHC)": "MET_IHC",
        "age range": "Age_Range",
        "ID #": "Sample_ID",
        "Date of Surgery": "SurgeryDate",
        "DATE of SURGERY": "SurgeryDate",
        "Status K-RAS": "KRAS_Mutation",
        "START TREAT 1": "Therapy_1_StartDate",
        "TREAT 1": "Therapy_1",
        "END TREAT 1": "Therapy_1_EndDate",
        "START TREAT 2": "Therapy_2_StartDate",
        "TREAT 2": "Therapy_2",
        "END TREAT 2": "Therapy_2_EndDate",
        "START TREAT 3": "Therapy_3_StartDate",
        "TREAT 3": "Therapy_3",
        "END TREAT 3": "Therapy_3_EndDate",
        "TIME from FIRST DIAGNOSIS (months)": "Time_From_Diagnosis_Months",
        "FOLLOW UP (months)": "FollowUp_Months",
        "EXCISION DATE OF COLON": "ExcisionDate_Colon",
        "EXCISON DATE OF LIVER MET": "ExcisionDate_LiverMet",
        "DATE OF DIAGNOSIS": "DateDX",
        "Status KRAS": "KRAS_Mutation",
        "Status EGFR (IHC)": "EGFR_Mutation_IHC",
        "follow up 1": "FollowUp_1",
        "AGE at diagnosis": "AgeDX",
        "Clinical Diagnosis": "Clinical_Diagnosis",
        "Pathology DIAGNOSIS": "Pathological_Diagnosis",
        "date of surgery": "SurgeryDate",
        "start theraphy 1": "Therapy_1_StartDate",
        "theraphy 1": "Therapy_1",
        "end therapy 1": "Therapy_1_EndDate",
        "start therapy 2": "Therapy_2_StartDate",
        "thrapy 2": "Therapy_2",
        "end therapy 2": "Therapy_2_EndDate",
        "start therapy 3": "Therapy_3_StartDate",
        "therapy 3": "Therapy_3",
        "end therapy 3": "Therapy_3_EndDate",
        "site of progression": "Progression_Site",
        "Last follow up": "Last_FollowUp",
        "HER2 ISH": "HER2_ISH",
        "HER2 score": "HER2_Score",
        "PD-L1(TPS)": "PDL1_TPS",
        "B7-H3 - 1+,2+,3+(%)": "B7H3_1Plus_2Plus_3Plus_Perc",
        "B7-H3 - 2+,3+(%)": "B7H3_2Plus_3Plus_Perc",
        "B7-H3 - 3+(%)": "B7H3_3Plus_Perc",
        "B7-H3 Neg(%)": "B7H3_NegPerc",
        "B7-H3 - H score": "B7H3_HScore",
        "B7-H3 - IHC Result": "B7H3_IHCResult",
        "B7-H4 - 1+,2+,3+(%)": "B7H4_1Plus_2Plus_3Plus_Perc",
        "B7-H4 - 2+,3+(%)": "B7H4_2Plus_3Plus_Perc",
        "B7-H4 - 3+(%)": "B7H4_3Plus_Perc",
        "B7-H4 - Neg(%)": "B7H4_NegPerc",
        "B7-H4 - H score": "B7H4_HScore",
        "B7-H4 - IHC Result": "B7H4_IHCResult",
        "Nectin-4 - 1+,2+,3+(%)": "Nectin4_1Plus_2Plus_3Plus_Perc",
        "Nectin-4 - 2+,3+(%)": "Nectin4_2Plus_3Plus_Perc",
        "Nectin-4 - 3+(%)": "Nectin4_3Plus_Perc",
        "Nectin-4 - Neg(%)": "Nectin4_NegPerc",
        "Nectin-4 - H score": "Nectin4_HScore",
        "Nectin-4 - IHC Result": "Nectin4_IHCResult",
        "TROP2 (1:1000) - 1+,2+,3+(%)": "TROP2_1_1000_1Plus_2Plus_3Plus_Perc",
        "TROP2 (1:1000) - 2+,3+(%)": "TROP2_1_1000_2Plus_3Plus_Perc",
        "TROP2 (1:1000) - 3+(%)": "TROP2_1_1000_3Plus_Perc",
        "TROP2 (1:1000) - Neg(%)": "TROP2_1_1000_NegPerc",
        "TROP2 (1:1000) - H score": "TROP2_1_1000_HScore",
        "TROP2 (1:1000) - IHC Result": "TROP2_1_1000_IHCResult",
        "TROP2 (1:3000) - 1+,2+,3+(%)21": "TROP2_1_3000_1Plus_2Plus_3Plus_Perc",
        "TROP2 (1:3000) - 2+,3+(%)": "TROP2_1_3000_2Plus_3Plus_Perc",
        "TROP2 (1:3000) - 3+(%)": "TROP2_1_3000_3Plus_Perc",
        "TROP2 (1:3000) - Neg(%)": "TROP2_1_3000_NegPerc",
        "TROP2 (1:3000) - H score": "TROP2_1_3000_HScore",
        "TROP2 (1:3000) - IHC Result": "TROP2_1_3000_IHCResult",
        "Number of Donor Blocks": "Num_Donor_Blocks",
        "Block #s": "Block_Number",
        "Marked areas": "Marked_Areas",
        "Number of Donor Blocks2": "Num_Donor_Blocks_2",
        "Block #s3": "Block_Number_1",
        "Cytological grading": "Cytological_Grading",
        "Gleason Score": "Gleason_Score",
        "III patron": "Patron_III",
        "Original ID": "Original_ID",
        "ID Ref": "Sample_ID",
        "MET SITE AT DIAGNOSIS": "Met_Site_At_Diagnosis",
        "Blast count percentage": "Blast_Count_Perc",
        "IPSS-R SCORE": "IPSS_R_Score",
        "Response to therapy": "Response_To_Therapy",
        "Follow up": "FollowUp",
        "AML Evolution": "AML_Evolution",
        "Donor Ref": "Sample_ID",
        "Pathological Diagnosis2": "Pathological_Diagnosis_2",
        "Clinical diagnosis5": "Clinical_Diagnosis_5",
        "smoker6": "Smoker_6",
        "previous treatment": "Previous_Treatment",
        "clinical response": "Clinical_Response",
        "1st line": "FirstLine_Treatment",
        "2º line": "SecondLine_Treatment",
        "3º line": "ThirdLine_Treatment",
        "Neoadjuvant Treatment12": "Neoadjuvant_Treatment_12",
        "Palliative QT": "Palliative_Chemotherapy",
        "sensitizing QT": "Sensitizing_Chemotherapy",
        "Treatment10": "Treatment_10",
        "Treatment11": "Treatment_11",
        "Treatment12": "Treatment_12",
        "Treatment13": "Treatment_13",
        "Treatment14": "Treatment_14",
        "Pathological Diagnose": "Pathological_Diagnosis",
        "Clinical diagnose": "Clinical_Diagnosis",
        "smoker5": "Smoker_5",
        "surgical date": "SurgeryDate",
        "quimioterapia": "Chemotherapy",
        "Treatment Group": "Treatment_Group",
        "Donor ref": "Sample_ID",
        "Donor Number": "Donor_Number",
        "Survival Status": "Survival_Status",
        "Date of diagnosis": "DateDX",
        "International Prognostic Index": "IPI_Score",
        "PET at diagnosis": "PET_At_Diagnosis",
        "Extranodal localization": "Extranodal_Localization",
        "Bone marrow": "Bone_Marrow_Involvement",
        "Treatment (schedula)": "Treatment_Schedule",
        "n. of cycles": "Num_Cycles",
        "maintainance therapy with rituximab ?": "Maintenance_Rituximab",
        "Date of relapse": "Relapse_Date",
        "last follow up": "Last_FollowUp",
        "Date of Death": "Death_Date",
        "Second treatment": "Second_Treatment",
        "death due to lymphoma?": "Death_Due_To_Lymphoma",
        "RESPONSE TO THERAPY": "Response_To_Therapy",
        "RESPONSE TO II line": "Response_To_SecondLine",
        "ID # Slides": "Slide_ID",
        "ID # Blocks": "Sample_ID",
        "HER2 (IHC)": "HER2_IHC",
        "Last Known Follow Up/Date of Death": "Last_Seen_Or_Death_Date",
        "M  (time 0)": "M_Stage_Time0",
        "CA9 - 1+ %": "CA9_1Plus_Perc",
        "CA9 - 2+ %": "CA9_2Plus_Perc",
        "CA9 - 3+ %": "CA9_3Plus_Perc",
        "CA9 - Neg %": "CA9_NegPerc",
        "CA9 - H score": "CA9_HScore",
        "CA9 - 1+,2+,3+ %": "CA9_1Plus_2Plus_3Plus_Perc",
        "CA9 - 2+,3+ %": "CA9_2Plus_3Plus_Perc",
        "CA9 - IHC result": "CA9_IHCResult",
        "IHC result": "IHC_Result",
        "Date of surgery": "SurgeryDate",
        "ID Project #": "Sample_ID",
        "NHL Subtype": "NHL_Subtype",
        "Date of Surgery                                 (mm/yy)": "SurgeryDate_MMYY",
        "Surgery & Site": "Surgery_And_Site",
        "Specimen 1 Collection Date  (mm/yy)": "Specimen_1_CollectionDate_MMYY",
        "Specimen 1 Type": "Specimen_1_Type",
        "Site of Diagnosis": "Diagnosis_Site",
        "Age at Diagnosis (years)": "AgeDX",
        "Date of DX (mm/yy)": "DateDX_MMYY",
        "Stage at DX": "StageAtDX",
        "Histological diagnosis": "Histological_Diagnosis",
        "Follow-up months": "FollowUp_Months",
        "Follow-up result": "FollowUp_Result",
        "Bl Primary": "Block_Primary",
        "Bl Met.": "Block_Metastatic",
        "Status (HPV)": "HPV_Status",
        "NUCLEAR GRADE": "Nuclear_Grade",
        "PD-L1 (TPS)": "PDL1_TPS",
        "PD-L1 (CPS)": "PDL1_CPS",
        "Date of Met Disease": "Metastatic_Disease_Date",
        "Met Disease": "Metastatic_Disease",
        "Radiation 1: Type": "Radiation_1_Type",
        "Radiation 2: Type": "Radiation_2_Type",
        "Adjuvant ChemoTx Regimen 1": "AdjuvantTx1_Regimen",
        "Adjuvant ChemoTx Regimen 2": "AdjuvantTx2_Regimen",
        "Immuno Tx": "Immunotherapy",
        "Metastatic Chemo/Targeted                     Tx Regimen 1": "Metastatic_Tx1_Regimen",
        "Metastatic Chemo/Targeted                     Tx Regimen 2": "Metastatic_Tx2_Regimen",
        "Immuno Met Tx": "Immunotherapy_Metastatic",
        "Other Therapy": "Other_Therapy",
        "Current Status": "CurrentStatus",
        "Date Last Seen": "Last_Seen_Date",
        "Death Date": "Death_Date",
        "Cause of Death": "CauseOfDeath",
        "Concurrent Medical Conditions": "ConcurrentMedicalConditions",
        "Sample ID#": "Sample_ID",
        "ID#": "Sample_ID",
        "Tissue type": "Tissue_Type",
        "Tumour sub-type or grade / differentiation": "Tumour_Subtype_Grade",
        "TS- number": "Sample_ID",
        "Primary lab number (block / slide)": "Primary_Lab_Number",
        "Core ID": "Core_ID",
        "TS- number2": "Sample_ID_2",
        "Secondary lab number (block / slide)": "Secondary_Lab_Number",
        "Tumour sub-type / Immunohistochemistry": "Tumour_Subtype_IHC",
        "Secondary cancer status": "Secondary_Cancer_Status",
        "Tumour site3": "SpecTumourSite_3",
        "Tissue type4": "Tissue_Type_4",
        "Tumour type5": "SpecTumourType_5",
        "BLOCK ID": "Sample_ID",
        "KRAS Status": "KRAS_Mutation",
        "Mutation or WT": "Mutation_Or_WT",
        "Tristar ID #": "Sample_ID",
        "Blocks and Slide ID": "Block_And_Slide_ID",
        "NEW ID": "Sample_ID",
        "PD-L1 (%)": "PDL1_Perc",
        "Her2 score": "HER2_Score",
        "new ID": "Sample_ID",
        "old ID": "Old_ID",
        "Lauren classification (Diffuse)": "Lauren_Classification",
        "Last known follow up": "Last_FollowUp",
        "Additional Information": "Additional_Info",
        "Cause of death": "CauseOfDeath",
        "Lab ID": "Lab_ID",
        "Country of Origin": "GeographicOrigin",
        "Collection year": "Collection_Year",
        "treatment 1": "Therapy_1",
        "follow up treatment 1": "Therapy_1_FollowUp",
        "treatment 2": "Therapy_2",
        "follow up treatment 2": "Therapy_2_FollowUp",
        "last follow up date": "Last_FollowUp_Date",
        "OS (months)": "OverallSurvival_Months",
        "DFS (months)": "DiseaseFreeSurvival_Months",
        "MSI status": "MSI_Status",
        "ID block": "Sample_ID",
        "ID block2": "Sample_ID_2",
        "Biobank ID": "Biobank_ID",
        "Pathology ID": "Pathology_ID",
        "Original ID#": "Original_ID",
        "Mutation Type": "Mutation_Type",
        "FFPE block ID": "Sample_ID",
        "Current StAlivetus": "CurrentStatus",
        "Age (years)": "AgeDX",
        "Survival (months)": "OverallSurvival_Months",
        "Reference ID": "Reference_ID",
        "Tissue ID": "Sample_ID",
        "Patient ID": "DonorID",
        "PSA data": "PSA_Data",
        "Follow up data": "FollowUp_Data",
        "Core No.": "Core_ID",
        "MONTH/YEAR OF INITIAL DIAGNOSIS OF PRIMARY CANCER": "Initial_Diagnosis_MonthYear",
        "MONTH/YEAR OF SURGERY/SAMPLE COLLECTION": "Sample_Collection_MonthYear",
        "TYPE OF CANCER/INDICATION": "Cancer_Type_Indication",
        "GENERAL SUBTYPES OF CANCER": "General_Cancer_Subtypes",
        "SPECIMEM NATURE": "Specimen_Nature",
        "COLLECTION PROCEDURE": "Collection_Procedure",
        "PRIMARY TUMOR SITE": "PrimTumourSite",
        "GROSS DESCRIPTION": "Gross_Description",
        "SIZE OF TUMOR mm": "Tumor_Size_MM",
        "HISTOLOGICAL TUMOR TYPE": "Histological_Tumour_Type",
        "HISTOLOGICAL GRADE": "Histological_Grade",
        "HISTOLOGIC FEATURES/MICROSCOPIC DESCRIPTION": "Microscopic_Description",
        "TUMOR INVASION": "Tumor_Invasion",
        "Extent of tumor": "Extent_Of_Tumor",
        "AJCC TNM (or other applicable)": "TNM_Stage",
        "AJCC Stage Grouping (or other applicable staging system)": "StageAtDX",
        "LYMPH NODES (Regional Lymph Notes)": "Lymph_Nodes_Status",
        "Metastatic Site (Distant)": "Metastatic_Site_Distant",
        "Lines of Therapy": "Lines_Of_Therapy",
        "Date of Disease Recurrence": "Recurrence_Date",
        "Recurrent Disease Type (s)": "Recurrent_Disease_Type",
        "Recurrent Disease Location (s)": "Recurrent_Disease_Location",
        "Cancer Status": "Cancer_Status",
        "Vital Status": "Vital_Status",
        "Date of Last Contact": "Last_Contact_Date",
        "Concurrent Diseases": "ConcurrentMedicalConditions",
        "Hormon Receptor Status": "Hormone_Receptor_Status",
        "HER2 Status": "HER2_Status",
        "PURCHASE-ORDER": "Purchase_Order",
        "PATIENT-ID": "DonorID",
        "MONTH-YEAR OF INITIAL DIAGNOSIS OF PRIMARY CANCER": "Initial_Diagnosis_MonthYear",
        "MONTH-YEAR OF SURGERY-SAMPLE COLLECTION": "Sample_Collection_MonthYear",
        "TYPES OF CANCER - INDICATION": "Cancer_Type_Indication",
        "GENERAL SUBTYPES OF  CANCER": "General_Cancer_Subtypes",
        "NEW PRIMARY OR RECURRENCE/": "New_Primary_Or_Recurrence",
        "Unique ID2": "Sample_ID",
        "Unique ID": "Sample_ID",
        "Pre/Post treatment": "Pre_Or_Post_Treatment",
        "Case Number": "Case_Number",
        "Collection Date": "Collection_Date",
        "Geographic Origin": "GeographicOrigin",
        "Surgery, Site and Side/Other info": "Surgery_Site_Side_Info",
        "Adjuvant/ Neo-Adjuvant Chemotx Regimen 1": "Adjuvant_Neoadjuvant_Regimen_1",
        "Adjuvant/Neo-Adjuvant Chemotx Regimen 2": "Adjuvant_Neoadjuvant_Regimen_2",
        "Age At Diagnosis": "AgeDX",
        "Ajcc Staging M": "M_Stage",
        "Ajcc Staging N": "N_Stage",
        "Ajcc Staging T": "T_Stage",
        "Alcohol Status": "Alcohol_Status",
        "Alcohol History": "Alcohol_History",
        "Smoking Status": "Smoking_Status",
        "Smoking History": "Smoking_History",
        "Cause Of Death": "CauseOfDeath",
        "Collection Date5": "Collection_Date_5",
        "Collection Time": "Collection_Time",
        "Contact Person": "Contact_Person",
        "Expiration Date": "Expiration_Date",
        "Gender7": "Gender_7",
        "Geographic Origin8": "GeographicOrigin_8",
        "Histopathologic Data/ Subtype": "Histopathologic_Data_Subtype",
        "Immuno-histopathologic Data": "Immunohistopathologic_Data",
        "Molecular Data": "Molecular_Data",
        "Other Therapy / Hormone Treatment": "Other_Therapy_Hormone_Treatment",
        "Sample Type": "Sample_Type",
        "Set Availability": "Set_Availability",
        "Specimen 1 Collection Date": "Specimen_1_CollectionDate",
        "Stage At Diagnosis": "StageAtDX",
        "Study Name": "Study_Name",
        "Surgery, Site And Side/Other info9": "Surgery_Site_Side_Info_9",
        "Immuno Tx (Regimen, type, duration)": "Immunotherapy_Details",
        "Cancer status (added)": "Cancer_Status",
        "Tumour Grade (added)": "Grade",
        "Estimated treatment dates (added)": "Estimated_Treatment_Dates",
        "Period between disease progression in months (added)": "PFS_Months",
        "Case Number10": "Case_Number_10",
        "Collection Date11": "Collection_Date_11",
        "Geographic Origin12": "GeographicOrigin_12",
        "Invasive edge/Tumor": "Invasive_Edge_Tumor",
        "RACE/ ETHNICITY": "Race_Ethnicity",
        "Final Tissue ID": "Sample_ID",
        "dapB comment": "DapB_Comment",
        "Grade or Gleason score as applicable": "Grade_Or_Gleason",
        "Type (Primary)": "Primary_Type",
        "Sub-Type or location": "Subtype_Or_Location",
        "TMA Core Number": "Core_ID",
        "Adjuvant/Neo-Adjuvant Chemotx Regimen 1": "Adjuvant_Neoadjuvant_Regimen_1",
        "Collection Date2": "Collection_Date_2",
        "Geographic Origin4": "GeographicOrigin_4",
        "Immunohistopathologic Data": "Immunohistopathologic_Data",
        "Surgery, Site And Side/Other info5": "Surgery_Site_Side_Info_5",
        "Survival status (added)": "Survival_Status",
        "Date last known alive (added)": "Last_Known_Alive_Date",
        "NHSB comments (added)": "NHSB_Comments",
        "Ki 67": "Ki67_Perc",
        "Date of collection": "Collection_Date",
        "Therapy 1": "Therapy_1",
        "Follow up 1": "FollowUp_1",
        "Therapy 2": "Therapy_2",
        "Follow up 2": "FollowUp_2",
        "date last follow up": "Last_FollowUp_Date",
        "TPS (PD-L1)": "PDL1_TPS",
        "CPS (PD-L1)": "PDL1_CPS",
        "RNAscope QC": "RNAscope_QC",
        "Year of collection": "Collection_Year",
        "Specimen Type": "Specimen_Type",
        "PPIB (TA3135-B)": "PPIB_TA3135_B",
        "PPIB (TA3136-B)": "PPIB_TA3136_B",
        "HER2 per TMA": "HER2_Per_TMA",
        "HER2+ per large section": "HER2_Positive_Per_LargeSection",
        "Collection date": "Collection_Date",
        "pTNM stage": "Pathological_TNM_Stage",
        "KRAS Mutation": "KRAS_Mutation",
        "EGFR Mutation": "EGFR_Mutation",
        "Date of clinical diagnosis": "Clinical_Diagnosis_Date",
        "Date of histological diagnosis": "Histological_Diagnosis_Date",
        "Neoadjuvant type of treatment": "Neoadjuvant_Treatment_Type",
        "Beginning of treatment": "Treatment_StartDate",
        "End of treatment date": "Treatment_EndDate",
        "Survival Month": "OverallSurvival_Months",
        "Treatment Subgrouping": "SpecTreatmentSubgrouping",
        "DATE OF  SURGERY": "SurgeryDate",
        "DATE OF TREATMENT 1": "Therapy_1_StartDate",
        "Stage2": "StageAtDX_2",
        "date of treatment 2": "Therapy_2_StartDate",
        "date of sample collection": "Collection_Date",
        "PPIB  (2025.12.05)": "PPIB_Score",
        "B7-H3 - Neg(%)": "B7H3_NegPerc",
        "TROP2 (1:3000) - 1+,2+,3+(%)": "TROP2_1_3000_1Plus_2Plus_3Plus_Perc",
        "PDL1_TPS_2": "PDL1_TPS_2",
        "PDL1_CPS2": "PDL1_CPS_2",
        "month of surgery": "Surgery_Month",
        "PPIB (C)": "PPIB_Score_C",
        "Overall survival (OS)": "OverallSurvival_Months",
        "Progression free survival (PFS)": "PFS_Months",
        "Responder/Non responder": "Responder_Status",
        "Best response or response vs nonresponse": "Best_Response",
        "date of relapse": "Relapse_Date",
        "Last date to follow-up": "Last_FollowUp_Date",
        "Progression free survival (PFS)15": "PFS_Months_15",
        "Responder/Non responder16": "Responder_Status_16",
        "Best response or response vs nonresponse17": "Best_Response_17",
        "date of relapse18": "Relapse_Date_18",
        "Last date to follow-up19": "Last_FollowUp_Date_19",
        "Death20": "Death_Date_20",
        "Overall survival (OS) (in months)": "OverallSurvival_Months",
        "Progression free survival (PFS) (in months)": "PFS_Months",
        "Comments on OS & PFS from clinical team": "Comments_ClinicalTeam",
        "PDL1 %": "PDL1_Perc",
        "Mutation data": "Molecular_Data",
        "TPS(PDL1)": "PDL1_TPS",
        "CPS(PDL1)": "PDL1_CPS",
        "Country of collection": "GeographicOrigin",
        "Site Code": "Site_Code",
        "Sample type": "Sample_Type",
        "RNAscope PPIB score": "RNAscope_PPIB_Score",
        "Patient Race": "Ethnicity",
        "Height, cm": "Height_CM",
        "Weight, kg": "Weight_KG",
        "MPL-Tissue - Sample weight, mg": "Sample_Weight_MG",
        "MPL-Tissue -  Date of tissue collection": "Tissue_Collection_Date",
        "MPL-Tissue - Time of tumour excision": "Tumour_Excision_Time",
        'MPL-Tissue - Minutes (") before fixation to Formalin': "Fixation_Delay_Minutes",
        "MPL-Tissue - Storage start date": "Storage_StartDate",
        "MPL-Tissue - Storage start time": "Storage_StartTime",
        "MPL-Tissue - Storage conditions": "Storage_Conditions",
        "Total volume, ml": "Total_Volume_ML",
        "Type  of vacutainer": "Vacutainer_Type",
        "Date of blood collection": "Blood_Collection_Date",
        "Time of blood collection": "Blood_Collection_Time",
        "Start time of processing": "Processing_StartTime",
        "Storage start date": "Storage_StartDate",
        "Storage start time": "Storage_StartTime",
        "Storage conditions": "Storage_Conditions",
        "Tumour Location": "SpecTumourSite",
        "Surgical procedure": "Surgical_Procedure",
        "Method of verification": "Verification_Method",
        "Tumor size (cm)-Height  X width X Length": "Tumor_Size_Dimensions_CM",
        "Distant metastasis location": "Distant_Metastasis_Location",
        "Cancer Stage": "StageAtDX",
        "Histologic Tumour Type": "Histological_Tumour_Type",
        "Gross description": "Gross_Description",
        "Microscopic description": "Microscopic_Description",
        "Subject received anticancer therapy in the preceding 6 months": "Prior_Therapy_6Months",
        "Subject has a personal history of a malignancy of any type": "Malignancy_History",
        "Subject has previous history of organ transplantation or blood product transfusion": "Transplant_Transfusion_History",
        "Type of Therapy/Dose": "Therapy_Dose_Type",
        "Hep B": "Hep_B_Status",
        "Hep C": "Hep_C_Status",
        "ICF date": "ICF_Date",
        "EGFR Mutated": "EGFR_Mutation",
        "KRAS Mutated": "KRAS_Mutation",
        "Protein Change": "Protein_Change",
        "PPIB (A)": "PPIB_Score_A",
        "PPIB (B)": "PPIB_Score_B",
        "Year of cellection": "Collection_Year",
        "LH number": "Sample_ID",
        "Block part": "Block_Part",
        "Study Number": "Study_Number",
        "Histological Diagnosis": "Histological_Diagnosis",
        "Sample obtained pre or post treatment": "Pre_Or_Post_Treatment",
        "Current status": "CurrentStatus",
        "Date of death": "Death_Date",
        "Histological Sub-Type": "Histological_Subtype",
        "Anonymized ID (Primary sample)": "AnonymisedID_PrimarySample",
        "Primary sample Block Number": "Primary_Sample_Block_Number",
        "Primary Surgery Date (Q/YYYY)": "Primary_Surgery_Date_QYYYY",
        "Site of primary tumour": "PrimTumourSite",
        "Anonymized ID Matched normal": "AnonymisedID_MatchedNormal",
        "Matched normal Block Number": "MatchedNormalTissue_BlockNumber",
        "Matched Normal Surgery Date (Q/YYYY)": "MatchedNormal_Surgery_Date_QYYYY",
        "Date of Dx (Q/YYYY)": "DateDX_QYYYY",
        "Resection margins": "Resection_Margins",
        "Block Tested": "Block_Tested",
        "P16 (pos/neg)": "P16_Status",
        "PDL1 (pos/neg)": "PDL1_Status",
        "PDL1 antibody": "PDL1_Antibody",
        "% cells positive": "Positive_Cells_Perc",
        "Immunotherapy naive (Y/N)": "Immunotherapy_Naive",
        "Immunotherapy Y/N": "AnyImmunotherapy",
        "Immunotherapy Agent": "Immunotherapy_Agent",
        "Response to immunotherapy": "Immunotherapy_Response",
        "Pre or Post": "Pre_Or_Post_Treatment",
        "Adjuvant Tx Start Date": "Adjuvant_Tx_StartDate",
        "Adjuvant Tx Finish Date": "Adjuvant_Tx_EndDate",
        "Age at Diagnosis": "AgeDX",
        "PDL1(TPS)": "PDL1_TPS",
        "PDL1(CPS)": "PDL1_CPS",
        "PDL1(TPS)2": "PDL1_TPS_2",
        "PDL1(CPS)3": "PDL1_CPS_3",
        "TA4384(A)": "TA4384_A",
        "PPIB (D)": "PPIB_Score_D",
        "Specimen surgery date": "SpecSurgeryDate",
        "Specimen surgery": "SpecSurgery",
        "Specimen tumour site": "SpecTumourSite",
        "Date of DX": "DateDX",
        "Specimen Histological Diagnosis": "SpecHistologicalDiagnosis",
        "Staging/grading info provided for:": "StagingGradingInfoProvidedFor",
        "Resection margins clear": "ResectionMarginsClear",
        "TNM stage": "TNM_Stage",
        "AJCC staging": "StagingCriteria",
        "mCRPC or mHSPC (added)": "Prostate_Cancer_Subtype",
        "Donor ID": "DonorID",
        "PPIB from TA3168a": "PPIB_Score_TA3168a",
        "ID written on the  TL sample (NSCLC)": "Sample_ID",
        "ID written on the  TS sample (Non-malignant)": "Sample_ID",
        "Pre-treatment ( neoadjuvant)": "Pretreatment_Neoadjuvant",
        "BRAF mutation": "BRAF_Mutation",
        "PD-L1": "PDL1_Status",
        "Other molecular data": "Molecular_Data",
        "Any Immunotherapy": "AnyImmunotherapy",
        "Specimen treatment subgrouping": "SpecTreatmentSubgrouping",
        "Specimen is a recurrence": "SpecIsRecurrence",
        "Primary Surgery date": "PrimSurgeryDate",
        "Primary tumour site": "PrimTumourSite",
        "Anonymized ID Normal tissue": "AnonymisedID_NormalTissue",
        "Matched Normal tissue Block Number": "MatchedNormalTissue_BlockNumber",
        "Matched Normal tissue Site": "MatchedNormalTissue_Site",
        "Matched Normal tissue Surgery Date": "MatchedNormalTissue_SurgeryDate",
        "Matched Normal treatment subgrouping": "MatchedNormal_TreatmentSubgrouping",
        "Stage at DX (AJCC)": "StageAtDX",
        "Comments from clinical team": "Comments_ClinicalTeam",
        "Specimen Block number 12": "SpecimenBlockNumber_12",
        "Specimen Block number 23": "SpecimenBlockNumber_23",
        "Specimen Block number 34": "SpecimenBlockNumber_34",
        "Block number": "Block_Number",
        "Specimen Block number": "SpecimenBlockNumber",
        "Specimen Surgery Date (Q/YYYY)": "Specimen_Surgery_Date_QYYYY",
        "Date of DX (Q/YYYY)": "DateDX_QYYYY",
        "Resection margins clear (Y/N)": "ResectionMarginsClear",
        "Primary surgery date (Month/Yr)": "Primary_Surgery_Date_MonthYear",
        "Block Type": "Block_Type",
        "Mets Site": "Mets_Site",
        "Mets Surgery Date": "Mets_Surgery_Date",
        "Mets treatment subgrouping": "Mets_Treatment_Subgrouping",
        "Primary surgery date": "PrimSurgeryDate",
        "Primary treatment subgrouping": "Primary_Treatment_Subgrouping",
        "KRAS mutation": "KRAS_Mutation",
        "NRAS mutation": "NRAS_Mutation",
        "% cells": "Positive_Cells_Perc",
        "EGFR mutation": "EGFR_Mutation",
        "Primary surgery (Year)": "Primary_Surgery_Year",
        "MMR result (normal/abnormal)": "MMR_Result_Normal_Abnormal",
        "MLH1 IHC (pos / neg / clonal neg)": "MLH1_IHC_Pos_Neg_ClonalNeg",
        "PMS2 IHC (pos/ neg / clonal neg": "PMS2_IHC_Pos_Neg_ClonalNeg",
        "MSH2 IHC (pos/ neg / clonal neg": "MSH2_IHC_Pos_Neg_ClonalNeg",
        "MSH6 IHC (pos/ neg/ clonal neg)": "MSH6_IHC_Pos_Neg_ClonalNeg",
        "FFPE Block ID": "Sample_ID",
        "DapB Score": "DapB_Score",
        "PPIB Score": "PPIB_Score",
        "std conditions": "Standard_Conditions",
        "QC Pass/Fail": "QC_Pass_Fail",
        "Test date": "Test_Date",
        "Matched met or normal site": "Matched_Met_Or_Normal_Site",
        "Matched met or normal surgery date": "Matched_Met_Or_Normal_SurgeryDate",
        "Matched met or normal treatment subgrouping": "Matched_Met_Or_Normal_TreatmentSubgrouping",
        "MACROSCOPIC DESCRIPTION - GROSS DESCRIPTION": "Gross_Description",
        "MACROSCOPIC DESCRIPTION - SIZE OF TUMOR": "Tumor_Size_MM",
        "MICROSCOPIC DESCRIPTION/HISTOLOGIC FEATURES - HISTOLOGICAL TUMOR TYPE": "Histological_Tumour_Type",
        "MICROSCOPIC DESCRIPTION/HISTOLOGIC FEATURES - HISTOLOGICAL GRADE": "Histological_Grade",
        "HISTOLOGIC FEATURES/ MICROSCOPIC DESCRIPTION": "Microscopic_Description",
        "MICROSCOPIC DESCRIPTION/HISTOLOGIC FEATURES -TUMOR INVASION": "Tumor_Invasion",
        "Extent of tumor": "Extent_Of_Tumor",
        "AJCC TNM (or other applicable) Clinical": "Clinical_TNM_Stage",
        "AJCC TNM (or other applicable) Pathological": "Pathological_TNM_Stage",
        "AJCC Tumor staging category (pT)": "Pathological_T_Stage",
        "AJCC Stage Grouping (or other applicable staging system)": "StageAtDX",
        "LYMPH NODES (Regional Lymph Notes) - Status": "Lymph_Nodes_Status",
        "LYMPH NODES (Regional Lymph Notes) - No. Assessed": "Lymph_Nodes_Num_Assessed",
        "LYMPH NODES (Regional Lymph Notes) - No. Positive": "Num_Pos_LymphNodes",
        "LYMPH NODES (Regional Lymph Notes) - Localization": "Lymph_Nodes_Localization",
        "Metastatic Site (Distant) - Status": "Metastatic_Status_Distant",
        "Metastatic Site (Distant) - Site": "Metastatic_Site_Distant",
        "Estrogen Receptor - ER": "ER_Status",
        "Progesterone Receptor - PR": "PR_Status",
        "Human Epidermal growth factor Receptor 2 - HER2": "HER2_Status",
        "Nectin-4 - 1+,2+,3+(%)2": "Nectin4_1Plus_2Plus_3Plus_Perc_2",
        "Nectin-4 - 2+,3+(%)2": "Nectin4_2Plus_3Plus_Perc_2",
        "T code": "T_Code",
        "M code": "M_Code",
        "Plasma tube ID": "Plasma_Tube_ID",
        "plasma tube ID": "Plasma_Tube_ID",
        "Normal site": "Normal_Tissue_Site",
        "Date of PPIB test": "PPIB_Test_Date",
        "Selected block": "Sample_ID",
        "Stage at dx (TNM)": "TNM_Stage",
        "Stage at dx (AJCC stage grouping)": "StageAtDX",
        "Item per 09/24 discussion - see Criteria tab for explanation": "Discussion_Item_Note",
        "Pretreatment history": "Pretreatment_History",
        "Neoadjuvant summary": "Neoadjuvant_Summary",
        "Neoadjuvant treatment 5": "NeoadjuvantTx5",
        "Adjuvant summary": "Adjuvant_Summary",
        "Subsequent treatment summary (all post-recurrence/progression treatments that preceded sample surgery)": "Subsequent_Treatment_Summary",
        "tissue type": "Tissue_Type",
        "Highlighted samples must be included in the TMA blocks. Select the remainder as needed to arrive at 108 samples in total across 9 TMA blocks": "TMA_Selection_Instructions",
        "Specimen Treatment Subgrouping": "SpecTreatmentSubgrouping",
        "Specimen Is A Recurrence": "SpecIsRecurrence",
        "Selected Specimen Anonymised ID": "Sample_ID",
        "Specimen Surgery Date": "SpecSurgeryDate",
        "Specimen Surgery": "SpecSurgery",
        "Specimen Tumour Site": "SpecTumourSite",
        "Primary Surgery Date": "PrimSurgeryDate",
        "Primary Tumour Site": "PrimTumourSite",
        "Anonymized ID Normal Tissue": "AnonymisedID_NormalTissue",
        "Matched Normal Tissue Block Number": "MatchedNormalTissue_BlockNumber",
        "Matched Normal Tissue Site": "MatchedNormalTissue_Site",
        "Matched Normal Tissue Surgery Date": "MatchedNormalTissue_SurgeryDate",
        "Matched Normal Treatment Subgrouping": "MatchedNormal_TreatmentSubgrouping",
        "Date Of Dx": "DateDX",
        "Staging/grading Info Provided For:": "StagingGradingInfoProvidedFor",
        "Resection Margins Clear": "ResectionMarginsClear",
        "TNM Stage": "TNM_Stage",
        "Staging Criteria": "StagingCriteria",
        "Stage At Dx": "StageAtDX",
        "Any Neoadjuvant Therapy": "AnyNeoadjuvantTherapy",
        "Response To Neoadjuvant Therapy (TRS Score)": "ResponseToNeoTherapy_TRScore",
        "Death Tumour Related": "DeathTumourRelated",
        "Comments From Clinical Team": "Comments_ClinicalTeam",
        "TTE Specimen Surgery Date": "TTE_SpecimenSurgeryDate",
        "TTE Primary Surgery Date": "TTE_PrimarySurgeryDate",
        "TTE Matched Normal Tissue Surgery Date": "TTE_MatchedNormalTissue_SurgeryDate",
        "TTE Matched Normal Tissue Surgery Date9": "TTE_MatchedNormalTissue_SurgeryDate_9",
        "TTE Adjuvant Tx1 Start Date10": "TTE_Adjuvant_Tx1_StartDate_10",
        "TTE Adjuvant Tx1 End Date11": "TTE_Adjuvant_Tx1_EndDate_11",
        "TTE Date Recurrence/ Progression 112": "TTE_DateRecurrenceProgression1_12",
        "TTE Subsequent Tx1 Start Date13": "TTE_Subsequent_Tx1_StartDate_13",
        "TTE Subsequent Tx1 End Date14": "TTE_Subsequent_Tx1_EndDate_14",
        "TTE Recurrence (PFS)": "TTE_Recurrence_PFS",
        "TTE Recurrence (PFS)15": "TTE_Recurrence_PFS_15",
        "TTE Death (OS And PFS)": "TTE_Death_OS_PFS",
        "TTE Death (OS And PFS)16": "TTE_Death_OS_PFS_16",
        "TTE Date Last Follow-Up (PFS)": "TTE_Date_LastFollowUp",
        "TTE Date Last Follow-Up (PFS)17": "TTE_Date_LastFollowUp_17",
        "TTE Date Last Seen (OS)": "TTE_DateLastSeen_OS",
        "TTE Date Last Seen (OS)18": "TTE_DateLastSeen_OS_18",
        "TTE Overall Survival (OS) In Months": "TTE_OverallSurvival_Months",
        "TTE Overall Survival (OS) In Months19": "TTE_OverallSurvival_Months_19",
        "TTE Progression Free Survival (PFS) In Months": "TTE_PFS_Months",
        "TTE Progression Free Survival (PFS) In Months20": "TTE_PFS_Months_20",
        "Event (Recurrence/Progression)": "Event_RecurrenceProgression",
        "Event (Recurrence/Progression)21": "Event_RecurrenceProgression_21",
        "Event (Cancer Death)": "Event_CancerDeath",
        "Event (Cancer Death)22": "Event_CancerDeath_22",
        "Event (Death; All Causes)": "Event_DeathAllCauses",
        "Event (Death; All Causes)23": "Event_DeathAllCauses_23",
        "Event (Recurrence Or Cancer Death": "Event_RecurrenceCancerDeath",
        "Event (Recurrence Or Cancer Death24": "Event_RecurrenceCancerDeath_24",
        "Treatment Type": "Treatment_Type",
        "MMR Result": "MMR_Result",
        "Neoadjuvant therapy (Yes/No)": "AnyNeoadjuvantTherapy",
        "Adjuvant therapy (Yes/No)": "AnyAdjuvantTherapy",
        "Matched met site": "MatchedMet_Site",
        "Matched met surgery date": "MatchedMet_SurgeryDate",
        "Matched met treatment subgrouping": "MatchedMet_TreatmentSubgrouping",
        "TTE Specimen Surgery Date2": "TTE_SpecimenSurgeryDate_2",
        "TTE Primary Surgery Date3": "TTE_PrimarySurgeryDate_3",
        "Condition Status": "Condition_Status",
        "Date of Dx": "DateDX",
        "Donor ID2": "DonorID_2",
        "ISUP Grade": "ISUP_Grade",
        "Tumor type": "SpecTumourType",
        "Psa (Ug/L) At Dx": "PSA_At_DX_UgPerLitre",
        "Castrate_Resistant_Prostate_Cancer (Yes = 1, No = 0)": "Castrate_Resistant_Prostate_Cancer",
        "Psa When Last Followed Up (Ug/L; Detection Limits Vary)": "PSA_Last_FollowUp_UgPerLitre",
        "Details Of Biochemical Recurrence": "Biochemical_Recurrence_Details",
        "Post Treatment Specimen 1 Path Number": "Post_Treatment_Specimen_1_Path_Number",
        "Last Treatment Prior To Specimen 1 Surgery": "Last_Treatment_Prior_To_Specimen_1_Surgery",
        "KRAS Mutation Group": "KRAS_Group",
        "Side_L_R2": "Side_L_R_2",
        "SpecNumber_Used_For_MolecularTesting_IfDifferent3": "SpecNumber_MolecularTesting_3",
        "Selected Sample ID": "SelSpecAnonymisedID",
        "Met site": "Met_Site",
        "Met surgery date": "Met_SurgeryDate",
        "Met treatment subgrouping": "Met_TreatmentSubgrouping",
        "Treatment Category": "Treatment_Category",
        "Specimen ID": "Sample_ID",
        "PPIB date": "PPIB_Test_Date",
        "Selected sample ID": "SelSpecAnonymisedID",
        "ID": "Sample_ID",
        "# case": "Sample_ID",
    }
)

# Rename all the columns in the tma_dfs
for tma_name, tma_df in tma_dfs.items():
    renamed_df = tma_df.rename(columns=column_name_map)
    renamed_df = renamed_df.loc[:, ~renamed_df.columns.duplicated(keep="first")]
    tma_dfs[tma_name] = renamed_df

######################################################################################################
# Upload the TMA dataframes to the database
######################################################################################################
# for tma_name, tma_df in tma_dfs.items():
#     # Upload the dataframe to the database
#     tma_df.to_sql(tma_name, con=engine, if_exists='replace', schema='tma', index=False)
#     print(f"Uploaded {tma_name} to the database.")

######################################################################################################
# Check the dataframes
######################################################################################################
# In[]:
for name, df in tma_dfs.items():
    if "Sample_ID" not in df.columns and "DonorID" not in df.columns:
        print(name)
