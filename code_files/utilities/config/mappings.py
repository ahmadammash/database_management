# Import required libraries
import json
import re

import ipywidgets as widgets
from IPython.display import display
from sqlalchemy.dialects.mssql import (
    BIGINT,
    BIT,
    CHAR,
    DATE,
    FLOAT,
    INTEGER,
    TINYINT,
    VARCHAR,
)

# Import the required functions
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

# load the json file for the cohort mapping
json_path = project_root / "json_files" / "cohort_names_mapping.json"
with open(json_path, "r") as f:
    cohort_mapping = json.load(f)

# ----------------------------------------------------------------------------------------------------------------------------------------


# define the function to set the data types of all columns
def set_data_types(main_workbook, dataset_type):
    """
    This function defines the mapping of column names to data types in the database
    """
    # Start with the base dictionary
    final_types = {
        "CohortName": VARCHAR(60),
        "DonorID": CHAR(10),
        "NHS_Number": BIGINT,
        "AnyImmunotherapy": BIT,
        "SpecTreatmentSubgrouping": VARCHAR(50),
        "SpecIsRecurrence": BIT,
        "SelSpecAnonymisedID": CHAR(50),
        "SpecBlockNum1": VARCHAR(10),
        "SpecBlockNum2": VARCHAR(10),
        "SpecSurgery": VARCHAR(500),
        "SpecTumourSite": VARCHAR(300),
        "PrimTumourSite": VARCHAR(100),
        "AnonymisedID_NormalTissue": CHAR(15),
        "MatchedNormalTissue_BlockNumber": VARCHAR(10),
        "MatchedNormalTissue_Site": VARCHAR(100),
        "MatchedNormal_TreatmentSubgrouping": VARCHAR(50),
        "Gender": VARCHAR(10),
        "AgeDX": TINYINT,
        "SpecHistologicalDiagnosis": VARCHAR(300),
        "StagingGradingInfoProvidedFor": VARCHAR(50),
        "ResectionMarginsClear": BIT,
        "pT": VARCHAR(300),
        "pN": VARCHAR(300),
        "pM": VARCHAR(300),
        "cM": VARCHAR(300),
        "TNM_Stage": VARCHAR(300),
        "Grade": VARCHAR(50),
        "StagingCriteria": VARCHAR(300),
        "StageAtDX": VARCHAR(100),
        "AnyNeoadjuvantTherapy": BIT,
        "ResponseToNeoTherapy_TRScore": VARCHAR(50),
        "CurrentStatus": VARCHAR(30),
        "DeathTumourRelated": BIT,
        "CauseOfDeath": VARCHAR(300),
        "Event_RecurrenceProgression": BIT,
        "Event_CancerDeath": BIT,
        "Event_DeathAllCauses": BIT,
        "Event_RecurrenceCancerDeath": BIT,
        "ConcurrentMedicalConditions": VARCHAR(300),
        "Comments_ClinicalTeam": VARCHAR(650),
        "CohortName_SpecimenNumber": VARCHAR(30),
        "Side_L_R": VARCHAR(30),
        "SpecNumber_Used_For_MolecularTesting_IfDifferent": VARCHAR(300),
        "Block_Used_For_LynchScreening_MMR_IHC": VARCHAR(100),
        "MMR_Result": VARCHAR(50),
        "MLH1_IHC_Pos_Neg_ClonalNeg": VARCHAR(50),
        "PMS2_IHC_Pos_Neg_ClonalNeg": VARCHAR(50),
        "MSH2_IHC_Pos_Neg_ClonalNeg": VARCHAR(50),
        "MSH6_IHC_Pos_Neg_ClonalNeg": VARCHAR(50),
        "KRAS_Mutation": VARCHAR(400),
        "NRAS_Mutation": VARCHAR(400),
        "BRAF_Mutation": VARCHAR(400),
        "PrimSurgery_Transplant_Resection": VARCHAR(30),
        "ResectionMarginsClear_HepaticRecurrences": BIT,
        "pT_HepaticRecurrence": VARCHAR(30),
        "pN_HepaticRecurrence": VARCHAR(30),
        "pM_HepaticRecurrence": VARCHAR(30),
        "cM_HepaticRecurrence": VARCHAR(30),
        "Grade_HepaticRecurrence": VARCHAR(30),
        "TNM_Stage_HepaticRecurrence": VARCHAR(30),
        "AJCC_Stage_HepaticRecurrence": VARCHAR(30),
        "Tumour_Type": VARCHAR(100),
        "Spec_Number_Used_For_Testing": VARCHAR(100),
        "Block_Used_For_Testing": VARCHAR(100),
        "PDL1": VARCHAR(30),
        "Percent_Cells": VARCHAR(30),
        "Ethnicity": VARCHAR(30),
        "Leibovich_Score": INTEGER,
        "PDL1_PosNeg": VARCHAR(100),
        "PDL1_Antibody": VARCHAR(30),
        "Percent_Cells_Pos": VARCHAR(50),
        "Block_Used_For_Screening": VARCHAR(20),
        "Spec_Receptor_Status": VARCHAR(30),
        "P16_PosNeg": VARCHAR(100),
        "Secondary_Block_Used_For_Screening": VARCHAR(30),
        "Secondary_PDL1_PosNeg": VARCHAR(30),
        "Secondary_PDL1_Antibody": VARCHAR(30),
        "Secondary_PDL1_Percent_Cells_Pos": VARCHAR(30),
        "Third_Block_Used_For_Screening": VARCHAR(30),
        "Third_PDL1_PosNeg": VARCHAR(30),
        "Third_PDL1_Antibody": VARCHAR(30),
        "Third_PDL1_Percent_Cells_Pos": VARCHAR(30),
        "Block_Used_For_Testing_P16": VARCHAR(100),
        "Details_Provided_For": VARCHAR(100),
        "PSA_At_Diagnosis_ugPerLitre": VARCHAR(50),
        "Gleason_Score": VARCHAR(100),
        "MSIH_MSS": VARCHAR(30),
        "EGFR_Mutation": VARCHAR(300),
        "Detection_Platform": VARCHAR(50),
        "ALK_Rearrangement_Fusion": VARCHAR(50),
        "Other_Genetic_Lesions": VARCHAR(50),
        "Notes": VARCHAR(650),
        "Any_Other_Variants_Detected": VARCHAR(50),
        "Androgen_Deprivation_Therapy": VARCHAR(50),
        "Biochemical_Relapse_During_ADT": VARCHAR(50),
        "Clinical_Relapse_During_ADT": VARCHAR(50),
        "Castrate_Resistant_Prostate_Cancer": BIT,
        "PSA_Nadir_ugPerLitre_DetectionLimitsVary": VARCHAR(50),
        "PSA_At_Last_FollowUp_ugPerLitre_DetectionLimitsVary": VARCHAR(50),
        "TTE_CRPC_Biochemical_From_ADT_StartDate": FLOAT,
        "TTE_CRPC_Clinical_From_ADT_StartDate": FLOAT,
        "TTE_CRPC_Biochemical_From_DateDiagnosis": FLOAT,
        "TTE_CRPC_Clinical_From_DateDiagnosis": FLOAT,
        "P16_Pos_Neg": VARCHAR(100),
        "HistopathologicalSubcategory_BiopsiesResections": VARCHAR(50),
        "BiopsyTreatmentSubgrouping": VARCHAR(50),
        "SelBiopsyAnonymisedID": CHAR(50),
        "BiopsyBlockNum1": VARCHAR(10),
        "BiopsyBlockNum2": VARCHAR(10),
        "Biopsy_HistopathologicalDiagnosis_PathReport": VARCHAR(300),
        "ResectionTreatmentSubgrouping": VARCHAR(50),
        "ResectionAnonymisedID": CHAR(50),
        "ResectionBlockNum1": VARCHAR(10),
        "ResectionBlockNum2": VARCHAR(10),
        "Resection_HistopathologicalDiagnosis_PathReport": VARCHAR(500),
        "ResectionSurgery": VARCHAR(300),
        "ResectionTumourSite": VARCHAR(300),
        "SpecBlockNum": VARCHAR(20),
        "Spec_IO_Naive": BIT,
        "Spec_Obtained_Pre_Post_Treatment": VARCHAR(100),
        "MatchedMet1_IO_Naive": BIT,
        "MatchedMet1_Obtained_Pre_Post_Treatment": VARCHAR(100),
        "MatchedMet2_TreatmentSubgrouping": VARCHAR(50),
        "MatchedMet2_IO_Naive": BIT,
        "MatchedMet2_Obtained_Pre_Post_Treatment": VARCHAR(100),
        "MatchedMet3_IO_Naive": BIT,
        "MatchedMet3_Obtained_Pre_Post_Treatment": VARCHAR(100),
        "MatchedMet3_Treatment_Subgrouping": VARCHAR(50),
        "SpecSurgeryDate": DATE if "unanonymised" in dataset_type else VARCHAR(50),
        "PrimSurgeryDate": DATE if "unanonymised" in dataset_type else VARCHAR(50),
        "Homologous_Recombination_Repair_Deficiency_Test_HRD": VARCHAR(100),
        "BRCA1_Mutation": VARCHAR(100),
        "BRCA2_Mutation": VARCHAR(100),
        "P53_Status_IHC": VARCHAR(100),
        "Inferred_P53_Mutation_Status": VARCHAR(100),
        "Platinum_Sensitive_Resistant_Refractory": VARCHAR(100),
        "Macroscopic_Surgical_Tumour_Debulking": VARCHAR(100),
        "CD56": VARCHAR(75),
        "Castrate_Resistant_Disease_At_Time_Sample_Surgery": BIT,
        "Met_Disease_at_DX_distMets_nonRegionalLN": BIT,
        "Met_Disease_At_Time_Sample_Surgery_distMets_nonRegionalLN": BIT,
        "TTE_Progression_To_Met_Disease_From_DateDX": FLOAT,
        "Disease_Status_At_Sample_Surgery_Met_NonMet_CRPC_HSPC": VARCHAR(50),
        "Date_Progression_To_Metastatic_Disease": DATE
    }

    dynamic_patterns = [
        (r"AnonymisedID_MatchedMet", CHAR(15)),
        (r"MatchedMet\d_BlockNumber", VARCHAR(20)),
        (r"MatchedMet\d_Site", VARCHAR(200)),
        (r"MatchedMet\d_TreatmentSubgrouping", VARCHAR(200)),
        (r"NeoadjuvantTx\d$", VARCHAR(300)),
        (r"NeoadjuvantTx\d_Agent", VARCHAR(300)),
        (r"TumourResponse_NeoadjuvantTx\d", VARCHAR(100)),
        (r"AdjuvantTx\d$", VARCHAR(300)),
        (r"AdjuvantTx\d_Agent", VARCHAR(300)),
        (r"TumourResponse_AdjuvantTx\d", VARCHAR(300)),
        (r"SubsequentTx\d$", VARCHAR(300)),
        (r"SubsequentTx\d_Agent", VARCHAR(300)),
        (r"TumourResponse_SubsequentTx\d", VARCHAR(300)),
        (r"Sites_RecurrenceProgression\d", VARCHAR(400)),
        (r"[A-Za-z]+_Dysplasia_Block(\d)$", VARCHAR(20)),
        (r"[A-Za-z]+_Dysplasia_Block(\d)_Site", VARCHAR(100)),
        (r"[A-Za-z0-9]+_Mutation_Shortened", VARCHAR(100)),
    ]

    # Iterate through the final execution columns
    for col in main_workbook.columns:
        col_name = str(col).strip()
        col_name_lower = col_name.lower()

        # 1. IMMEDIATE RULE FOR TTE: Intercept and catch ALL time-to-event columns first,
        # mapping them to FLOAT as they represent a number duration rather than a date stamp.
        if "tte_" in col_name_lower:
            final_types[col_name] = FLOAT
            continue

        # 2. CALENDAR DATE RULE: Catch remaining columns containing 'date' or 'dob'
        if "date" in col_name_lower or "dob" in col_name_lower:
            final_types[col_name] = (
                DATE if "unanonymised" in dataset_type.lower() else VARCHAR(50)
            )
            continue

        # 3. Skip if it's already explicitly defined inside the base dictionary
        if col_name in final_types:
            continue

        # 4. Handle Remaining dynamic patterns
        for regex, dtype in dynamic_patterns:
            if re.search(regex, col_name, re.IGNORECASE):
                final_types[col_name] = dtype
                break

    return final_types


# define a function to create the dictionary for renaming the columns in the main data
def create_column_map():
    """
    This function creates the column mapping dictionary used when uploading data to the database
    """
    column_name_map = {
        "NHS": "NHS_Number",
        "Nhs": "NHS_Number",
        "Nhs Number": "NHS_Number",
        "NHS number": "NHS_Number",
        "Specimen Id (Tumour)": "SelSpecAnonymisedID",
        "Donor Id": "DonorID",
        "Any Immunotherapy": "AnyImmunotherapy",
        "Specimen Treatment Subgrouping": "SpecTreatmentSubgrouping",
        "Specimen Is A Recurrence": "SpecIsRecurrence",
        "Selected Specimen Anonymised Id": "SelSpecAnonymisedID",
        "Specimen Block Number 1": "SpecBlockNum1",
        "Specimen Block Number 2": "SpecBlockNum2",
        "Specimen Surgery Date": "SpecSurgeryDate",
        "Specimen Surgery": "SpecSurgery",
        "Specimen Tumour Site": "SpecTumourSite",
        "Primary Surgery Date": "PrimSurgeryDate",
        "Primary Tumour Site": "PrimTumourSite",
        "Anonymised Id Normal Tissue": "AnonymisedID_NormalTissue",
        "Matched Normal Tissue Block Number": "MatchedNormalTissue_BlockNumber",
        "Matched Normal Tissue Site": "MatchedNormalTissue_Site",
        "Matched Normal Tissue Surgery Date": "MatchedNormalTissue_SurgeryDate",
        "Matched Normal Treatment Subgrouping": "MatchedNormal_TreatmentSubgrouping",
        "Date Of Dx": "DateDX",
        "Age At Diagnosis": "AgeDX",
        "Specimen Histological Diagnosis": "SpecHistologicalDiagnosis",
        "Staging/Grading Info Provided For:": "StagingGradingInfoProvidedFor",
        "Resection Margins Clear": "ResectionMarginsClear",
        "Tnm Stage": "TNM_Stage",
        "Staging Criteria": "StagingCriteria",
        "Stage At Dx": "StageAtDX",
        "Any Neoadjuvant Therapy": "AnyNeoadjuvantTherapy",
        "Response To Neoadjuvant Therapy (Trs Score)": "ResponseToNeoTherapy_TRScore",
        "Sites Recurrence/ Progression 1": "Sites_RecurrenceProgression1",
        "Current Status": "CurrentStatus",
        "Death Tumour Related": "DeathTumourRelated",
        "Cause Of Death": "CauseOfDeath",
        "Tte Neoadjuvant Tx1 End Date": "TTE_NeoadjuvantTx1EndDate",
        "Tte Specimen Surgery Date": "TTE_SpecimenSurgeryDate",
        "Tte Primary Surgery Date": "TTE_PrimarySurgeryDate",
        "Tte Matched Normal Tissue Surgery Date": "TTE_MatchedNormalTissue_SurgeryDate",
        "Tte Adjuvant Tx1 Start Date": "TTE_AdjuvantTx1StartDate",
        "Tte Adjuvant Tx1 End Date": "TTE_AdjuvantTx1EndDate",
        "Tte Date Recurrence/ Progression 1": "TTE_DateRecurrenceProgression1",
        "Tte Subsequent Tx1 Start Date": "TTE_SubsequentTx1StartDate",
        "Tte Subsequent Tx1 End Date": "TTE_SubsequentTx1EndDate",
        "Tte Recurrence (Pfs)": "TTE_Recurrence_PFS",
        "Tte Death (Os And Pfs)": "TTE_Death_OS_PFS",
        "Tte Date Last Follow-Up (Pfs)": "TTE_Date_LastFollowUp",
        "Tte Date Last Seen (Os)": "TTE_DateLastSeen_OS",
        "Tte Overall Survival (Os) In Months": "TTE_OverallSurvival_Months",
        "Tte Progression Free Survival (Pfs) In Months": "TTE_PFS_Months",
        "Event (Recurrence/Progression)": "Event_RecurrenceProgression",
        "Event (Cancer Death)": "Event_CancerDeath",
        "Event (Death; All Causes)": "Event_DeathAllCauses",
        "Event (Recurrence Or Cancer Death)": "Event_RecurrenceCancerDeath",
        "Event (Recurrence Or Cancer Death": "Event_RecurrenceCancerDeath",
        "Pt": "pT",
        "Pn": "pN",
        "Pm": "pM",
        "Cm": "cM",
        "Concurrent Medical Conditions": "ConcurrentMedicalConditions",
        "Comments From Clinical Team": "Comments_ClinicalTeam",
        "Death Date": "Death_Date",
        "Date Last Follow-Up": "Date_LastFollowUp",
        "Date Last Seen": "Date_LastSeen",
        "Mmr Result (Normal/Abnormal)": "MMR_Result",
        "Side L/R": "Side_L_R",
        "Specimen Number Used For Molecular Testing (If Different)": "SpecNumber_Used_For_MolecularTesting_IfDifferent",
        "Block Used For Lynch Screening (Mmr Ihc)": "Block_Used_For_LynchScreening_MMR_IHC",
        "Mlh1 Ihc (Pos / Neg / Clonal Neg)": "MLH1_IHC_Pos_Neg_ClonalNeg",
        "Pms2 Ihc (Pos/ Neg / Clonal Neg": "PMS2_IHC_Pos_Neg_ClonalNeg",
        "Msh2 Ihc (Pos/ Neg / Clonal Neg": "MSH2_IHC_Pos_Neg_ClonalNeg",
        "Msh6 Ihc (Pos/ Neg/ Clonal Neg)": "MSH6_IHC_Pos_Neg_ClonalNeg",
        "Kras Mutation": "KRAS_Mutation",
        "Nras Mutation": "NRAS_Mutation",
        "Braf Mutation": "BRAF_Mutation",
        "Primary Surgery: Transplant / Resection": "PrimSurgery_Transplant_Resection",
        "Resection Margins Clear (Hepatic Recurrences)": "ResectionMarginsClear_HepaticRecurrences",
        "Pt (Hepatic Recurrence)": "pT_HepaticRecurrence",
        "Pn (Hepatic Recurrence)": "pN_HepaticRecurrence",
        "Pm (Hepatic Recurrence)": "pM_HepaticRecurrence",
        "Cm (Hepatic Recurrence)": "cM_HepaticRecurrence",
        "Grade (Hepatic Recurrence)": "Grade_HepaticRecurrence",
        "Tnm Stage (Hepatic Recurrence)": "TNM_Stage_HepaticRecurrence",
        "Ajcc Stage (Hepatic Recurrence)": "AJCC_Stage_HepaticRecurrence",
        "Tumour Type": "Tumour_Type",
        "Specimen Number Used For Testing": "Spec_Number_Used_For_Testing",
        "Block Used For Testing": "Block_Used_For_Testing",
        "Block Used For Screening": "Block_Used_For_Screening",
        "Pdl-1": "PDL1_PosNeg",
        "Pdl1 (Pos/Neg)": "PDL1_PosNeg",
        "Pdl1 Antibody": "PDL1_Antibody",
        "% Cells Positive": "Percent_Cells_Pos",
        "% Cells": "Percent_Cells_Pos",
        "Leibovich Score": "Leibovich_Score",
        "Ethinicity": "Ethnicity",
        "Specimen Donor Id": "DonorID",
        "Specimen Receptor Status": "Spec_Receptor_Status",
        "P16 (Pos/Neg)": "P16_PosNeg",
        "P16": "P16_PosNeg",
        "Pdl1": "PDL1_PosNeg",
        "Secondary Block Used For Screening": "Secondary_Block_Used_For_Screening",
        "Secondary Pdl1 Test (Pos/Neg)": "Secondary_PDL1_PosNeg",
        "Secondary Pdl1 Antibody": "Secondary_PDL1_Antibody",
        "Secondary Pdl1 Test % Cells Positive": "Secondary_PDL1_Percent_Cells_Pos",
        "Third Block Used For Screening": "Third_Block_Used_For_Screening",
        "Third Pdl1 Test (Pos/Neg)": "Third_PDL1_PosNeg",
        "Third Pdl1 Antibody": "Third_PDL1_Antibody",
        "Third Pdl1 Test % Cells Positive": "Third_PDL1_Percent_Cells_Pos",
        "Block Used For Testing P16": "Block_Used_For_Testing_P16",
        "Specimen Id": "SelSpecAnonymisedID",
        "Details Provided For:": "Details_Provided_For",
        "Psa At Diagnosis (Ug/L)": "PSA_At_Diagnosis_ugPerLitre",
        "Gleason Score": "Gleason_Score",
        "Specimen Used For Testing If Different": "SpecNumber_Used_For_MolecularTesting_IfDifferent",
        "Msi-H/Mss": "MSIH_MSS",
        "Specimen Used For Testing (If Different)": "SpecNumber_Used_For_MolecularTesting_IfDifferent",
        "Egfr Mutation(S)": "EGFR_Mutation",
        "Detection Platform": "Detection_Platform",
        "Braf Muation": "BRAF_Mutation",
        "Alk Rearrangement": "ALK_Rearrangement_Fusion",
        "Pd-L1": "PDL1_PosNeg",
        "Pd-L1 % Positive Cells": "Percent_Cells_Pos",
        "Other Genetic Lesions": "Other_Genetic_Lesions",
        "Detection Platform (Egfr)": "Detection_Platform",
        "Specimen Number Used For Testing If Different": "SpecNumber_Used_For_MolecularTesting_IfDifferent",
        "% Tumour Cells Pd-L1 Positive": "Percent_Cells_Pos",
        "Egfr Mutation": "EGFR_Mutation",
        "Alk Rearrangement/Fusion": "ALK_Rearrangement_Fusion",
        "% Positive Cells (Pd-L1)": "Percent_Cells_Pos",
        "Any Other Variants Detected": "Any_Other_Variants_Detected",
        "Specimen": "SelSpecAnonymisedID",
        "Psa (Ug/L)": "PSA_At_Diagnosis_ugPerLitre",
        "Androgen Deprivation Therapy (Adt)": "Androgen_Deprivation_Therapy",
        "Biochemical Relapse During Adt": "Biochemical_Relapse_During_ADT",
        "Clinical Relapse During Adt": "Clinical_Relapse_During_ADT",
        "Castrate Resistant Prostate Cancer (Crpc)": "Castrate_Resistant_Prostate_Cancer",
        "Psa Nadir (Ug/L; Detection Limits Vary)": "PSA_Nadir_ugPerLitre_DetectionLimitsVary",
        "Psa At Last Follow Up (Ug/L; Detection Limits Vary)": "PSA_At_Last_FollowUp_ugPerLitre_DetectionLimitsVary",
        "Tte Crpc (Biochemical) (From Adt Start Date)": "TTE_CRPC_Biochemical_From_ADT_StartDate",
        "Tte Crpc (Clinical) (From Adt Start Date)": "TTE_CRPC_Clinical_From_ADT_StartDate",
        "Tte Crpc (Biochemical) From Date Diagnosis": "TTE_CRPC_Biochemical_From_DateDiagnosis",
        "Tte Crpc (Clinical) From Date Diagnosis": "TTE_CRPC_Clinical_From_DateDiagnosis",
        "Histopathological Subcategory (Biopsies And Resections)": "HistopathologicalSubcategory_BiopsiesResections",
        "Biopsy Treatment Subgrouping": "BiopsyTreatmentSubgrouping",
        "Selected Biopsy Anonymised Id": "SelBiopsyAnonymisedID",
        "Biopsy Block Number 1": "BiopsyBlockNum1",
        "Biopsy Block Number 2": "BiopsyBlockNum2",
        "Biopsy Histopathological Diagnosis (Pathology Report)": "Biopsy_HistopathologicalDiagnosis_PathReport",
        "Biopsy Surgery Date": "BiopsySurgeryDate",
        "Resection Treatment Subgrouping": "ResectionTreatmentSubgrouping",
        "Resection Anonymised Id": "ResectionAnonymisedID",
        "Resection Block Number 1": "ResectionBlockNum1",
        "Resection Block Number 2": "ResectionBlockNum2",
        "Resection Histopathological Diagnosis (Pathology Report)": "Resection_HistopathologicalDiagnosis_PathReport",
        "Resection Surgery Date": "ResectionSurgeryDate",
        "Resection Surgery": "ResectionSurgery",
        "Resection Tumour Site": "ResectionTumourSite",
        "Specimen Block Number": "SpecBlockNum",
        "Specimen Io Naive (Y/N)": "Spec_IO_Naive",
        "Specimen Obtained Pre Or Post Treatment": "Spec_Obtained_Pre_Post_Treatment",
        "Specimen Is A Recurrence (Y/N)": "SpecIsRecurrence",
        "Specimen Treatment Subgrouping.1": "MatchedMet1_TreatmentSubgrouping",
        "Specimen Io Naive (Y/N).1": "MatchedMet1_IO_Naive",
        "Specimen Obtained Pre Or Post Treatment.1": "MatchedMet1_Obtained_Pre_Post_Treatment",
        "Specimen Treatment Subgrouping.2": "MatchedMet2_TreatmentSubgrouping",
        "Specimen Io Naive (Y/N).2": "MatchedMet2_IO_Naive",
        "Specimen Io Naive (Y/N).3": "MatchedMet3_IO_Naive",
        "Specimen Treatment Subgrouping.3": "MatchedMet3_Treatment_Subgrouping",
        "Dob (Dd/Mm/Yyyy)": "DOB",
        "Date Of Dx (Dd/Mm/Yyyy)": "DateDX",
        "Age At Diagnosis (Years)": "AgeDX",
        "Resection Margins Clear Y/N)": "ResectionMarginsClear",
        "Stage At Dx (Ajcc)": "StageAtDX_AJCC",
        "Treatment Subgrouping": "SpecTreatmentSubgrouping",
        "Specimen Io Naive? Y/N": "Spec_IO_Naive",
        "Treatment Subgrouping.1": "MatchedMet1_TreatmentSubgrouping",
        "Specimen Io Naive? Y/N.1": "MatchedMet1_IO_Naive",
        "Resection Margins Clear (Y/N)": "ResectionMarginsClear",
        "Stage": "StageAtDX",
        "Progesterone Receptor Status": "PR_Status",
        "Oestrogen Receptor Status": "ER_Status",
        "Her2 Ihc Score": "HER2_IHC_Score",
        "Her2 Fish Score": "HER2_FISH_Score",
        "Nottingham Prognostic Index": "Nottingham_Prognostic_Index",
        "Pathogenic Germline Gene Variants Identified": "Pathogenic_Germline_Gene_Variants_Identified",
        "Genes Tested": "Genes_Tested",
        "Oncotype Dx Score (Done On Resection)": "Oncotype_DX_Score_DoneOnResection",
        "Anonymised Id Data Taken From": "AnonymisedID_Data_Taken_From",
        "Msi Status": "MSI_Status",
        "Specimen Number": "SelSpecAnonymisedID",
        "Figo Staging": "FIGO_Stage",
        "Specimen Used For P53 Testing (If Different)": "Spec_Used_For_P53_Testing_IfDifferent",
        "P53 Status (Ihc)": "P53_Status_IHC",
        "Inferred P53 Mutation Status": "Inferred_P53_Mutation_Status",
        "Specimen Number Used For Mmr Screening (If Different)": "Spec_Used_For_MMR_Screening_IfDifferent",
        "Block Used For Mmr Ihc": "Block_Used_For_MMR_IHC",
        "Mlh1 Promoter Hypermethylation (Detected/Not Detected/Not Tested)": "MLH1_Promoter_Hypermethylation_DetectionResult",
        "Lynch Syndrome": "Lynch_Syndrome",
        "Inferred Msi Status": "Inferred_MSI_Status",
        "Specimen Number Used For Screening (If Different)": "Spec_Used_For_Screening_IfDifferent",
        "Who Grade": "WHO_Grade",
        "Mgmt Promoter Methylation": "MGMT_Promoter_Methylation",
        "Tert Promoter Mutation": "TERT_Promoter_Mutation",
        "Ngs Idh1 Mutation": "NGS_IDH1_Mutation",
        "Ihc Idh1 Mutation": "IHC_IDH1_Mutation",
        "Idh2 Mutation": "IDH2_Mutation",
        "Histone H3F3A Mutation": "Histone_H3F3A_Mutation",
        "Histone Hist1H3B Mutation": "Histone_HIST1H3B_Mutation",
        "Egfr Amplification": "EGFR_Amplification",
        "Cdkn2A Loss": "CDKN2A_Loss",
        "Chromosome 1P36 Loss": "Chromosome_1P36_Loss",
        "Chromosome 19Q13 Loss": "Chromosome_19Q13_Loss",
        "Details Of Biochemical Recurrence": "Details_Of_Biochemical_Recurrence",
        "Psa When Last Followed Up (Ug/L; Detection Limits Vary)": "PSA_At_Last_FollowUp_ugPerLitre_DetectionLimitsVary",
        "Tumour Type (Primary)": "Tumour_Type",
        "Macroscopic Surgical Tumour Debulking: R0 (No Residual Disease), R1 (<1 Cm Residual Disease), R2 (1 Cm And Over Residual Disease)": "Macroscopic_Surgical_Tumour_Debulking",
        "Msi Result": "MSI_Status",
        "Focal High Grade Dysplasia Block Number": "Focal_High_Grade_Dysplasia_BlockNumber",
        "Focal High Grade Dysplasia Site": "Focal_High_Grade_Dysplasia_Site",
        "Focal Low Grade Dysplasia Block Number": "Focal_Low_Grade_Dysplasia_BlockNumber",
        "Focal Low Grade Dysplasia Site": "Focal_Low_Grade_Dysplasia_Site",
        "Histopathological Subcategory (Biopsies And Resections Project)": "Histopathological_Subcategory_BiopsiesResections_Project",
        "Histopathological Subtype (Biopsies And Resections Project)": "Histopathological_Subtype_BiopsiesResections_Project",
        "Isup Grade": "ISUP_Grade",
        "Low Grade Bilin Block 1": "Low_Grade_Biliin_Block1",
        "Low Grade Bilin Block 1 Site": "Low_Grade_Bilin_Block1_Site",
        "Low Grade Dysplasia Block Number": "Low_Grade_Dysplasia_BlockNumber",
        "Low Grade Dysplasia Site": "Low_Grade_Dysplasia_Site",
        "Matched Met Treatment Subgrouping": "MatchedMet_TreatmentSubgrouping",
        "Mitotic Count": "Mitotic_Count",
        "Mixed Low/High Grade Dysplasia Block Number": "Mixed_Low_High_Grade_Dysplasia_BlockNumber",
        "Mixed Low/High Grade Dysplasia Site": "Mixed_Low_High_Grade_Dysplasia_Site",
        "Msi Result (Stable/Unstable)": "MSI_Status",
        "Number Positive Lymph Nodes": "Num_Pos_LymphNodes",
        "Pd-L1 Combined Positive Score (Cps) ( Dako Pd-L1 Ihc 22C3 Pharmdx Test)": "PDL1_Combined_Pos_Score_CPS_DAKO_PDL1_IHC_Pharmdx_Test",
        "Pd-L1 Proportion Tumour Infiltrating Immune Cells (Ic) (Ventana Pd-L1 Sp142 Assay)": "PDL1_Proportion_Tumour_Infiltrating_ImmuneCells_Ventana_PDL1_SP142_Assay",
        "Pd-L1 Score For Atesolisumab": "PDL1_Score_For_Atesolisumab",
        "Pd-L1 Score For Nivolumab": "PDL1_Score_For_Nivolumab",
        "Pd-L1 Score For Pembrolisumab": "PDL1_Score_For_Pembrolisumab",
        "Pd-L1 Tumour Positive Score (Tps) ( Dako Pd-L1 28-8 Pharmdx Test)": "PDL1_Tumour_Pos_Score_TPS_DAKO_PDL1_28_8_Pharmdx_Test",
        "Pn Stage (Ajcc 7)": "pN_Stage_AJCC7",
        "Pn Stage (Ajcc 8)": "pN_Stage_AJCC8",
        "Post Treatment Specimen 1 Path Number": "Post_Treatment_Spec1_PathNumber",
        "Psa (Ug/L) At Dx": "PSA_At_Diagnosis_ugPerLitre",
        "Resection Margins Clear (1Mm And Above)": "ResectionMarginsClear",
        "Selected Specimen": "SelSpecAnonymisedID",
        "Specimen Is Hormone Receptor Positive & Her2 Negative": "Spec_Is_Hormone_Receptor_Pos_And_HER2_Neg",
        "Specimen Number Used For Testing (If Different)": "SpecNumber_Used_For_MolecularTesting_IfDifferent",
        "Specimen Obtained Pre Or Post Treatment.2": "Spec_Obtained_Pre_Post_Treatment",
        "Specimen Obtained Pre Or Post Treatment.3": "Spec_Obtained_Pre_Post_Treatment",
        "Specimen Site": "SpecTumourSite",
        "Stage Group (Ajcc 7)": "Stage_AJCC7",
        "Stage Group (Ajcc 8)": "Stage_AJCC8",
        "Tte Biopsy Surgery Date": "TTE_Biopsy_SurgeryDate",
        "Tte Resection Surgery Date": "TTE_Resection_SurgeryDate",
        "Tumor Necrosis": "Tumour_Necrosis",
        "Tumour Site Area": "Tumour_Site_Area",
        "Homologous Recombination Repair Deficiency Test (Hrd)": "Homologous_Recombination_Repair_Deficiency_Test_HRD",
        "Mmr Deficient/Proficient (Ihc)": "MMR_Result",
        "Brca1 Mutation": "BRCA1_Mutation",
        "Brca2 Mutation": "BRCA2_Mutation",
        "P53 Staining Pattern (Ihc)": "P53_Status_IHC",
        "Platinum Sensitive/Resistant/Refractory": "Platinum_Sensitive_Resistant_Refractory",
        "Cd56": "CD56",
        "Pd-L1 Combined Positive Score (Cps)": "PDL1_Combined_Pos_Score_CPS",
        "Gastric Her2 Ihc": "Gastric_HER2_IHC",
        "Tumour Block 1": "SpecBlockNum1",
        "Year Of Specimen Surgery": "SpecSurgeryDate",
        "Pt Stage": "pT",
        "Specimen Is Post Platinum-Based Chemotherapy (Yes/No)": "Spec_Obtained_Post_Platinum_Based_Chemotherapy",
        "Site (Tumour Block 1)": "Block1_Site",
        "Site (Tumour Block 2)": "Block2_Site",
        "Normal Tissue Id": "Normal_Tissue_ID",
        "Normal Block 1": "Normal_SpecBlockNum1",
        "Normal Tissue Surgery": "Normal_Tissue_Surgery",
        "Year Of Normal Tissue Surgery": "Normal_Tissue_SurgeryDate",
        "Normal Tissue Is Post Platinum-Based Chemotherapy (Yes/No)": "Normal_Spec_Obtained_Post_Platinum_Based_Chemotherapy",
        "Site (Normal Tissue Block 1)": "Normal_Block1_Site",
        "Date Weight And Height Taken": "Date_Weight_Height_Taken",
        "% Positive Cells": "Percent_Cells_Pos",
        "Notes_Scan2": "Notes_Scan_2",
        "Reason For Stopping Therapy 2": "Reason_For_Stopping_Therapy_2",
        "Tte Cystectomy": "TTE_Cystectomy",
        "Test Details": "Test_Details",
        "Sample Obtained Pre Or Post Treatment.1": "Spec_Obtained_Pre_Post_Treatment",
        "Histopathological Diagnosis": "SpecHistologicalDiagnosis",
        "Specimen Surgery Date (Q/Yyyy)": "SpecSurgeryDate",
        "Block Used For P53 Testing": "Block_Used_For_P53_Testing",
        "Ki67": "KI_67",
        "Who Grade (Primary)": "WHO_Grade_Primary",
        "Any Other Mutation": "Any_Other_Mutation",
        "Dob": "DOB",
        "Date Biochemical Recurrence/Progression During Adt": "Date_Biochemical_RecurrenceProgression_During_ADT",
        "Smoking Status At Dx": "Smoking_Status_At_DX",
        "Height (M)": "Height_meters",
        "S1 = Elevated Up To 1.5 X Upper Normal Limit (<645)": "S1_Elevated_Up_To_1.5xUpperNormalLimit_LessThan645",
        "Therapy 2 Name": "Therapy2_Name",
        "Weight (Kg)": "Weight_KG",
        "How To Record": "How_To_Record",
        "Specimen Used For Pd-L1 Testing (If Different)": "Spec_Used_For_PDL1_Testing_If_Different",
        "Pt Stage (Cystectomy)": "pT_Cystectomy",
        "Specimen Used For Pole Testing (If Different)": "Spec_Used_For_POLE_Testing_If_Different",
        "Tte Date Height And Weight Taken": "TTE_Date_Height_Weight_Taken",
        "AnonymisedID": "SelSpecAnonymisedID",
        "Matched Met 1 Surgery Date (Q/Yyyy)": "MatchedMet1_SurgeryDate",
        "Date Clinical Recurrence/Progression During Adt": "Date_Clinical_RecurrenceProgression_During_ADT",
        "Her2 Score": "HER2_Score",
        "Therapy 6 Name": "Therapy6_Name",
        "Reason For Stopping Therapy 10": "Reason_For_Stopping_Therapy_10",
        "Serum Afp At Dx (Ku/L; Ref. <7 Ku/L)": "Serum_AFP_At_DX_KuPerL_RefLessThan7KuPerL",
        "Cystectomy Date": "Cystectomy_Date",
        "Reason For Stopping Therapy 3": "Reason_For_Stopping_Therapy_3",
        "Sample Obtained Pre Or Post Treatment": "Spec1_Obtained_Pre_Post_Treatment",
        "Tte Crpc (Clinical) From Date Diagnosis.1": "TTE_CRPC_Clinical_From_DateDiagnosis_1",
        "Tumour Grade": "Grade",
        "Lh Number": "SelSpecAnonymisedID",
        "TreatmentSubgrouping": "SpecTreatmentSubgrouping",
        "Sample Obtained Pre Or Post Treatment.2": "Spec2_Obtained_Pre_Post_Treatment",
        "Subsequent Tx16 Agent": "Subsequent_Tx16_Agent",
        "Mmr Result (Normal(Proficient)/Abnormal (Deficient)": "MMR_Result",
        "Figo Grading": "FIGO_Grade",
        "Response To Neoadjuvant Therapy (Crs Score)": "ResponseToNeoTherapy_CRScore",
        "S Stage Hcg (S0= Normal; S1= 5.0 - 5000 Iu/L": "S_Stage_HCG_S0EqNormal_S1BW_5_5000IuPerL",
        "Stic Specimen Block Number 2": "STIC_SpecBlockNum2",
        "S3>10 X (>4300)": "S3_Above_10X_Above4300",
        "S Stage Afp": "S_Stage_AFP",
        "Subsequent Therapy (Post-Cystectomy)": "Subsequent_Therapy_Post_Cystectomy",
        "Therapy 3 Name": "Therapy_3_Name",
        "Date Adt Therapy First Initiated": "Date_ADT_Therapy_First_Initiated",
        "Germline Brca1 Status": "Germline_BRCA1_Status",
        "Age At Cystectomy": "AgeAtCystectomy",
        "SpecimenNumber": "SelSpecAnonymisedID",
        "Pole Mutation": "POLE_Mutation",
        "Therapy 10 Name": "Therapy_10_Name",
        "BlockNum": "SpecBlockNum",
        "Specimen Histopathological Diagnosis": "SpecHistologicalDiagnosis",
        "Reason For Stopping Therapy 5": "Reason_For_Stopping_Therapy_5",
        "Therapy 1 Name": "Therapy_1_Name",
        "Date Dx": "DateDX",
        "Therapy 7 Name": "Therapy_7_Name",
        "Date Of Dx (Q/Yyyy)": "Date_DX",
        "Block Tested": "Block_Tested",
        "Subsequent Tx16": "Subsequent_Tx16",
        "LH_Number": "SelSpecAnonymisedID",
        "Tumour Response To Subsequent Tx16": "TumourResponse_SubsequentTx16",
        "Serum Total Pregnancy Hcg At Dx (Ref. <5 U/L)": "Serum_Total_Pregnancy_HGC_At_DX_RefLessThan5UPerL",
        "Therapy 9 Name": "Therapy_9_Name",
        "Specimen Surgery (Cystectomy/Turbt)": "Specimen_Surgery_Cystectomy_TURBT",
        "000 Iu/L": "000_IuPerL",
        "Er (Ihc) Allred Score": "ER_IHC_AllRed_Score",
        "Reason For Stopping Therapy 1": "Reason_For_Stopping_Therapy1",
        "Psa": "PSA",
        "Reason For Stopping Therapy 9": "Reason_For_Stopping_Therapy_9",
        "Stic Specimen Block Number 1": "STIC_SpecimenBlockNum",
        "Pd-L1 Status": "PDL1_Status",
        "Specimen Tumour Block Number 1": "SpecBlockNum1",
        "Figo 2018 Grading": "FIGO_2018_Grade",
        "Block Used For Msi Testing": "Block_Used_For_MSI_Testing",
        "Serum Ldh At Dx (Ref. Range 160-430 Iu/L)": "Serum_LDH_At_DX_Ref_BW_160_430_IuPerL",
        "Summary S Stage": "Summary_S_Stage",
        "Histopathological Subcategory (Biopsies And Resectios)": "Histopathological_Subcategory_BiopsiesResections",
        "Pd-L1 Cps Score": "PDL1_CPS_Score",
        "Pt Stage (Preceding Turbt)": "pT_Stage_Preceding_TURBT",
        "Specimen Used For Msi Testing (If Different)": "Spec_Used_For_MSI_Testing_If_Different",
        "Therapy 4 Name": "Therapy_4_Name",
        "Therapy 5 Name": "Therapy_5_Name",
        "Select For Bms": "Select_For_BMS",
        "Date Diagnosis": "DateDX",
        "Msi Plus Test": "MSI_Plus_Test",
        "Primary Cancer Type": "Primary_Cancer_Type",
        "Molecular Testing/Grading Provided For:": "Molecular_Testing_Grading_Info_Provided_For",
        "Smoking History At Dx": "Smoking_History_At_DX",
        "Reason For Stopping Therapy 7": "Reason_For_Stopping_Therapy_7",
        "Germline Brca2 Status": "Germline_BRCA2_Status",
        "S2 = 5000-50": "S2_5000_50",
        "Reason For Stopping Therapy 6": "Reason_For_Stopping_Therapy_6",
        "Reason For Stopping Therapy 4": "Reason_For_Stopping_Therapy_4",
        "Therapy 8 Name": "Therapy_8_Name",
        "Specimen Tumour Block Number 2": "SpecBlockNum2",
        "Ca125 At Diagnosis (Ku/L)": "CA125_At_DX_KuPerL",
        "Primary Tumour Block 1": "Primary_SpecBlockNum1",
        "Other Tests": "Other_Tests",
        "Stage Ajcc 8": "AJCC_8_Stage",
        "Sex": "Gender",
        "Thy (1-5)": "Thy_1_5",
        "Neoadjuvant Therapy (Radiotherapy/Systemic Therapy)": "NeoadjuvantTx_Radiotherapy_SystemicTherapy",
        "Primary Tumour Block 2": "Primary_SpecBlockNum2",
        "Gastric Her2 Fish": "Gastric_HER2_FISH",
        "S2 1.5 To 10 X (645 - 4300)": "S2_1.5_To_10X_645_4300",
        "S Stage Ldh (S0 = Normal": "S_Stage_LDH_S0EqNormal",
        "Castrate Resistant Disease At Time Of Sample Surgery": "Castrate_Resistant_Disease_At_Time_Sample_Surgery",
        "Metastatic Disease At Diagnosis (Distant mets/Non-Regional Lymph Nodes)": "Met_Disease_at_DX_distMets_nonRegionalLN",
        "Metastatic Disease At Time Of Sample Surgery (Distant Mets/Non-Regional Lymph Nodes)": "Met_Disease_At_Time_Sample_Surgery_distMets_nonRegionalLN",
        "Tte Progression To Metastatic Disease (From Date Of Diagnosis)": "TTE_Progression_To_Met_Disease_From_DateDX",
        "Disease Status At Sample Surgery (Metastatic/Non-Metastatic/Castrate Resistant/Hormone Sensitive)": "Disease_Status_At_Sample_Surgery_Met_NonMet_CRPC_HSPC",
        "Date Progression To Metastatic Disease": "Date_Progression_To_Metastatic_Disease"
    }

    # Loop through the columns with similar names and rename them
    for i in range(1, 8):
        column_name_map[f"Anonymised Id Matched Met {i}"] = (
            f"AnonymisedID_MatchedMet{i}"
        )
        column_name_map[f"Matched Met {i} Block Number"] = f"MatchedMet{i}_BlockNumber"
        column_name_map[f"Matched Met {i} Site"] = f"MatchedMet{i}_Site"
        column_name_map[f"Matched Met {i} Surgery Date"] = f"MatchedMet{i}_SurgeryDate"
        column_name_map[f"Matched Met {i} Treatment Subgrouping"] = (
            f"MatchedMet{i}_TreatmentSubgrouping"
        )
        column_name_map[f"Tte Matched Met {i} Surgery Date"] = (
            f"TTE_MatchedMet{i}_SurgeryDate"
        )

    for i in range(1, 10):
        column_name_map[f"Turbt {i} Date"] = f"TURBT_{i}_Date"
        column_name_map[f"Turbt {i} Pt Stage"] = f"TURBT_{i}_pT_Stage"
        column_name_map[f"Tte Turbt {i}"] = f"TTE_TURBT_{i}"

    for i in range(1, 11):
        column_name_map[f"Neoadjuvant Tx{i}"] = f"NeoadjuvantTx{i}"
        column_name_map[f"Neoadjuvant Tx{i} Agent"] = f"NeoadjuvantTx{i}_Agent"
        column_name_map[f"Neoadjuvant Tx{i} Start Date"] = f"NeoadjuvantTx{i}_StartDate"
        column_name_map[f"Neoadjuvant Tx{i} End Date"] = f"NeoadjuvantTx{i}_EndDate"
        column_name_map[f"Tumour Response Neoadjuvant Tx{i}"] = (
            f"TumourResponse_NeoadjuvantTx{i}"
        )
        column_name_map[f"Tte Neoadjuvant Tx{i} Start Date"] = (
            f"TTE_NeoadjuvantTx{i}StartDate"
        )
        column_name_map[f"Tte Neoadjuvant Tx{i} End Date"] = (
            f"TTE_NeoadjuvantTx{i}EndDate"
        )

    for i in range(1, 16):
        column_name_map[f"Adjuvant Tx{i}"] = f"AdjuvantTx{i}"
        column_name_map[f"Adjuvant Tx{i} Agent"] = f"AdjuvantTx{i}_Agent"
        column_name_map[f"Adjuvant Tx{i} Start Date"] = f"AdjuvantTx{i}_StartDate"
        column_name_map[f"Adjuvant Tx{i} End Date"] = f"AdjuvantTx{i}_EndDate"
        column_name_map[f"Tumour Response Adjuvant Tx{i}"] = (
            f"TumourResponse_AdjuvantTx{i}"
        )
        column_name_map[f"Subsequent Tx{i}"] = f"SubsequentTx{i}"
        column_name_map[f"Subsequent Tx{i} Agent"] = f"SubsequentTx{i}_Agent"
        column_name_map[f"Subsequent Tx{i} Start Date"] = f"SubsequentTx{i}_StartDate"
        column_name_map[f"Subsequent Tx{i} End Date"] = f"SubsequentTx{i}_EndDate"
        column_name_map[f"Subsequent Tx{i} Finish Date"] = f"SubsequentTx{i}_EndDate"
        column_name_map[f"Tumour Response To Subsequent Tx{i}"] = (
            f"TumourResponse_SubsequentTx{i}"
        )
        column_name_map[f"Sites Recurrence/ Progression {i}"] = (
            f"Sites_RecurrenceProgression{i}"
        )
        column_name_map[f"Tte Date Recurrence/ Progression {i}"] = (
            f"TTE_DateRecurrenceProgression{i}"
        )
        column_name_map[f"Tte Adjuvant Tx{i} Start Date"] = (
            f"TTE_AdjuvantTx{i}StartDate"
        )
        column_name_map[f"Tte Adjuvant Tx{i} End Date"] = f"TTE_AdjuvantTx{i}EndDate"
        column_name_map[f"Tte Subsequent Tx{i} Start Date"] = (
            f"TTE_SubsequentTx{i}StartDate"
        )
        column_name_map[f"Tte Subsequent Tx{i} End Date"] = (
            f"TTE_SubsequentTx{i}EndDate"
        )
        column_name_map[f"Date Recurrence/ Progression {i}"] = (
            f"Date_RecurrenceProgression{i}"
        )
        column_name_map[f"Tte Subsequent Tx{i} Finish Date"] = (
            f"TTE_SubsequentTx{i}EndDate"
        )

    for i in range(1, 16):
        column_name_map[f"Tte Date Recurrence/ Progression {i}"] = (
            f"TTE_DateRecurrenceProgression{i}"
        )

    for i in range(1, 16):
        column_name_map[f"Tte Subsequent Tx{i} Start Date"] = (
            f"TTE_SubsequentTx{i}StartDate"
        )
        column_name_map[f"Tte Subsequent Tx{i} End Date"] = (
            f"TTE_SubsequentTx{i}EndDate"
        )
        column_name_map[f"Tte Subsequent Tx{i} Finish Date"] = (
            f"TTE_SubsequentTx{i}EndDate"
        )

    for i in range(1, 3):
        column_name_map[f"Low Grade Dysplasia Block {i}"] = (
            f"LowGrade_Dysplasia_Block{i}"
        )
        column_name_map[f"Low Grade Dysplasia Block {i} Site"] = (
            f"LowGrade_Dysplasia_Block{i}_Site"
        )
        column_name_map[f"High Grade Dysplasia Block {i}"] = (
            f"HighGrade_Dysplasia_Block{i}"
        )
        column_name_map[f"High Grade Dysplasia Block {i} Site"] = (
            f"HighGrade_Dysplasia_Block{i}_Site"
        )
        column_name_map[f"FocalHigh Grade Dysplasia Block {i}"] = (
            f"FocalHighGrade_Dysplasia_Block{i}"
        )
        column_name_map[f"FocalHigh Grade Dysplasia Block {i} Site"] = (
            f"FocalHighGrade_Dysplasia_Block{i}_Site"
        )
        column_name_map[f"Focal High Grade Dysplasia Block {i}"] = (
            f"FocalHighGrade_Dysplasia_Block{i}"
        )
        column_name_map[f"Focal High Grade Dysplasia Block {i} Site"] = (
            f"FocalHighGrade_Dysplasia_Block{i}_Site"
        )
        column_name_map[f"High Grade Bilin Block {i}"] = f"High_Grade_Bilin_Block{i}"
        column_name_map[f"High Grade Bilin Block {i} Site"] = (
            f"High_Grade_Bilin_Block{i}_Site"
        )
        column_name_map[f"High Grade Dysplasia {i} Site"] = (
            f"High_Grade_Dysplasia{i}_Site"
        )
        column_name_map[f"High Grade Dysplasia Block Number {i}"] = (
            f"High_Grade_Dysplasia_BlockNumber{i}"
        )
        column_name_map[f"Intracholecystic Papillary Neoplasia Block {i}"] = (
            f"Intracholecystic_Papillary_Neoplasia_Block{i}"
        )
        column_name_map[f"Intracholecystic Papillary Neoplasia Block {i} Site"] = (
            f"Intracholecystic_Papillary_Neoplasia_Block{i}_Site"
        )
        column_name_map[f"Intraductal Papillary Neoplasia Block {i}"] = (
            f"Intraductal_Papillary_Neoplasia_Block{i}"
        )
        column_name_map[f"Intraductal Papillary Neoplasia Block {i} Site"] = (
            f"Intraductal_Papillary_Neoplasia_Block{i}_Site"
        )
        column_name_map[f"Matched Normal Tissue Block {i} Site"] = (
            f"MatchedNormalTissue_Block{i}_Site"
        )
        column_name_map[f"Matched Normal Tissue Block {i}"] = (
            f"MatchedNormalTissue_Block{i}"
        )
        column_name_map[f"Matched Normal Tissue Block Number {i}"] = (
            f"MatchedNormalTissue_Block{i}"
        )
        column_name_map[f"Micropapillary Lesion Block {i}"] = (
            f"Micropapillary_Lesion_Block{i}"
        )
        column_name_map[f"Micropapillary Lesion Block {i} Site"] = (
            f"Micropapillary_Lesion_Block{i}_Site"
        )
        column_name_map[f"Mixed Low And High Grade Bilin Block {i}"] = (
            f"Mixed_Low_High_Grade_Bilin_Block{i}"
        )
        column_name_map[f"Mixed Low And High Grade Bilin Block {i} Site"] = (
            f"Mixed_Low_High_Grade_Bilin_Block{i}_Site"
        )
        column_name_map[f"Papillary Lesion Block {i}"] = f"Papillary_Lesion_Block{i}"
        column_name_map[f"Papillary Lesion Block {i} Site"] = (
            f"Papillary_Lesion_Block{i}_Site"
        )
        column_name_map[f"Polypoid Mucinous Cystic Neoplasm Block {i}"] = (
            f"Polypoid_Mucinous_Cystic_Neoplasm_Block{i}"
        )
        column_name_map[f"Polypoid Mucinous Cystic Neoplasm Block {i} Site"] = (
            f"Polypoid_Mucinous_Cystic_Neoplasm_Block{i}_Site"
        )
        column_name_map[f"Ungraded Biliary Intraepithelial Neoplasia Block {i}"] = (
            f"Ungraded_Biliary_Intraepithelial_Neoplasia_Block{i}"
        )
        column_name_map[
            f"Ungraded Biliary Intraepithelial Neoplasia Block {i} Site"
        ] = f"Ungraded_Biliary_Intraepithelial_Neoplasia_Block{i}_Site"
        column_name_map[f"Ungraded Dyplasia {i} Site"] = f"Ungraded_Dyplasia{i}_Site"
        column_name_map[f"Ungraded Dyplasia Block Number {i}"] = (
            f"Ungraded_Dyplasia_Block{i}"
        )

    for i in range(1, 6):
        column_name_map[f"Additional Metastatic Specimen Number {i}"] = (
            f"Additional_Metastatic_SpecimenNumber{i}"
        )
        column_name_map[f"Site {i}"] = f"Site{i}"

    for i in range(1, 4):
        column_name_map[f"Mucinous Cystic Neoplasm Block {i}"] = (
            f"Mucinous_Cystic_Neoplasm_Block{i}"
        )
        column_name_map[f"Mucinous Cystic Neoplasm Block {i} Site"] = (
            f"Mucinous_Cystic_Neoplasm_Block{i}_Site"
        )
        column_name_map[f"Last Treatment Prior To Specimen {i} Surgery"] = (
            f"Last_Treatment_Prior_To_Spec{i}_Surgery"
        )
        column_name_map[f"Post Treatment Specimen {i} Anonymised Id"] = (
            f"Post_Treatment_Spec{i}_Anonymised_ID"
        )

    return column_name_map


# ----------------------------------------------------------------------------------------------------------------------------------------


# define a function to set the data types of the patient list datasheet
def set_data_types_patientList(main_workbook):
    """
    This function maps the patient list columns to the data types in the database
    """
    # Define the data types for the Patient List Data
    data_types_PatientList = {
        "DonorID": CHAR(10),
        "SpecimenNumber": VARCHAR(15),
        "NHS_Number": VARCHAR(20),
        "NumTumourBlocksSentToTriStar": INTEGER,
        "NumNormalBlocksSnetToTriStar": INTEGER,
        "TotalNumBlocksSentToTriStar": INTEGER,
        "DateBlocksSent": VARCHAR(50),
        "ManifestID": CHAR(10),
        "CohortName": VARCHAR(200),
        "FileName": VARCHAR(200),
        "AllBlocksReturnedFromTriStar": BIT,
        "NumBlocksReturnedFromTriStar": INTEGER,
        "DateBlocksReturnedFromTriStar": DATE,
        "TrackingNumber": VARCHAR(30),
        "BlockReturnComments": VARCHAR(200),
        "PatientNumber": INTEGER,
        "DiseaseEpisode": VARCHAR(20),
        "Recall": BIT,
        "SentToTriStar": BIT,
        "Tissue_Pathway": VARCHAR(100),
        "CohortID": VARCHAR(30),
        "Comments": VARCHAR(300),
        "CohortName_Mapped": VARCHAR(200),
    }
    return data_types_PatientList


# ----------------------------------------------------------------------------------------------------------------------------------------


def set_data_types_tristar_pathway(main_workbook):
    """
    This function maps the tristar pathway columns to the data types in the database
    """
    # Define the data types of the tristar pathway data
    data_types_tristar_pathway = {
        "Cohort": VARCHAR(100),
        "Cohort_ID": VARCHAR(75),
        "Cohort_Details": VARCHAR(300),
        "Manifest": VARCHAR(75),
        "Shipment_Date": DATE,
        "Tracking_Number": VARCHAR(100),
        "Data_Sent": VARCHAR(50),
        "Version": VARCHAR(30),
        "Folder": VARCHAR(10),
        "MTA": VARCHAR(10),
        "New_Name": VARCHAR(300),
        "Holding_Name": VARCHAR(200),
        "Box_Updated": DATE,
        "Current_Name": VARCHAR(300),
    }
    return data_types_tristar_pathway


# ----------------------------------------------------------------------------------------------------------------------------------------


# Define a Function to set the data types of all columns in scan logs
def set_data_types_scan_logs(main_workbook):
    """
    This function maps the scan logs columns to the data types in the database
    """
    # Define the data types
    data_types = {
        "LH_Number": VARCHAR(50),
        "Part_Number": VARCHAR(50),
        "Anonymised_Code": VARCHAR(50),
        "Svs_Number_To_Transfer": VARCHAR(50),
        "Scan1_QC": VARCHAR(50),
        "Scan2_QC": VARCHAR(50),
        "ErrorCode_Scan1": VARCHAR(50),
        "ErrorCode_Scan2": VARCHAR(50),
        "Notes_Scan1": VARCHAR(50),
        "Notes_Scan2": VARCHAR(50),
        "Indication": VARCHAR(50),
    }
    return data_types


# ----------------------------------------------------------------------------------------------------------------------------------------


# define a function to create the dictionary for renaming the columns in the patient list
def map_cohort_names_patientList(main_workbook):
    """
    This function adds the CohortName_Mapped column to a dataframe using the cohort mapping dictionary
    """
    # Add the cohort names mapped column
    main_workbook["CohortName_Mapped"] = main_workbook["CohortName"].map(cohort_mapping)
    return main_workbook


# -----------------------------------------------------------------------------------------------------------------------------------------


def add_cohortName_column():
    """
    This function adds the cohort name column to a dataset to be uploaded to the database - after retrieving
    the selection from the user
    """
    # Set the cohort names variable as the keys() in the cohort mapping dictionary - which would be the options list in the dropdown
    cohort_names = sorted(cohort_mapping.keys())

    # Add the cohort names to the dropdown menu
    dropdown = widgets.Dropdown(
        options=cohort_names,
        value=None,
        description="Please select the cohort: ",
        disabled=False,
        layout=widgets.Layout(width="450px"),
        style={"description_width": "initial"},
    )

    # Define a variable to store the selected cohort
    selected_cohort = {"value": dropdown.value}

    # observe the widget for changes in the selection and update accordingly
    def on_change(change):
        if change["type"] == "change" and change["name"] == "value":
            selected_cohort["value"] = change["new"]
            print(f"You selected the cohort named {selected_cohort['value']}.")

    dropdown.observe(on_change, names="value")
    display(dropdown)

    return selected_cohort
