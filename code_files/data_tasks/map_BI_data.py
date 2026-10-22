# Import the required libraries
import re

import numpy as np
import pandas as pd
from openpyxl import load_workbook

pd.set_option("display.max_rows", None)
pd.set_option("display.max_columns", None)

# Load the BI data workbook
wb = load_workbook(
    filename="C:/Users/medaamm/Box/Data file (AA)/Metadata projects/Boehringer Ingelheim/BI sample data (AA).xlsx"
)

# Define the set of table names which need to be selected
tbls_to_select = {
    "leeds_fu1",
    "ovarian_500_fu1",
    "leeds_basic1",
    "leeds_supplementary",
    "hgugm_fu1",
    "hgugm_fu2",
    "hgugm_fu3",
    "hgugm_basic1",
    "hannover_fu1",
    "chu_nice_fu1",
    "nhsb_fu1",
    "nhsb_fu2",
    "nhsb_basic1",
    "nhsb_basic2",
    "nhsb_basic3",
    "iom_fu1",
    "iom_fu2",
    "iom_fu3",
    "iom_fu4",
    "iom_fu5",
    "iom_fu6",
    "iom_basic1",
    "iom_basic2",
    "iom_basic3",
}


# Define a function to process the date columns
def process_dates(series):
    # Standardize to string and trim whitespace
    s_str = series.astype(str).str.strip()

    # Parse valid date/time strings cleanly
    parsed = pd.to_datetime(series, errors="coerce", utc=True, format="mixed")
    formatted = parsed.dt.strftime("%Y-%m-%d")

    # Keep parsed date if successful; otherwise revert to original raw string
    return formatted.fillna(s_str)


# Extract target tables directly into mapped dictionary
tables_mapped = {}

for ws in wb.worksheets:
    for table_name, table in ws.tables.items():
        if table_name in tbls_to_select:
            rows = [[cell.value for cell in row] for row in ws[table]]
            df = pd.DataFrame(rows[1:], columns=rows[0])
            tables_mapped[f"{table_name}_mapped"] = df

# format the date columns in the non-leeds data
for df_name, df in tables_mapped.items():
    for col in df.columns:
        if "date" in col.lower():
            tables_mapped[df_name][col] = process_dates(df[col])


column_mapping = {
    "Identifiers_and_Metadata": {
        "Anonymized_ID": "Patient ID",
        "Blocks and Slide ID": "Vendor Sample ID",
        "DonorID": "Patient ID",
        "Order Number": "Vendor Order ID",
        "Patient_ID": "Patient ID",
        "Sample ID": "Vendor Sample ID",
        "Sample_ID": "Vendor Sample ID",
        "Biobank_ID": "Patient ID",
    },
    "Demographics_and_History": {
        "Age": "Patient Age at diagnosis",
        "AgeAtDX": "Patient Age at diagnosis",
        "AgeDX": "Patient Age at diagnosis",
        "Age_at_Diagnosis": "Patient Age at diagnosis",
        "ConcurrentMedicalConditions": "Comorbidities",
        "Concurrent_Diseases": "Comorbidities",
        "Concurrent_Medical_Conditions": "Comorbidities",
        "Gender": "Gender",
        "Smoking_History": "Smoking history (Descr.)",
        "Smoking_Status": "Tobacco History",
        "Viral_Status": None,
        "Ethnicity": None,
    },
    "Diagnosis_and_Classification": {
        "Cancer_Status": "Tissue description",
        "New_Primary_Or_Recurrence_SpecimenNature": "Tissue description",
        "Site": "Organ_Combined",
        "Site_Of_Diagnosis": "Organ_Combined",
        "Site_of_Diagnosis": "Organ_Combined",
        "TumourSite": "Organ_Combined",
        "TumourSite_Location": "Organ_Combined",
        "Tumour_Site_Area": "Organ_2",
        "Side_L_R": "Organ_1",
    },
    "Staging_Grading_and_Scores": {"StagingCriteria": "TNM Edition"},
    "Specimen_Sample_and_QC": {
        "Block_Type": "Tissue description",
        "DescriptionTumour_MacroscopicTumourType_GrossDescription": "Vendor tissue size (mm)",
        "Divided": "Vendor tissue size (mm)",
        "Format": "Sample Type",
        "QC_Results_Necrotic_Tissue_Perc": "Vendor Necrosis %",
        "QC_Results_TumourTissue_Perc": "Vendor Tumor %",
        "QC_Results_Tumour_Size": "Vendor tissue size (mm)",
        "Sample_Format": "Sample_Type",
        "Sample_Type": "Sample_Type",
        "Specimen1_Type": "Type of Surgery",
        "Specimen_1_Type": "Sample_Type",
        "Specimen_Nature": "Tissue description",
        "Specimen_Type": "Tissue description",
        "Surgical_Specimen": "Organ_Combined",
        "TumourInvasion_Extent": "Vendor Tumor %",
        "Tumour_Necrosis": "Vendor Necrosis %",
        "Tumour_Size": "Vendor tissue size (mm)",
    },
    "Biomarkers_Genetics_and_IHC": {
        "BRAF_Mutation": None,
        "CDKN2A_Loss": None,
        "Chromosome_19Q13_Loss": None,
        "Chromosome_1P36_Loss": None,
        "EGFR_Amplification": None,
        "EGFR_Mutation": None,
        "ER": "Estrogen Receptor",
        "HER2_IHC": "HER2",
        "Histone_H3F3A_Mutation": None,
        "Histone_HIST1H3B_Mutation": None,
        "IDH2_Mutation": None,
        "KRAS_Mutation": None,
        "MGMT_Promoter_Methylation": None,
        "MMR_Result": "MSI Status",
        "NGS_IDH1_Mutation": None,
        "NRAS_Mutation": None,
        "PR": "Progesteron Receptor",
        "TERT_Promoter_Mutation": None,
    },
    "Treatments_and_Procedures": {
        "Surgery_Site": "Organ_Combined",
    },
    "Disease_Progression_and_Metastasis": {
        "Lymph_Nodes_Num_Assessed": "TNM: total nodes",
        "Lymph_Nodes_Num_Positive": "TNM: Nodes positive",
    },
    "Dates_and_Outcomes": {
        "CauseOfDeath": "Cause of Death",
        "Cause_Of_Death": "Cause of Death",
        "Cause_of_Death": "Cause of Death",
        "Collection_Date": "Collection Date",
        "DateDX": "Date of Diagnosis",
        "Date_Of_Collection": "Collection Date",
        "Date_Of_Diagnosis": "Date of Diagnosis",
        "Date_Of_Histological_Diagnosis": "Date of Diagnosis",
        "Date_Of_Surgery": "Collection Date",
        "Date_of_Diagnosis": "Date of Diagnosis",
        "Date_of_Surgery": "Collection Date",
        "Death": "Date of Death",
        "Death_Date": "Date of Death",
        "Month_Year_Initial_Diagnosis_PrimaryCancer": "Date of Diagnosis",
        "Month_Year_Surgery_SampleCollection": "Collection Date",
        "Specimen_1_Collection_Date": "Collection Date",
        "SurgeryDate": "Collection Date",
        "Year_Of_Collection": "Collection Date",
    },
    "Administrative_and_Notes": {
        "Comments": "Comments",
    },
}

exclude_columns = [
    "Site_ID",
    "Project_Code",
    "Purchase_Order",
    "Study_Name",
    "Age_At_DX",
    "Geographic_Origin",
    "Symptoms",
    "Case_Number",
    "Histopathology_Subtype",
    "AJCC_pT_Or_Other_Applicable",
    "Bulky",
    "Differentiation",
    "International_Prognostic_Index",
    "Nuclear_Grade",
    "S_Stage_AFP",
    "StagingGradingInfoProvidedFor",
    "Summary_S_Stage",
    "Block_A",
    "Block_B",
    "Block_C",
    "Block_D",
    "Block_E",
    "Block_F",
    "Block_G",
    "Block_H",
    "Block_Number",
    "Block_Used_For_LynchScreening_MMR_IHC",
    "Block_Used_For_Testing",
    "Block_Used_For_Testing_P16",
    "ID_Format",
    "Line_Shipping_List",
    "Met_Block_1",
    "Met_Block_2",
    "Met_Block_3",
    "Met_Block_4",
    "Position_In_Box",
    "ResectionMarginsClear",
    "SpecNumber_Used_For_MolecularTesting_IfDifferent",
    "TRK_Incoming",
    "IHC_IDH1_Mutation",
    "Immuno_Histopathologic_Data",
    "Immunohistochemistry_KI67_Expression",
    "KI67",
    "KI_67",
    "MLH1_IHC_Pos_Neg_ClonalNeg",
    "MSH2_IHC_Pos_Neg_ClonalNeg",
    "MSH6_IHC_Pos_Neg_ClonalNeg",
    "Mitotic_Count",
    "P16_PosNeg",
    "PDL1_Antibody",
    "PDL1_Percent",
    "PDL1_PosNeg",
    "PMS2_IHC_Pos_Neg_ClonalNeg",
    "PSA_At_Diagnosis_ugPerLitre",
    "PSA_At_Last_FollowUp_ugPerLitre_DetectionLimitsVary",
    "PSA_Nadir_ugPerLitre_DetectionLimitsVary",
    "Percent_Cells_Pos",
    "Serum_AFP_At_DX_KuPerL_RefLessThan7KuPerL",
    "Serum_LDH_At_DX_Ref_BW_160_430_IuPerL",
    "Serum_Total_Pregnancy_HGC_At_DX_RefLessThan5UPerL",
    "Androgen_Deprivation_Therapy",
    "AnyImmunotherapy",
    "AnyNeoadjuvantTherapy",
    "Castrate_Resistant_Prostate_Cancer",
    "Chemo_pre_or_post_surgery",
    "Immuno_Met_Tx",
    "Maintenance_Therapy_With_Rituximab",
    "N_Cycles",
    "Second_Treatment",
    "SpecSurgery",
    "Surgery_Site_Side",
    "Surgery_and_Site",
    "Thy_1_5",
    "Transplantation",
    "TreatmentSubgrouping",
    "Treatment_Schedule",
    "mCRPC_mHSPC",
    "Biochemical_Relapse_During_ADT",
    "Bone_Marrow",
    "Clinical_Relapse_During_ADT",
    "Extranodal_Localization",
    "Lymph_Nodes_Localization",
    "Met_Disease_1",
    "Met_Disease_2",
    "Metastatic_Site_Distant_Status",
    "Nodal_Mets",
    "Responder_NonResponder",
    "ResponseToNeoTherapy_TRScore",
    "DFS",
    "Molecular_Date",
    "OS",
    "OS_Months",
    "PET_At_Diagnosis",
    "PFS_Months",
    "Period_Between_Disease_Progression_Months",
    "Survival",
    "Survival_Months",
    "Survival_Status",
    "TTE_AdjuvantTx1EndDate",
    "TTE_AdjuvantTx1StartDate",
    "TTE_AdjuvantTx2EndDate",
    "TTE_AdjuvantTx2StartDate",
    "TTE_AdjuvantTx3EndDate",
    "TTE_AdjuvantTx3StartDate",
    "TTE_AdjuvantTx4EndDate",
    "TTE_AdjuvantTx4StartDate",
    "TTE_CRPC_Biochemical_From_ADT_StartDate",
    "TTE_CRPC_Biochemical_From_DateDiagnosis",
    "TTE_DateLastSeen_OS",
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
    "TTE_Date_LastFollowUp",
    "TTE_Death_OS_PFS",
    "TTE_NeoadjuvantTx1EndDate",
    "TTE_NeoadjuvantTx1StartDate",
    "TTE_NeoadjuvantTx2EndDate",
    "TTE_NeoadjuvantTx2StartDate",
    "TTE_NeoadjuvantTx3EndDate",
    "TTE_NeoadjuvantTx3StartDate",
    "TTE_OverallSurvival_Months",
    "TTE_PFS_Months",
    "TTE_PrimarySurgeryDate",
    "TTE_Recurrence_PFS",
    "TTE_SubsequentTx1EndDate",
    "TTE_SubsequentTx1StartDate",
    "TTE_SubsequentTx2EndDate",
    "TTE_SubsequentTx2StartDate",
    "TTE_SubsequentTx3EndDate",
    "TTE_SubsequentTx3StartDate",
    "TTE_SubsequentTx4EndDate",
    "TTE_SubsequentTx4StartDate",
    "TTE_SubsequentTx5EndDate",
    "TTE_SubsequentTx5StartDate",
    "TTE_SurgeryDate",
    "CohortName",
    "Details_Provided_For",
    "Molecular_Testing_Grading_Info_Provided_For",
    "Tristar ID #",
    "Original ID",
    "Original_ID",
    "Pathology_ID",
]

# Drop excluded columns safely
for df_name, df in tables_mapped.items():
    cols_to_drop = [col for col in exclude_columns if col in df.columns]
    if cols_to_drop:
        tables_mapped[df_name] = df.drop(columns=cols_to_drop)


# Helper for string combination
def combine_columns(df, columns):
    existing_cols = [col for col in columns if col in df.columns]
    if not existing_cols:
        return pd.Series("", index=df.index)
    return df[existing_cols].apply(
        lambda x: " | ".join(
            s
            for s in x.dropna().astype(str).str.strip()
            if s not in ("", "nan", "None", "NR")
        ),
        axis=1,
    )


def combine_with_column_name(df, columns):
    existing_cols = [col for col in columns if col in df.columns]
    if not existing_cols:
        return pd.Series("", index=df.index)
    return df[existing_cols].apply(
        lambda x: " | ".join(
            f"{col}: {val}"
            for col, val in zip(existing_cols, x)
            if pd.notnull(val) and str(val).strip() not in ("", "nan", "None", "NR")
        ),
        axis=1,
    )


# 1. Identifiers and Metadata Processing
identifiers_mapping = {
    k: v for k, v in column_mapping["Identifiers_and_Metadata"].items() if v
}
for df_name, df in tables_mapped.items():
    df = df.rename(columns=identifiers_mapping)
    if "Patient ID" not in df.columns and "Vendor Sample ID" in df.columns:
        df["Patient ID"] = df["Vendor Sample ID"]
        cols = list(df.columns)
        cols.remove("Patient ID")
        insert_pos = cols.index("Vendor Sample ID") + 1
        cols.insert(insert_pos, "Patient ID")
        df = df[cols]
    tables_mapped[df_name] = df

# 2. Demographics & Clinical History
demographics_mapping = {
    k: v for k, v in column_mapping["Demographics_and_History"].items() if v
}
gender_replace = {
    "-": "NR",
    2: "Female",
    "MALE": "Male",
    "FEMALE": "Female",
    "F": "Female",
    "M": "Male",
}
tobacco_replace = {
    "Smoker": "Current Use",
    "Ex-smoker": "Previous Use",
    "Non-smoker": "Never Used",
    "Not stated": "Unknown",
}
age_columns = ["Age", "AgeAtDX", "AgeDX", "Age_at_Diagnosis"]

for df_name, df in tables_mapped.items():
    df = df.rename(columns=demographics_mapping)
    if "Gender" in df.columns:
        df["Gender"] = df["Gender"].replace(gender_replace).astype(str)
    if "Tobacco History" in df.columns:
        df["Tobacco History"] = (
            df["Tobacco History"].replace(tobacco_replace).fillna("Unknown").astype(str)
        )
    df["Patient Age at diagnosis"] = combine_columns(df, age_columns)
    df = df.drop(columns=[col for col in age_columns if col in df.columns])
    tables_mapped[df_name] = df

# 3. Diagnostics & Cancer Classification
diagnosis_columns = [
    "CancerType_Indication",
    "CancerType_Indication_Diagnosis",
    "Diagnosis",
    "General_Cancer_Subtypes",
    "General_Subtypes_Malignant_Neoplasm_MaleGenitalOrgans_ICD_Code",
    "General_Subtypes_Malignant_Neoplasm_MesothelialAndSoftTissue_IC",
    "Histological_Diagnosis",
    "Histological_Subtype",
    "Pathological_Diagnosis",
    "SpecHistologicalDiagnosis",
    "Sub-Type",
    "Sub_Type",
    "Type",
    "Histological_Tumour_Type",
    "Histological_Tumour_Type_WHO",
    "Gross_Description",
    "Histopathologic_Subtype",
]

diagnosis_indication_mapping = {
    "ADC | Colon Pri. And Met.": "Colorectal Cancer (CRC)",
    "ADC | Colon Pri. and Met.": "Colorectal Cancer (CRC)",
    "ADC | Gastric Pri.": "Gastric Cancer",
    "ADC | Gastric Pri. and Nodal Met.": "Gastric Cancer",
    "ADC | Lung Pri.": "Lung Cancer (NSCLC)",
    "ADC | Pancreas Pri.": "Pancreatic ductal adenocarcinoma (PDAC)",
    "ADC | Prostate Pri.": "Prostate Cancer",
    "ADC/ Metasynchronous | Colon Pri. And Met. (Liver)": "Colorectal Cancer (CRC)",
    "ADC/ Synchronous | Colon Pri. And Met. (Liver)": "Colorectal Cancer (CRC)",
    "Adenocarcinoma of the gastro-oesophageal junction": "Gastroesophageal Junction Adenocarcinoma",
    "Adenocarcinoma, intestinal type": "Gastric Cancer",
    "Adenocarcinoma, intestinal type (background Helicobacter gastritis)": "Gastric Cancer",
    "Adenocarcinoma, mixed diffuse and intestinal type": "Gastric Cancer",
    "Angiosarcoma": "Sarcoma",
    "Breast Pri. | IDC": "Breast Cancer",
    "Breast Pri. | ILC": "Breast Cancer",
    "C22.0 Malignant neoplasm: Liver cell carcinoma | Hepatocellular Carcinoma | Hepatocellular carcinoma | Multiple nodose tumors": "Hepatocellular Carcinoma (HCC)",
    "C22.0 Malignant neoplasm: Liver cell carcinoma | Hepatocellular Carcinoma | Hepatocellular carcinoma | N.A.": "Hepatocellular Carcinoma (HCC)",
    "C22.0 Malignant neoplasm: Liver cell carcinoma | Hepatocellular Carcinoma | Hepatocellular carcinoma | Nodose tumoe": "Hepatocellular Carcinoma (HCC)",
    "C22.0 Malignant neoplasm: Liver cell carcinoma | Hepatocellular Carcinoma | Hepatocellular carcinoma | Nodose tumor": "Hepatocellular Carcinoma (HCC)",
    "C25 Malignant neoplasm of pancreas | Ductal adenocarcinoma | Nodose tumor": "Pancreatic ductal adenocarcinoma (PDAC)",
    "C25.0 Malignant neoplasm: Head of pancreas | Ductal adenocarcinoma partially mucinous | Nodose tumor": "Pancreatic ductal adenocarcinoma (PDAC)",
    "C25.0 Malignant neoplasm: Head of pancreas | Ductal adenocarcinoma | Infiltrative tumor": "Pancreatic ductal adenocarcinoma (PDAC)",
    "C25.0 Malignant neoplasm: Head of pancreas | Ductal adenocarcinoma | Nodose tumor": "Pancreatic ductal adenocarcinoma (PDAC)",
    "C25.0 Malignant neoplasm: Head of pancreas | Ductal adenocarcinoma | Solid-cystic tumor": "Pancreatic ductal adenocarcinoma (PDAC)",
    "C25.1 Malignant neoplasm: Body of pancreas | Ductal adenocarcinoma | Nodose tumor": "Pancreatic ductal adenocarcinoma (PDAC)",
    "C25.2 Malignant neoplasm: Tail of pancreas | Ductal adenocarcinoma | Nodose tumor": "Pancreatic ductal adenocarcinoma (PDAC)",
    "C61 Malignant neoplasm of prostate | Adenocarcinoma": "Prostate Cancer",
    "CDI - (NST) | Breast Pri.": "Breast Cancer",
    "CLL": "Lymphoma",
    "Chromophobe renal cell carcinoma with rhabdoid differentiation": "Renal Cell Carcinoma",
    "Classical Seminoma": "Testicular cancer",
    "Clear cell  Carcinoma | Ovary Pri.": "Ovarian Cancer",
    "Conventional clear cell renal carcinoma with sarcomatoid differentiation": "Renal Cell Carcinoma",
    "Conventional renal cell carcinoma": "Renal Cell Carcinoma",
    "Conventional renal clear cell carcinoma": "Renal Cell Carcinoma",
    "Cystic papillary thyroid carcinoma": "Thyroid Cancer",
    "DIFFUSE GASTRIC ADC | GASTRIC CARCINOMA": "Gastric Cancer",
    "DLBCL | Lymph Node Pri.": "Lymphoma",
    "Diffuse type adenocarcinoma": "Gastric Cancer",
    "Diffuse type adenocarcinoma with signet ring cells": "Gastric Cancer",
    "Diffuse type gastric adenocarcinoma": "Gastric Cancer",
    "Diffuse-type gastric adenocarcinoma": "Gastric Cancer",
    "Extrapulmonary high grade neuroendocrine carcinoma": "Extra-pulmonary neuroendocrine carcinoma (epNEC)",
    "FL | Colon Pri.": "Lymphoma",
    "FL | Lymph Node Pri.": "Lymphoma",
    "Fibromyxoid sarcoma": "Sarcoma",
    "GASTRIC ADC INTESTINAL TYPE | GASTRIC CARCINOMA": "Gastric Cancer",
    "GASTRIC ADC SIGNET RING CELL | GASTRIC CARCINOMA": "Gastric Cancer",
    "GASTRIC ADC TUBULAR TYPE | GASTRIC CARCINOMA": "Gastric Cancer",
    "GASTRIC ADC | GASTRIC CARCINOMA": "Gastric Cancer",
    "Gastric Pri. and Met. (Lymph.) | ADC": "Gastric Cancer",
    "Gastric Pri. | ADC": "Gastric Cancer",
    "Glioblastoma": "Brain Cancer - Glioblastoma",
    "Glioblastoma (primary tumour anaplastic astrocytoma)": "Brain Cancer - Glioblastoma",
    "Glioblastoma (primary tumour diffuse astrocytoma WHO grade II)": "Brain Cancer - Glioblastoma",
    "Glioblastoma (primary tumour grade III anaplastic astrocytoma)": "Brain Cancer - Glioblastoma",
    "Hepatocellular carcinoma": "Hepatocellular Carcinoma (HCC)",
    "Hepatocellular carcinoma and cholangiocarcinoma": "Mixed indications",
    "Hurthle cell carcinoma": "Thyroid Cancer",
    "Intestinal type adenocarcinoma": "Gastric Cancer",
    "Intraductal tubopapillary neoplasm with invasive adenocarcinoma": "Pancreatic cancer (other)",
    "Intrahepatic cholangiocarcinoma": "Cholangiocarcinoma",
    "Invasive UC | Bladder Pri.": "Bladder Cancer/Urothelial Carcinoma",
    "MUCINOUS GASTRIC ADC | GASTRIC CARCINOMA": "Gastric Cancer",
    "Malignant astrocytoma | C71.9 Malignant neoplasm: Brain, unspecified - Glioblastoma | Gliobastoma": "Brain Cancer - Glioblastoma",
    "Medullary thyroid carcinoma": "Thyroid Cancer",
    "Merkel cell carcinoma": "Skin cancer (other)",
    "Metastasis": "",
    "Metastatic hepatocellular carcinoma": "Hepatocellular Carcinoma (HCC)",
    "Mixed intestinal and diffuse adenocarcinoma": "Gastric Cancer",
    "Myxofibrosarcoma": "Sarcoma",
    "Nodular malignant melanoma": "Melanoma",
    "Normal tissue": "",
    "ORAL CAVITY CANCER | EPIDERMOID CARCINOMA": "Head and Neck Squamous Cell Carcinoma (HNSCC)",
    "PAPILLARY GASTRIC ADC | GASTRIC CARCINOMA": "Gastric Cancer",
    "Pancreas Pri. and Met. | ADC": "Pancreatic ductal adenocarcinoma (PDAC)",
    "Papillary Serous  Carcinoma | Ovary Pri.": "Ovarian Cancer",
    "Papillary renal cell carcinoma": "Renal Cell Carcinoma",
    "Papillary thyroid carcinoma": "Thyroid Cancer",
    "Papillary thyroid carcinoma (and Hurthle cell follicular adenoma)": "Thyroid Cancer",
    "Papillary thyroid carcinoma (follicular variant)": "Thyroid Cancer",
    "Perhilar cholangiocarcinoma": "Cholangiocarcinoma",
    "Perihilar cholangiocarcinoma": "Cholangiocarcinoma",
    "Poorly differentiated adenocarcinoma of the GOJ": "Gastroesophageal Junction Adenocarcinoma",
    "Poorly differentiated adenocarcinoma of the cardio-oesophageal junction": "Gastroesophageal Junction Adenocarcinoma",
    "Prostate cancer": "Prostate Cancer",
    "Prostate cancer | C61 Malignant neoplasm of prostate | Acinar adenocarcinoma | N.A.": "Prostate Cancer",
    "Prostate cancer | C61 Malignant neoplasm of prostate | Adenocarcinoma": "Prostate Cancer",
    "Prostatic adenocarcinoma": "Prostate Cancer",
    "Recurrent hepatocellular carcinoma": "Hepatocellular Carcinoma (HCC)",
    "Renal cell carcinoma": "Renal Cell Carcinoma",
    "Renal cell carcinoma, unclassified (with clear cell areas)": "Renal Cell Carcinoma",
    "UROTELIAL CARCINOMA | BLADDER CARCINOMA": "Bladder Cancer/Urothelial Carcinoma",
    "lung cancer | Adenocarcinoma": "Lung Cancer (NSCLC)",
}

vendor_disease_term_cleaned = {
    "": "NR",
    "ADC | Gastric Pri.": "Gastric Adenocarcinoma",
    "ADC | Lung Pri.": "Lung Adenocarcinoma",
    "Hepatocellular adenomas": "Hepatocellular Adenoma",
    "Mucinous": "Mucinous Carcinoma",
    "C25.1 Malignant neoplasm: Body of pancreas | Ductal adenocarcinoma | Nodose tumor": "Pancreatic (Body) Ductal Adenocarcinoma with Nodular Morphology",
    "ADC | Pancreas Pri.": "Pancreatic Ductal Adenocarcinoma",
    "Invasive UC | Bladder Pri.": "Invasive Urothelial Carcinoma",
    "C25.2 Malignant neoplasm: Tail of pancreas | Ductal adenocarcinoma | Nodose tumor": "Pancreatic (Tail) Ductal Adenocarcinoma with Nodular Morphology",
    "Gastric Pri. | ADC": "Gastric Adenocarcinoma",
    "Mucinous G2": "Mucinous Carcinoma",
    "ADC/ Synchronous | Colon Pri. And Met. (Liver)": "Primary Colon Adenocarcinoma with Synchronous Hepatic Metastasis",
    "PAPILLARY GASTRIC ADC | GASTRIC CARCINOMA": "Papillary Gastric Adenocarcinoma",
    "ADC | Colon Pri. and Met.": "Colon Adenocarcinoma with Metastatic Disease",
    "C22.0 Malignant neoplasm: Liver cell carcinoma | Hepatocellular Carcinoma | Hepatocellular carcinoma | Multiple nodose tumors": "Hepatocellular Carcinoma with Nodular Morphology",
    "Renal cell carcinoma, unclassified (with clear cell areas)": "Renal Cell Carcinoma with Clear Cell Areas",
    "C22.0 Malignant neoplasm: Liver cell carcinoma | Hepatocellular Carcinoma | Hepatocellular carcinoma | N.A.": "Hepatocellular Carcinoma",
    "GASTRIC ADC TUBULAR TYPE | GASTRIC CARCINOMA": "Tubular Gastric Adenocarcinoma",
    "Adenocarcinoma, diffuse type": "Diffuse-Type Adenocarcinoma",
    "GASTRIC ADC | GASTRIC CARCINOMA": "Gastric Adenocarcinoma",
    "C25.0 Malignant neoplasm: Head of pancreas | Ductal adenocarcinoma | Infiltrative tumor": "Infiltrative Pancreatic (Head) Ductal Adenocarcinoma",
    "Diffuse type gastric adenocarcinoma": "Diffuse-Type Gastric Adenocarcinoma",
    "Papillary thyroid carcinoma (and Hurthle cell follicular adenoma)": "Papillary Thyroid Carcinoma and Hurthle Cell Follicular Adenoma",
    "Adenocarcinoma of the gastro-oesophageal junction": "Gastroesophageal Junction Adenocarcinoma",
    "Pancreas Pri. and Met. | ADC": "Pancreatic Ductal Adenocarcinoma with Metastatic Disease",
    "Breast Pri. | ILC": "Breast Invasive Lobular Carcinoma",
    "Poorly differentiated adenocarcinoma of the GOJ": "Poorly Differentiated Gastroesophageal Junction Adenocarcinoma",
    "C25.0 Malignant neoplasm: Head of pancreas | Ductal adenocarcinoma | Nodose tumor": "Pancreatic (Head) Ductal Adenocarcinoma with Nodular Morphology",
    "C22.0 Malignant neoplasm: Liver cell carcinoma | Hepatocellular Carcinoma | Hepatocellular carcinoma | Nodose tumoe": "Hepatocellular Carcinoma with Nodular Morphology",
    "Metastasis": "Metastatic Tumor",
    "Neuroendocrine tumour (well differentiated)": "Well Differentiated Neuroendocrine Tumor",
    "Adenocarcinoma, mixed type (NOS)": "Mixed Type Adenocarcinoma (NOS)",
    "C61 Malignant neoplasm of prostate | Adenocarcinoma": "Prostate Adenocarcinoma",
    "Adenocarcinoma, intestinal type": "Intestinal-Type Adenocarcinoma",
    "Gastric Pri. and Met. (Lymph.) | ADC": "Gastric Adenocarcinoma with Lymph Node Metastasis",
    "FL | Colon Pri.": "Primary Colonic Follicular Lymphoma",
    "ADC | Colon Pri. And Met.": "Colon Adenocarcinoma with Metastatic Disease",
    "Adenocarcinoma, mixed diffuse and intestinal type": "Mixed Diffuse and Intestinal-Type Adenocarcinoma",
    "Ductal ADC": "Ductal Adenocarcinoma",
    "ADC/ Metasynchronous | Colon Pri. And Met. (Liver)": "Metachronous Colon Adenocarcinoma with Liver Metastasis",
    "C25 Malignant neoplasm of pancreas | Ductal adenocarcinoma | Nodose tumor": "Pancreatic Ductal Adenocarcinoma with Nodular Morphology",
    "CDI - (NST) | Breast Pri.": "Invasive Ductal Carcinoma (NST)",
    "ORAL CAVITY CANCER | EPIDERMOID CARCINOMA": "Oral Cavity Squamous Cell Carcinoma",
    "Diffuse-type gastric adenocarcinoma": "Diffuse-Type Gastric Adenocarcinoma",
    "Clear cell": "Clear Cell Carcinoma",
    "ADC | Gastric Pri. and Nodal Met.": "Gastric Adenocarcinoma with Lymph Node Metastasis",
    "FL | Lymph Node Pri.": "Primary Nodal Follicular Lymphoma",
    "CLL": "Chronic Lymphocytic Leukemia",
    "Adenocarcinoma, mixed (mucinous with a diffuse growth pattern)": "Mixed Mucinous Adenocarcinoma with a Diffuse Growth Pattern",
    "DIFFUSE GASTRIC ADC | GASTRIC CARCINOMA": "Diffuse-Type Gastric Adenocarcinoma",
    "Parathyroid": "Parathyroid Carcinoma",
    "UROTELIAL CARCINOMA | BLADDER CARCINOMA": "Urothelial Carcinoma",
    "Normal tissue": "No Tumor (Normal Tissue)",
    "MUCINOUS GASTRIC ADC | GASTRIC CARCINOMA": "Mucinous Gastric Adenocarcinoma",
    "C25.0 Malignant neoplasm: Head of pancreas | Ductal adenocarcinoma | Solid-cystic tumor": "Pancreatic (Head) Ductal Adenocarcinoma with Solid-Cystic Morphology",
    "Prostate cancer | C61 Malignant neoplasm of prostate | Adenocarcinoma": "Prostate Adenocarcinoma",
    "Papillary Serous  Carcinoma | Ovary Pri.": "Primary Ovarian Serous Papillary Carcinoma",
    "GASTRIC ADC SIGNET RING CELL | GASTRIC CARCINOMA": "Signet Ring Cell Gastric Adenocarcinoma",
    "DLBCL | Lymph Node Pri.": "Primary Nodal Diffuse Large B-Cell Lymphoma",
    "Clear cell  Carcinoma | Ovary Pri.": "Primary Ovarian Clear Cell Carcinoma",
    "Adenocarcinoma, intestinal type (background Helicobacter gastritis)": "Intestinal-Type Gastric Adenocarcinoma with Background Helicobacter Gastritis",
    "Breast Pri. | IDC": "Invasive Ductal Carcinoma",
    "C25.0 Malignant neoplasm: Head of pancreas | Ductal adenocarcinoma partially mucinous | Nodose tumor": "Pancreatic (Head) Partially Mucinous Ductal Adenocarcinoma with Nodular Morphology",
    "C22.0 Malignant neoplasm: Liver cell carcinoma | Hepatocellular Carcinoma | Hepatocellular carcinoma | Nodose tumor": "Hepatocellular Carcinoma with Nodular Morphology",
    "Adenocarcinoma (NOS; variable histology)": "Adenocarcinoma (NOS)",
    "ADC": "Adenocarcinoma",
    "Papillary thyroid carcinoma (follicular variant)": "Follicular Variant Papillary Thyroid Carcinoma",
    "Glioblastoma (primary tumour grade III anaplastic astrocytoma)": "Glioblastoma",
    "Malignant astrocytoma | C71.9 Malignant neoplasm: Brain, unspecified - Glioblastoma | Gliobastoma": "Glioblastoma",
    "ADC | Prostate Pri.": "Primary Prostate Adenocarcinoma",
    "Prostate cancer | C61 Malignant neoplasm of prostate | Acinar adenocarcinoma | N.A.": "Prostate Acinar Adenocarcinoma",
    "Glioblastoma (primary tumour diffuse astrocytoma WHO grade II)": "Glioblastoma",
    "GASTRIC ADC INTESTINAL TYPE | GASTRIC CARCINOMA": "Intestinal-Type Gastric Adenocarcinoma",
    "lung cancer | Adenocarcinoma": "Lung Adenocarcinoma",
    "Poorly differentiated adenocarcinoma of the cardio-oesophageal junction": "Poorly Differentiated Cardio-Oesophageal Junction Adenocarcinoma",
    "Glioblastoma (primary tumour anaplastic astrocytoma)": "Glioblastoma",
}

crc_ids = {
    "LDS0000253",
    "LDS0000254",
    "LDS0000258",
    "LDS0000259",
    "LDS0000260",
    "LDS0000262",
    "LDS0000269",
    "LDS0000270",
    "LDS0000271",
    "LDS0000272",
    "LDS0000273",
    "LDS0000274",
    "LDS0000275",
    "LDS0000277",
    "LDS0000280",
    "LDS0000281",
    "LDS0000282",
    "LDS0000285",
    "LDS0000288",
    "LDS0000289",
    "LDS0000290",
    "LDS0000292",
    "LDS0000295",
    "LDS0000296",
    "LDS0000299",
    "LDS0000300",
    "LDS0000308",
    "LDS0000309",
    "LDS0000313",
    "LDS0000314",
    "LDS0000316",
    "LDS0000317",
    "LDS0000318",
    "LDS0000319",
    "LDS0000321",
    "LDS0000322",
    "LDS0000323",
    "LDS0000328",
    "LDS0000329",
    "LDS0000330",
    "LDS0000331",
    "LDS0000334",
    "LDS0000335",
    "LDS0000337",
    "LDS0000340",
    "LDS0000342",
    "LDS0000346",
    "LDS0000349",
    "LDS0000352",
    "LDS0000353",
    "LDS0000355",
    "LDS0000358",
    "LDS0000362",
    "LDS0000363",
    "LDS0000365",
    "LDS0000366",
    "LDS0000382",
    "LDS0000369",
    "LDS0000370",
    "LDS0000373",
    "LDS0000374",
    "LDS0000377",
    "LDS0000379",
    "LDS0000380",
    "LDS0000383",
    "LDS0000386",
    "LDS0000387",
    "LDS0000388",
    "LDS0000389",
}
hnscc_ids = {
    "LDS0001773",
    "LDS0001774",
    "LDS0001906",
    "LDS0001908",
    "LDS0001922",
    "LDS0001926",
    "LDS0001934",
    "LDS0001981",
    "LDS0001982",
    "LDS0001983",
    "LDS0001986",
    "LDS0001987",
    "LDS0001988",
    "LDS0001989",
    "LDS0001990",
    "LDS0001991",
    "LDS0001992",
    "LDS0001993",
    "LDS0001994",
    "LDS0001997",
    "LDS0002002",
    "LDS0002003",
    "LDS0002004",
    "LDS0002005",
    "LDS0000825",
    "LDS0000826",
    "LDS0000830",
    "LDS0000831",
    "LDS0000832",
    "LDS0000834",
    "LDS0000835",
    "LDS0000836",
    "LDS0000837",
    "LDS0000839",
    "LDS0000841",
    "LDS0000842",
    "LDS0000843",
    "LDS0000846",
    "LDS0000847",
    "LDS0000848",
    "LDS0000851",
    "LDS0000852",
    "LDS0000853",
    "LDS0000854",
    "LDS0000855",
    "LDS0000858",
    "LDS0000859",
    "LDS0000860",
    "LDS0000861",
    "LDS0000862",
    "LDS0000863",
    "LDS0000864",
    "LDS0000865",
    "LDS0000866",
    "LDS0000867",
    "LDS0000868",
    "LDS0000870",
    "LDS0000871",
    "LDS0000872",
    "LDS0000874",
    "LDS0000875",
    "LDS0000876",
    "LDS0000877",
    "LDS0000878",
    "LDS0000879",
    "LDS0000880",
    "LDS0000881",
    "LDS0000882",
    "LDS0000883",
    "LDS0000884",
    "LDS0000887",
    "LDS0000888",
    "LDS0000891",
    "LDS0000892",
    "LDS0000894",
    "LDS0000895",
    "LDS0000897",
    "LDS0000899",
    "LDS0000900",
    "LDS0000901",
    "LDS0000902",
}
gallbladder_ids = {
    "LDS0000658",
    "LDS0000659",
    "LDS0000660",
    "LDS0000661",
    "LDS0000662",
    "LDS0000663",
    "LDS0000664",
    "LDS0000665",
    "LDS0000666",
    "LDS0000668",
    "LDS0000669",
    "LDS0000670",
    "LDS0000671",
}
gastric_ids = {
    "LDS0001946",
    "LDS0001947",
    "LDS0001952",
    "LDS0001953",
    "LDS0001956",
    "LDS0001958",
    "LDS0001959",
    "LDS0000759",
    "LDS0000767",
    "LDS0000684",
    "LDS0000689",
    "LDS0000691",
    "LDS0000692",
    "LDS0000703",
    "LDS0000705",
    "LDS0000719",
    "LDS0000722",
    "LDS0000729",
}
panc_other_ids = {
    "LDS0002920",
    "LDS0002921",
    "LDS0002922",
    "LDS0002923",
    "LDS0002924",
    "LDS0002925",
    "LDS0002927",
    "LDS0002928",
    "LDS0002929",
    "LDS0002911",
    "LDS0002912",
    "LDS0002913",
    "LDS0002918",
    "LDS0002931",
    "LDS0002932",
    "LDS0002933",
    "LDS0002934",
    "276P040-E",
    "276P040-M2",
}
gej_ids = {
    "LDS0000675",
    "LDS0000678",
    "LDS0000683",
    "LDS0000693",
    "LDS0000694",
    "LDS0000697",
    "LDS0000704",
    "LDS0000712",
    "LDS0000713",
    "LDS0000714",
    "LDS0000715",
    "LDS0000720",
    "LDS0000725",
    "LDS0000728",
    "LDS0000732",
}
pdac_ids = {
    "LDS0001181",
    "LDS0001183",
    "LDS0001191",
    "LDS0001194",
    "LDS0001196",
    "LDS0001204",
    "LDS0001205",
    "LDS0001206",
    "LDS0001207",
    "LDS0001208",
    "LDS0001210",
    "LDS0001214",
    "LDS0001219",
    "276P001-E",
    "276P048-F",
}
prostate_ids = {
    "LDS0001393",
    "LDS0001432",
    "LDS0001479",
    "LDS0001486",
    "LDS0001490",
    "LDS0001520",
    "LDS0001526",
    "LDS0001530",
}
ovarian_ids = {"LDS-0003", "LDS-0025", "LDS-0138"}
normal_ids = {
    "LDS0004682",
    "LDS0004683",
    "LDS0004687",
    "LDS0004705",
    "LDS0004706",
    "LDS0004707",
    "LDS0004708",
    "LDS0004709",
}
internal_prostate_ids = {
    8149,
    8098,
    8099,
    8110,
    8144,
    8106,
    8103,
    8104,
    8145,
    8107,
    8108,
}

cols_to_drop = ["Primary_Tumour_Site", "PrimTumourSite"]

for df_name, df in tables_mapped.items():
    existing_cols = [col for col in diagnosis_columns if col in df.columns]
    if existing_cols:
        raw_combined = (
            df[existing_cols]
            .fillna("")
            .replace("NR", "", regex=False)
            .astype(str)
            .agg(" | ".join, axis=1)
        )
        df["Vendor Disease Term"] = raw_combined.str.split(r"\s*\|\s*").apply(
            lambda vals: " | ".join(dict.fromkeys(v.strip() for v in vals if v.strip()))
        )
        df = df.drop(columns=existing_cols)

        df["Tumor indication"] = df["Vendor Disease Term"].map(
            diagnosis_indication_mapping
        )

        if "Patient ID" in df.columns:
            pid = df["Patient ID"]
            df.loc[pid.isin(crc_ids), "Tumor indication"] = "Colorectal Cancer (CRC)"
            df.loc[pid.isin(hnscc_ids), "Tumor indication"] = (
                "Head and Neck Squamous Cell Carcinoma (HNSCC)"
            )
            df.loc[pid.isin(gallbladder_ids), "Tumor indication"] = "Gallbladder cancer"
            df.loc[pid.isin(gastric_ids), "Tumor indication"] = "Gastric Cancer"
            df.loc[pid.isin(panc_other_ids), "Tumor indication"] = (
                "Pancreatic cancer (other)"
            )
            df.loc[pid.isin(gej_ids), "Tumor indication"] = (
                "Gastroesophageal Junction Adenocarcinoma"
            )
            df.loc[pid == "LDS0000917", "Tumor indication"] = (
                "Hepatocellular Carcinoma (HCC)"
            )
            df.loc[
                pid.isin(["LDS0004672", "LDS0004676", "LDS0004679"]), "Tumor indication"
            ] = ""
            df.loc[pid.isin(pdac_ids), "Tumor indication"] = (
                "Pancreatic ductal adenocarcinoma (PDAC)"
            )
            df.loc[pid.isin(prostate_ids), "Tumor indication"] = "Prostate Cancer"
            df.loc[pid == "LDS000236", "Tumor indication"] = "Renal Cell Carcinoma"
            df.loc[pid.isin(ovarian_ids), "Tumor indication"] = "Ovarian Cancer"

            normal_mask = pid.isin(normal_ids)
            df.loc[normal_mask, "Tissue description"] = "Normal"
            df.loc[normal_mask, "Tumor indication"] = "N/A"

        if "Internal ID" in df.columns:
            df.loc[
                df["Internal ID"].isin(internal_prostate_ids), "Tumor indication"
            ] = "Prostate Cancer"

        df["Vendor Disease Term"] = df["Vendor Disease Term"].map(
            lambda x: vendor_disease_term_cleaned.get(
                x, x.title() if isinstance(x, str) else x
            )
        )

    df = df.rename(columns=column_mapping["Diagnosis_and_Classification"])
    df = df.drop(columns=[c for c in cols_to_drop if c in df.columns])
    tables_mapped[df_name] = df

# 4. Organ & Sublocalization Mapping Logic
if "leeds_supplementary_mapped" in tables_mapped:
    ls_df = tables_mapped["leeds_supplementary_mapped"]
    if "Organ_1" in ls_df.columns and "Organ_2" in ls_df.columns:
        ls_df["Organ_Combined"] = ls_df["Organ_1"].combine_first(ls_df["Organ_2"])
        tables_mapped["leeds_supplementary_mapped"] = ls_df.drop(
            columns=["Organ_1", "Organ_2"]
        )

organs_mapping = {
    "Addominal wall": "Soft tissue",
    "Anterior maxilla": "Oral cavity",
    "Antrum": "Stomach",
    "Ascending colon": "Colon",
    "Base of tongue": "Tongue",
    "Base of tongue mucosa/left anterior neck soft": "Multiorgan",
    "Bilateral (apex, anterior)": "Prostate",
    "Bilateral (posterior, lateral)": "Prostate",
    "Body of stomach": "Stomach",
    "Brain, Left temporal lobe": "Brain",
    "Brain, Right Occipital": "Brain",
    "Brain, Right Parietal lobe": "Brain",
    "Brain, Right frontal": "Brain",
    "Brain, Right medial temporal": "Brain",
    "Brain, Right parietal-temporal": "Brain",
    "Brain, left anterior temporal": "Brain",
    "Brain, left frontal": "Brain",
    "Brain, left occipital": "Brain",
    "Brain, left temporal": "Brain",
    "Brain, left temporo-parietal": "Brain",
    "Brain, right frontal": "Brain",
    "Brain, right frontal temporal": "Brain",
    "Brain, right parietal": "Brain",
    "Brain, right temporal": "Brain",
    "Buccal mucosa": "Oral cavity",
    "Caecum": "Colon",
    "Cardia": "Stomach",
    "Cardia (note tumour spreads into GOJ but position of tumour epicentre not recorded; No Siewert-Stein classification)": "Stomach",
    "Cardia - extends into and GOJ and distal oesophagus; no tumour epicentre measurements recorded or Siewert-Stein classification given": "Multiorgan",
    "Cardia - undermines mucosa at GOJ; no site of tumour epicentre measurements recorded or Siewer-Stein classification given": "Multiorgan",
    "Central pancreas": "Pancreas",
    "Descending": "Colon",
    "Distal pancreas": "Pancreas",
    "Distal stomach": "Stomach",
    "Distal stomach and first part of duodenum": "Multiorgan",
    "Facial node": "Lymph node",
    "Floor of mouth": "Oral cavity",
    "Fundus of stomach": "Stomach",
    "Gastro-oesophageal junction": "Multiorgan",
    "Gastro-oesophageal junction (Siewert type I)": "Multiorgan",
    "Gastro-oesophageal junction, NOS (Siewert I/II)": "Multiorgan",
    "Greater curve": "Stomach",
    "HGUGM": "Oral cavity",
    "Hepatic flexure": "Colon",
    "Isthmus": "Thyroid",
    "Just above oesophagus distally": "Esophagus",
    "L": "Colon",
    "L Breast": "Breast",
    "L Ear": "Skin",
    "L Leg": "Soft tissue",
    "Left": "Colon",
    "Left (Liver)": "Liver",
    "Left (Omentum-Peritoneum-Liver)": "Multiorgan",
    "Left brachialis": "Skeletal muscle",
    "Left breast": "Breast",
    "Left buccal": "Oral cavity",
    "Left cheek": "Skin",
    "Left chest wall (local recurrence)": "Soft tissue",
    "Left ear": "Skin",
    "Left ear (local recurrence)": "Skin",
    "Left extended low maxillectomy": "Oral cavity",
    "Left hemiglossectomy": "Tongue",
    "Left kidney": "Kidney",
    "Left lateral oropharyngectomy": "Oropharynx",
    "Left leg": "Soft tissue",
    "Left level 1 node": "Lymph node",
    "Left level 2a node": "Lymph node",
    "Left lobe": "Thyroid",
    "Left lower leg": "Soft tissue",
    "Left lung": "Lung",
    "Left mandible": "Oral cavity",
    "Left maxilla": "Oral cavity",
    "Left oropharyngectomy": "Oropharynx",
    "Left oropharynx": "Oropharynx",
    "Left ovary ": "Ovary",
    "Left posterior hemiglossectomy": "Tongue",
    "Left retromolar": "Oral cavity",
    "Left retromolar trigone": "Oral cavity",
    "Left ring finger": "Skin",
    "Left soft palate": "Oral cavity",
    "Left testis": "Testis",
    "Left tongue": "Tongue",
    "Left tonsil": "Tonsil",
    "Left upper back": "Skin",
    "Level 1B node": "Lymph node",
    "Level 2 + 3 node": "Lymph node",
    "Level 2 node": "Lymph node",
    "Level 2A + 2B node": "Lymph node",
    "Level 2A node": "Lymph node",
    "Level 2B node": "Lymph node",
    "Level 3 node": "Lymph node",
    "Level 4 node": "Lymph node",
    "Liver segment, 4-8": "Liver",
    "Liver segment, 7": "Liver",
    "Liver segment, 8": "Liver",
    "Liver, Segment 2, 3": "Liver",
    "Liver, Segment 2,3,4 and 5": "Liver",
    "Liver, Segment 3": "Liver",
    "Liver, Segment 3, 4": "Liver",
    "Liver, Segment 4": "Liver",
    "Liver, Segment 5, 6": "Liver",
    "Liver, Segment 5/4b": "Liver",
    "Liver, Segment 6": "Liver",
    "Liver, Segment 6,7": "Liver",
    "Liver, Segment 8": "Liver",
    "Liver, caudate lobe, and segments 2, 4, 5, 7 & 8": "Liver",
    "Liver, segment 2, 3": "Liver",
    "Liver, segment 3": "Liver",
    "Liver, segment 4": "Liver",
    "Liver, segment 4b": "Liver",
    "Liver, segment 5": "Liver",
    "Liver, segment 5 & 7": "Liver",
    "Liver, segment 6": "Liver",
    "Liver, segment 6, 7": "Liver",
    "Liver, segment 6, 7 & 8": "Liver",
    "Liver, segment 7/8": "Liver",
    "Liver, segment 8": "Liver",
    "Liver, segment, 3,4,5,6,8": "Liver",
    "Liver, segments 2 & 3": "Liver",
    "Liver, segments 2, 3, 4, 5, 6 & 7": "Liver",
    "Liver, segments 3 & 5": "Liver",
    "Liver, segments 4,5 & 8": "Liver",
    "Liver, segments 4b, 5, 6 and gallbladder": "Multiorgan",
    "Liver, segments 6 & 7": "Liver",
    "Liver, segments 8 & 5": "Liver",
    "Liver. Segment 4,8": "Liver",
    "Liver. Segment 8": "Liver",
    "Lower back (above sacrum)": "Skin",
    "Lower oesophagus/GOJ": "Multiorgan",
    "Lung, left upper lobe": "Lung",
    "Lymph Node": "Lymph node",
    "Lymph node": "Lymph node",
    "Lymph node/ satellite tumour deposit": "Lymph node",
    "Lymph node/satellite tumour deposit": "Lymph node",
    "Lymph nodes": "Lymph node",
    "NR": "",
    "Node": "Lymph node",
    "Node (neck)": "Lymph node",
    "Nose": "Skin",
    "Oesophagus": "Esophagus",
    "Oesophagus - lower": "Esophagus",
    "Oesophagus - lower (just above GOJ)": "Esophagus",
    "Oesophagus, distal": "Esophagus",
    "Oesophagus, distal & GOJ": "Multiorgan",
    "Oesophagus, mid-lower": "Esophagus",
    "One lymph node": "Lymph node",
    "Pancreas (distal)": "Pancreas",
    "Pancreas (head)": "Pancreas",
    "Pancreas (main duct)": "Pancreas",
    "Pancreas (tail)": "Pancreas",
    "Pancreas (tail); infiltrating kidney (B10)": "Multiorgan",
    "Para-oesophageal lymph nodes": "Lymph node",
    "Paraoesophageal lymph node": "Lymph node",
    "Perigastric lymph node": "Lymph node",
    "Perigastric nodes": "Lymph node",
    "Perihilar region": "Lung",
    "Post auricular left": "Skin",
    "Posterior pharyngeal wall": "Pharynx",
    "Pyloric antrum": "Stomach",
    "Pylorus": "Stomach",
    "R": "Colon",
    "R Atrium": "Heart",
    "R Breast": "Breast",
    "R Tongue": "Tongue",
    "R extended hemiglossectomy": "Tongue",
    "R soft tissue and thyroid": "Multiorgan",
    "R upper back": "Skin",
    "Rectosigmoid colon": "Colon",
    "Rectum": "Colon",
    "Right": "Colon",
    "Right (Peritoneum)": "Mesothelium",
    "Right Kidney": "Kidney",
    "Right Scalp": "Skin",
    "Right anterior floor of mouth and partial glossectomy": "Multiorgan",
    "Right anterior tongue": "Tongue",
    "Right back": "Skin",
    "Right breast": "Breast",
    "Right chest wall": "Soft tissue",
    "Right ethmoidal sinus": "Nasopharynx",
    "Right floor of mouth tumour": "Oral cavity",
    "Right inferolateral leg": "Soft tissue",
    "Right knee": "Soft tissue",
    "Right lateral oropharyngectomy": "Oropharynx",
    "Right lateral tongue": "Tongue",
    "Right level 2a node": "Lymph node",
    "Right liver": "Liver",
    "Right lobe": "Thyroid",
    "Right lung": "Lung",
    "Right neck level 2A to 5": "Lymph node",
    "Right oropharyngectomy": "Oropharynx",
    "Right oropharynx": "Oropharynx",
    "Right pharyngeal wall (main tumour in supraglottis although tumour and high grade dysplasia also found in R pharyngeal wall and tongue base)": "Multiorgan",
    "Right posterior mandible": "Oral cavity",
    "Right posterior tongue": "Tongue",
    "Right temple": "Skin",
    "Right testis": "Testis",
    "Right tongue": "Tongue",
    "Right tongue base": "Tongue",
    "Right tonsil": "Tonsil",
    "Right tonsillar": "Tonsil",
    "Right undescended testis": "Testis",
    "Right upper chest": "Skin",
    "Right upper nose and ethmoid": "Nasopharynx",
    "Satellite tumour": "Adipose Tissue",
    "Satellite tumour deposit": "Adipose Tissue",
    "Satellite tumour deposit (pericolic fat)": "Adipose Tissue",
    "Satellite tumour deposit in mesocolon": "Adipose Tissue",
    "Satellite tumour deposit, pericolic fat": "Adipose Tissue",
    "Sigma": "Colon",
    "Sigma (Liver)": "Multiorgan",
    "Sigmoid colon": "Colon",
    "Skin, Right facial subcutaneous": "Skin",
    "Soft palate": "Oral cavity",
    "Splenic flexure": "Colon",
    "Stomach; involves right colon": "Multiorgan",
    "Tail of pancreas": "Pancreas",
    "Tissue around GOJ": "Multiorgan",
    "Tongue base": "Tongue",
    "Tonsil tumour": "Tonsil",
    "Tonsil, lateral pharynx": "Multiorgan",
    "Transverse (proximal third)": "Colon",
    "Transverse (proximal)": "Colon",
    "Transverse colon": "Colon",
    "Tumour": "Stomach",
    "Vocal cord (tumour originated in L pyriform fossa)": "Multiorgan",
    "lung": "Lung",
    "mandibular tuberosity": "Oral cavity",
    "right pterygoid/maxillary": "Oral cavity",
}


sublocalization_mapping = {
    "Body of stomach": "Body",
    "Brain, Left temporal lobe": "Left Temporal Lobe",
    "Brain, Right Occipital": "Right Occipital Lobe",
    "Brain, Right Parietal lobe": "Right Parietal Lobe",
    "Brain, Right frontal": "Right Frontal Lobe",
    "Brain, Right medial temporal": "Right Medial Temporal",
    "Brain, Right parietal-temporal": "Right Parietal Temporal",
    "Brain, left anterior temporal": "Left Anterior Temporal",
    "Brain, left frontal": "Left Frontal Lobe",
    "Brain, left occipital": "Left Occipital Lobe",
    "Brain, left temporal": "Left Temporal Lobe",
    "Brain, left temporo-parietal": "Left Tempo Parietal Lobe",
    "Brain, right frontal": "Right Frontal Lobe",
    "Brain, right frontal temporal": "Right Frontal Temporal Lobe",
    "Brain, right parietal": "Right Parietal Lobe",
    "Brain, right temporal": "Right Temporal Lobe",
    "Cardia (note tumour spreads into GOJ but position of tumour epicentre not recorded; No Siewert-Stein classification)": "Cardia and extends to GOJ",
    "Cardia - extends into and GOJ and distal oesophagus; no tumour epicentre measurements recorded or Siewert-Stein classification given": "Cardia and extends to GOJ",
    "Cardia - undermines mucosa at GOJ; no site of tumour epicentre measurements recorded or Siewer-Stein classification given": "Cardia and undermines Mucosa at GOJ",
    "Central pancreas": "Central Pancreas",
    "Descending": "Descending Colon",
    "Distal pancreas": "Distal Pancreas",
    "Distal stomach": "Distal Stomach",
    "Distal stomach and first part of duodenum": "Distal Stomach and first part of Duodenum",
    "Facial node": "Facial node",
    "Floor of mouth": "Floor of Mouth",
    "Fundus of stomach": "Stomach Fundus",
    "Gastro-oesophageal junction": "Gastro-Esophageal Junction",
    "Gastro-oesophageal junction (Siewert type I)": "Gastro-Esophageal Junction",
    "Gastro-oesophageal junction, NOS (Siewert I/II)": "Gastro-Esophageal Junction",
    "Greater curve": "Stomach Greater Curve",
    "HGUGM": "NR",
    "Just above oesophagus distally": "Just above the Oesophagus Distally",
    "L": "Left Colon",
    "L Breast": "Left Breast",
    "L Ear": "Left Ear",
    "L Leg": "Left Leg",
    "Left": "Left Colon",
    "Left (Liver)": "Left Liver",
    "Left (Omentum-Peritoneum-Liver)": "Left Omentum, Peritoneum, and Liver",
    "Left brachialis": "Left Brachialis",
    "Left breast": "Left Breast",
    "Left buccal": "Left Buccal",
    "Left cheek": "Left Cheek",
    "Left chest wall (local recurrence)": "Left Chest Wall",
    "Left ear": "Left Ear",
    "Left ear (local recurrence)": "Left Ear",
    "Left extended low maxillectomy": "Left Maxilla",
    "Left hemiglossectomy": "Left Tongue",
    "Left kidney": "Left Kidney",
    "Left lateral oropharyngectomy": "Left Lateral Oropharynx",
    "Left leg": "Left Leg",
    "Left lobe": "Left Lobe",
    "Left lower leg": "Left Lower Leg",
    "Left lung": "Left Lung",
    "Left mandible": "Left Mandible",
    "Left maxilla": "Left Maxilla",
    "Left oropharyngectomy": "Left Oropharynx",
    "Left oropharynx": "Left Oropharynx",
    "Left ovary ": "Left Ovary",
    "Left posterior hemiglossectomy": "Left Tongue",
    "Left retromolar": "Oral cavity",
    "Left retromolar trigone": "Oral cavity",
    "Level 1B node": "Lymph node",
    "Level 2 + 3 node": "Lymph node",
    "Level 2 node": "Lymph node",
    "Level 2A + 2B node": "Lymph node",
    "Level 2A node": "Lymph node",
    "Level 2B node": "Lymph node",
    "Level 3 node": "Lymph node",
    "Level 4 node": "Lymph node",
    "Liver segment, 4-8": "Segment 4-8",
    "Liver segment, 7": "Segment 7",
    "Liver segment, 8": "Segment 8",
    "Liver, Segment 2, 3": "Segment 2,3",
    "Liver, Segment 2,3,4 and 5": "Segment 2-5",
    "Liver, Segment 3": "Segment 3",
    "Liver, Segment 3, 4": "Segment 3,4",
    "Liver, Segment 4": "Segment 4",
    "Liver, Segment 5, 6": "Segment 5,6",
    "Liver, Segment 5/4b": "Segment 4b/5",
    "Liver, Segment 6": "Segment 6",
    "Liver, Segment 6,7": "Segment 6,7",
    "Liver, Segment 8": "Segment 8",
    "Liver, caudate lobe, and segments 2, 4, 5, 7 & 8": "Segment 1,2,4,5,7,8",
    "Liver, segment 2, 3": "Segment 2,3",
    "Liver, segment 3": "Segment 3",
    "Liver, segment 4": "Segment 4",
    "Liver, segment 4b": "Segment 4b",
    "Liver, segment 5": "Segment 5",
    "Liver, segment 5 & 7": "Segment 5,7",
    "Liver, segment 6": "Segment 6",
    "Liver, segment 6, 7": "Segment 6,7",
    "Liver, segment 6, 7 & 8": "Segment 6-8",
    "Liver, segment 7/8": "Segment 7,8",
    "Liver, segment 8": "Segment 8",
    "Liver, segment, 3,4,5,6,8": "Segment 3-6,8",
    "Liver, segments 2 & 3": "Segment 2,3",
    "Liver, segments 2, 3, 4, 5, 6 & 7": "Segment 2-7",
    "Liver, segments 3 & 5": "Segment 3,5",
    "Liver, segments 4,5 & 8": "Segment 4,5,8",
    "Liver, segments 4b, 5, 6 and gallbladder": "Segment 4b,5,6",
    "Liver, segments 6 & 7": "Segment 6,7",
    "Liver, segments 8 & 5": "Segment 5,8",
    "Liver. Segment 4,8": "Segment 4,8",
    "Liver. Segment 8": "Segment 8",
    "Lower back (above sacrum)": "Lower Back (Above Sacrum)",
    "Lower oesophagus/GOJ": "Lower Esophagus - Gastro-Esophageal Junction",
    "Lung, left upper lobe": "Left Upper Lobe",
    "Lymph Node": "NR",
    "Lymph node": "NR",
    "Lymph node/ satellite tumour deposit": "Lymph Node - Satellite Tumor Deposit",
    "Lymph node/satellite tumour deposit": "Lymph Node - Satellite Tumor Deposit",
    "Lymph nodes": "NR",
    "NR": "NR",
    "Node (neck)": "Neck Lymph Node",
    "Oesophagus": "NR",
    "Oesophagus - lower": "Lower Esophagus",
    "Oesophagus - lower (just above GOJ)": "Lower Esophagus - just above Gastro-Esophageal Junction",
    "Oesophagus, distal": "Distal Esophagus",
    "Oesophagus, distal & GOJ": "Distal Esophagus and Gastro-Esophageal Junction",
    "Oesophagus, mid-lower": "Mid-Lower Esophagus",
    "One lymph node": "NR",
    "Pancreas (distal)": "Distal Pancreas",
    "Pancreas (head)": "Head of Pancreas",
    "Pancreas (main duct)": "Main Duct of Pancreas",
    "Pancreas (tail)": "Tail of Pancreas",
    "Pancreas (tail); infiltrating kidney (B10)": "Tail of Pancreas and inflitrating Kidney",
    "Para-oesophageal lymph nodes": "Para-esophageal Lymph node",
    "Paraoesophageal lymph node": "Para-esophageal Lymph node",
    "Perigastric lymph node": "Perigastric Lymph node",
    "Perigastric nodes": "Perigastric Lymph node",
    "Perihilar region": "NR",
    "Post auricular left": "Left Post Auricular",
    "R": "Right Colon",
    "R Atrium": "Right Atrium",
    "R Breast": "Right Breast",
    "R Tongue": "Right Tongue",
    "R extended hemiglossectomy": "Right Tongue",
    "R soft tissue and thyroid": "Right Soft Tissue and Thyroid",
    "R upper back": "Right Upper Back",
    "Rectum": "Rectum",
    "Right": "Right Colon",
    "Right (Peritoneum)": "Right Peritoneum",
    "Right Kidney": "Right Kidney",
    "Right Scalp": "Right Scalp",
    "Right anterior floor of mouth and partial glossectomy": "Right Anterior Floor of Mouth and Tongue",
    "Right anterior tongue": "Right Anterior Tongue",
    "Right back": "Right Back",
    "Right breast": "Right Breast",
    "Right chest wall": "Right Chest Wall",
    "Right ethmoidal sinus": "Right Ethmoidal Sinus",
    "Right floor of mouth tumour": "Right Floor of Mouth",
    "Right lateral oropharyngectomy": "Right Lateral Oropharynx",
    "Right oropharyngectomy": "Right Oropharynx",
    "Right pharyngeal wall (main tumour in supraglottis although tumour and high grade dysplasia also found in R pharyngeal wall and tongue base)": "Right Pharyngeal Wall",
    "Right tonsillar": "Right Tonsil",
    "Satellite tumour": "Satellite Tumor Deposit",
    "Satellite tumour deposit in mesocolon": "Satellite Tumor Deposit (Mesocolon)",
    "Satellite tumour deposit, pericolic fat": "Satellite Tumor Deposit (Pericolic Fat)",
    "Sigma": "Sigmoid Colon",
    "Sigma (Liver)": "Sigmoid Colon and Liver",
    "Sigmoid colon": "Sigmoid Colon",
    "Skin, Right facial subcutaneous": "Right Facial Subcutaneous",
    "Soft palate": "Soft Palate",
    "Splenic flexure": "Splenic Flexure",
    "Stomach; involves right colon": "Stomach and involves Right Colon",
    "Tissue around GOJ": "Around the Gastro-Esophageal Junction",
    "Tongue base": "Base of Tongue",
    "Tonsil tumour": "Tonsil",
    "Tonsil, lateral pharynx": "Tonsil and Lateral Pharynx",
    "Transverse (proximal third)": "Transverse Colon (Proximal Third)",
    "Transverse (proximal)": "Transverse Colon (Proximal)",
    "Transverse colon": "Transverse Colon",
    "Tumour": "NR",
    "Vocal cord (tumour originated in L pyriform fossa)": "Vocal Cord originating in Left Pyriform Fossa",
    "lung": "NR",
}

for df_name, df in tables_mapped.items():
    if "Organ_Combined" in df.columns:
        df["Organ"] = df["Organ_Combined"].map(organs_mapping)
        df["Organ sublocalization"] = (
            df["Organ_Combined"]
            .map(sublocalization_mapping)
            .fillna(df["Organ_Combined"].str.title())
        )
        tables_mapped[df_name] = df


# 5. Staging, Grading, and Prognostic Scores Integration
tumor_stage_columns = [
    "AJCC_Stage_Ann_Arbor",
    "AJCC_Stage_Or_Other_Applicable_System",
    "AJCC_TNM_Clinical_Or_Other_Applicable",
    "AJCC_TNM_Pathological_Or_Other_Applicable",
    "Grade",
    "Histological_Grade",
    "M",
    "N",
    "Stage",
    "Stage At Dx",
    "StageAtDX",
    "Stage_at_DX",
    "T",
    "TNM",
    "TNM_Stage",
    "WHO_Grade",
    "WHO_Grade_Primary",
    "cM",
    "pM",
    "pN",
    "pT",
    "Gleason_Score",
]

staging_mapping = {
    "": {
        "Tumor Stage": "",
        "TNM: Tumor": "",
        "TNM: Node": "",
        "TNM: Metastasis": "",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "AJCC_Stage_Ann_Arbor: IIB | AJCC_TNM_Clinical_Or_Other_Applicable: N.A. | AJCC_TNM_Pathological_Or_Other_Applicable: pT2c N0 Mx | Histological_Grade: Gleason 8 (5+3)": {
        "Tumor Stage": "II B",
        "TNM: Tumor": "pT2c",
        "TNM: Node": "N0",
        "TNM: Metastasis": "Mx",
        "Grade": "",
        "Additional tumor classification": "Gleason 8 (5+3)",
    },
    "AJCC_Stage_Ann_Arbor: IV | AJCC_TNM_Clinical_Or_Other_Applicable: N.A. | AJCC_TNM_Pathological_Or_Other_Applicable: pT3b N1 Mx | Histological_Grade: Gleason 8 (4+4)": {
        "Tumor Stage": "IV",
        "TNM: Tumor": "pT3b",
        "TNM: Node": "N1",
        "TNM: Metastasis": "Mx",
        "Grade": "",
        "Additional tumor classification": "Gleason 8 (4+4)",
    },
    "AJCC_Stage_Ann_Arbor: IV | AJCC_TNM_Clinical_Or_Other_Applicable: N.A. | AJCC_TNM_Pathological_Or_Other_Applicable: pT3b N1 Mx | Histological_Grade: Gleason 9 (5+4)": {
        "Tumor Stage": "IV",
        "TNM: Tumor": "pT3b",
        "TNM: Node": "N1",
        "TNM: Metastasis": "Mx",
        "Grade": "",
        "Additional tumor classification": "Gleason 9 (5+4)",
    },
    "AJCC_Stage_Or_Other_Applicable_System: I (at least) | AJCC_TNM_Clinical_Or_Other_Applicable: T1 Nx Mx | AJCC_TNM_Pathological_Or_Other_Applicable: N.A. | Histological_Grade: G1, Edmondson-Steiner: GII": {
        "Tumor Stage": "I (at least)",
        "TNM: Tumor": "T1",
        "TNM: Node": "Nx",
        "TNM: Metastasis": "Mx",
        "Grade": "1",
        "Additional tumor classification": "Edmondson-Steiner: GII",
    },
    "AJCC_Stage_Or_Other_Applicable_System: II (at least) | AJCC_TNM_Clinical_Or_Other_Applicable: T2 Nx Mx | AJCC_TNM_Pathological_Or_Other_Applicable: N.A. | Histological_Grade: G1, Edmondson-Steiner: GII": {
        "Tumor Stage": "II (at least)",
        "TNM: Tumor": "T2",
        "TNM: Node": "Nx",
        "TNM: Metastasis": "Mx",
        "Grade": "1",
        "Additional tumor classification": "Edmondson-Steiner: GII",
    },
    "AJCC_Stage_Or_Other_Applicable_System: II (at least) | AJCC_TNM_Clinical_Or_Other_Applicable: T2 Nx Mx | AJCC_TNM_Pathological_Or_Other_Applicable: N.A. | Histological_Grade: G1; \nEdmondson-Steiner: GII": {
        "Tumor Stage": "II (at least)",
        "TNM: Tumor": "T2",
        "TNM: Node": "Nx",
        "TNM: Metastasis": "Mx",
        "Grade": "1",
        "Additional tumor classification": "Edmondson-Steiner: GII",
    },
    "AJCC_Stage_Or_Other_Applicable_System: II (at least) | AJCC_TNM_Clinical_Or_Other_Applicable: T2 Nx Mx | AJCC_TNM_Pathological_Or_Other_Applicable: N.A. | Histological_Grade: G2, Edmondson-Steiner: GIII": {
        "Tumor Stage": "II (at least)",
        "TNM: Tumor": "T2",
        "TNM: Node": "Nx",
        "TNM: Metastasis": "Mx",
        "Grade": "2",
        "Additional tumor classification": "Edmondson-Steiner: GIII",
    },
    "AJCC_Stage_Or_Other_Applicable_System: II (at least) | AJCC_TNM_Clinical_Or_Other_Applicable: T2 Nx Mx | AJCC_TNM_Pathological_Or_Other_Applicable: N.A. | Histological_Grade: G2;\nNuclear grade 2": {
        "Tumor Stage": "II (at least)",
        "TNM: Tumor": "T2",
        "TNM: Node": "Nx",
        "TNM: Metastasis": "Mx",
        "Grade": "2",
        "Additional tumor classification": "Nuclear grade 2",
    },
    "AJCC_Stage_Or_Other_Applicable_System: IIB | AJCC_TNM_Clinical_Or_Other_Applicable: T3 N1 M0 | AJCC_TNM_Pathological_Or_Other_Applicable: N.A. | Histological_Grade: G1": {
        "Tumor Stage": "II B",
        "TNM: Tumor": "T3",
        "TNM: Node": "N1",
        "TNM: Metastasis": "M0",
        "Grade": "1",
        "Additional tumor classification": "",
    },
    "AJCC_Stage_Or_Other_Applicable_System: IIB | AJCC_TNM_Clinical_Or_Other_Applicable: T3 N1 M0 | AJCC_TNM_Pathological_Or_Other_Applicable: N.A. | Histological_Grade: G2": {
        "Tumor Stage": "II B",
        "TNM: Tumor": "T3",
        "TNM: Node": "N1",
        "TNM: Metastasis": "M0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "AJCC_Stage_Or_Other_Applicable_System: IIB | AJCC_TNM_Clinical_Or_Other_Applicable: T3 N1 M0 | AJCC_TNM_Pathological_Or_Other_Applicable: N.A. | Histological_Grade: G3": {
        "Tumor Stage": "II B",
        "TNM: Tumor": "T3",
        "TNM: Node": "N1",
        "TNM: Metastasis": "M0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "AJCC_Stage_Or_Other_Applicable_System: III | AJCC_TNM_Clinical_Or_Other_Applicable: T2 N2 Mx | AJCC_TNM_Pathological_Or_Other_Applicable: N.A. | Histological_Grade: G2": {
        "Tumor Stage": "III",
        "TNM: Tumor": "T2",
        "TNM: Node": "N2",
        "TNM: Metastasis": "Mx",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "AJCC_Stage_Or_Other_Applicable_System: IIIA (at least) | AJCC_TNM_Clinical_Or_Other_Applicable: T3a Nx Mx | AJCC_TNM_Pathological_Or_Other_Applicable: N.A. | Histological_Grade: G1;\nEdmondson-Steiner: GII": {
        "Tumor Stage": "III A (at least)",
        "TNM: Tumor": "T3a",
        "TNM: Node": "Nx",
        "TNM: Metastasis": "Mx",
        "Grade": "1",
        "Additional tumor classification": "Edmondson-Steiner: GII",
    },
    "AJCC_Stage_Or_Other_Applicable_System: ISUP Grade 3 | AJCC_TNM_Clinical_Or_Other_Applicable: N.A. | AJCC_TNM_Pathological_Or_Other_Applicable: pT3b N1 | Histological_Grade: Gleason 7 (4+3)": {
        "Tumor Stage": "",
        "TNM: Tumor": "pT3b",
        "TNM: Node": "N1",
        "TNM: Metastasis": "",
        "Grade": "",
        "Additional tumor classification": "ISUP Grade 3, Gleason 7 (4+3)",
    },
    "AJCC_Stage_Or_Other_Applicable_System: ISUP Grade 5 | AJCC_TNM_Clinical_Or_Other_Applicable: N.A. | AJCC_TNM_Pathological_Or_Other_Applicable: pT3b N1 | Histological_Grade: Gleason 9 (4+5)": {
        "Tumor Stage": "",
        "TNM: Tumor": "pT3b",
        "TNM: Node": "N1",
        "TNM: Metastasis": "",
        "Grade": "",
        "Additional tumor classification": "ISUP Grade 5, Gleason 9 (4+5)",
    },
    "AJCC_Stage_Or_Other_Applicable_System: IV | AJCC_TNM_Clinical_Or_Other_Applicable: N.A. | AJCC_TNM_Pathological_Or_Other_Applicable: pT3a N1 Mx | Histological_Grade: Gleason 9 (4+5)": {
        "Tumor Stage": "IV",
        "TNM: Tumor": "pT3a",
        "TNM: Node": "N1",
        "TNM: Metastasis": "Mx",
        "Grade": "",
        "Additional tumor classification": "Gleason 9 (4+5)",
    },
    "AJCC_Stage_Or_Other_Applicable_System: IV | AJCC_TNM_Clinical_Or_Other_Applicable: N.A. | AJCC_TNM_Pathological_Or_Other_Applicable: pT3b N1 Mx | Histological_Grade: Gleason 7 (4+3)": {
        "Tumor Stage": "IV",
        "TNM: Tumor": "pT3b",
        "TNM: Node": "N1",
        "TNM: Metastasis": "Mx",
        "Grade": "",
        "Additional tumor classification": "Gleason 7 (4+3)",
    },
    "AJCC_Stage_Or_Other_Applicable_System: N.A. | AJCC_TNM_Clinical_Or_Other_Applicable: N.A. | AJCC_TNM_Pathological_Or_Other_Applicable: N.A. | Histological_Grade: IV": {
        "Tumor Stage": "",
        "TNM: Tumor": "",
        "TNM: Node": "",
        "TNM: Metastasis": "",
        "Grade": "4",
        "Additional tumor classification": "",
    },
    "Gleason_Score: 4 + 5 (9)": {
        "Tumor Stage": "",
        "TNM: Tumor": "",
        "TNM: Node": "",
        "TNM: Metastasis": "",
        "Grade": "",
        "Additional tumor classification": "Gleason 4+5=9",
    },
    "Gleason_Score: 5 + 4 (9)": {
        "Tumor Stage": "",
        "TNM: Tumor": "",
        "TNM: Node": "",
        "TNM: Metastasis": "",
        "Grade": "",
        "Additional tumor classification": "Gleason 5+4=9",
    },
    "Grade: - | M: - | N: - | StageAtDX: - | T: -": {
        "Tumor Stage": "",
        "TNM: Tumor": "",
        "TNM: Node": "",
        "TNM: Metastasis": "",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "Grade: - | M: x | N: 1 | StageAtDX: - | T: 1": {
        "Tumor Stage": "",
        "TNM: Tumor": "T1",
        "TNM: Node": "N1",
        "TNM: Metastasis": "Mx",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "Grade: / | M: / | N: / | Stage: IIIC | T: /": {
        "Tumor Stage": "III C",
        "TNM: Tumor": "",
        "TNM: Node": "",
        "TNM: Metastasis": "",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "Grade: / | M: 0 | N: 2 | T: 4a": {
        "Tumor Stage": "",
        "TNM: Tumor": "T4a",
        "TNM: Node": "N2",
        "TNM: Metastasis": "M0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "Grade: 1 | M: 0 | N: 0 | Stage: IA | T: 1c": {
        "Tumor Stage": "I A",
        "TNM: Tumor": "T1c",
        "TNM: Node": "N0",
        "TNM: Metastasis": "M0",
        "Grade": "1",
        "Additional tumor classification": "",
    },
    "Grade: 1 | M: 0 | N: 0 | Stage: IA2 | T: 1b": {
        "Tumor Stage": "I A2",
        "TNM: Tumor": "T1b",
        "TNM: Node": "N0",
        "TNM: Metastasis": "M0",
        "Grade": "1",
        "Additional tumor classification": "",
    },
    "Grade: 1 | M: x | N: 0 | StageAtDX: IIA | T: 3": {
        "Tumor Stage": "II A",
        "TNM: Tumor": "T3",
        "TNM: Node": "N0",
        "TNM: Metastasis": "Mx",
        "Grade": "1",
        "Additional tumor classification": "",
    },
    "Grade: 1 | StageAtDX: I | TNM_Stage: T1N0M0 | cM: 0 | pT: 1": {
        "Tumor Stage": "I",
        "TNM: Tumor": "pT1",
        "TNM: Node": "N0",
        "TNM: Metastasis": "cM0",
        "Grade": "1",
        "Additional tumor classification": "",
    },
    "Grade: 1 | StageAtDX: IA | TNM_Stage: T1N0M0 | cM: 0 | pT: 1": {
        "Tumor Stage": "I A",
        "TNM: Tumor": "pT1",
        "TNM: Node": "N0",
        "TNM: Metastasis": "cM0",
        "Grade": "1",
        "Additional tumor classification": "",
    },
    "Grade: 1 | StageAtDX: IB | TNM_Stage: T2N0M0 | cM: 0 | pT: 2": {
        "Tumor Stage": "I B",
        "TNM: Tumor": "pT2",
        "TNM: Node": "N0",
        "TNM: Metastasis": "cM0",
        "Grade": "1",
        "Additional tumor classification": "",
    },
    "Grade: 1.0": {
        "Tumor Stage": "",
        "TNM: Tumor": "",
        "TNM: Node": "",
        "TNM: Metastasis": "",
        "Grade": "1",
        "Additional tumor classification": "",
    },
    "Grade: 2 | M: - | N: - | StageAtDX: - | T: -": {
        "Tumor Stage": "",
        "TNM: Tumor": "",
        "TNM: Node": "",
        "TNM: Metastasis": "",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2 | M: / | N: / | Stage: / | T: /": {
        "Tumor Stage": "",
        "TNM: Tumor": "",
        "TNM: Node": "",
        "TNM: Metastasis": "",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2 | M: 0 | N: 0 | Stage: / | T: 1": {
        "Tumor Stage": "",
        "TNM: Tumor": "T1",
        "TNM: Node": "N0",
        "TNM: Metastasis": "M0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2 | M: 0 | N: 0 | Stage: / | T: 2": {
        "Tumor Stage": "",
        "TNM: Tumor": "T2",
        "TNM: Node": "N0",
        "TNM: Metastasis": "M0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2 | M: 0 | N: 0 | Stage: / | T: 3": {
        "Tumor Stage": "",
        "TNM: Tumor": "T3",
        "TNM: Node": "N0",
        "TNM: Metastasis": "M0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2 | M: 0 | N: 0 | Stage: IIA | T: 2": {
        "Tumor Stage": "II A",
        "TNM: Tumor": "T2",
        "TNM: Node": "N0",
        "TNM: Metastasis": "M0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2 | M: 0 | N: 1 | Stage: / | T: 2": {
        "Tumor Stage": "",
        "TNM: Tumor": "T2",
        "TNM: Node": "N1",
        "TNM: Metastasis": "M0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2 | M: 0 | N: 1 | Stage: / | T: 3": {
        "Tumor Stage": "",
        "TNM: Tumor": "T3",
        "TNM: Node": "N1",
        "TNM: Metastasis": "M0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2 | M: 0 | N: 1 | Stage: 0 | T: 2b": {
        "Tumor Stage": "0",
        "TNM: Tumor": "T2b",
        "TNM: Node": "N1",
        "TNM: Metastasis": "M0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2 | M: 0 | N: 1 | Stage: IIB | T: 2": {
        "Tumor Stage": "II B",
        "TNM: Tumor": "T2",
        "TNM: Node": "N1",
        "TNM: Metastasis": "M0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2 | M: 0 | N: 1 | Stage: IIIB | T: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "T3",
        "TNM: Node": "N1",
        "TNM: Metastasis": "M0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2 | M: 0 | N: 1 | StageAtDX: IIA | T: 3": {
        "Tumor Stage": "II A",
        "TNM: Tumor": "T3",
        "TNM: Node": "N1",
        "TNM: Metastasis": "M0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2 | M: 0 | N: 2 | Stage: IIIB | T: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "T3",
        "TNM: Node": "N2",
        "TNM: Metastasis": "M0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2 | M: 0 | N: 2a | Stage: / | T: 3": {
        "Tumor Stage": "",
        "TNM: Tumor": "T3",
        "TNM: Node": "N2a",
        "TNM: Metastasis": "M0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2 | M: 0 | N: 2a | Stage: IIIB | T: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "T3",
        "TNM: Node": "N2a",
        "TNM: Metastasis": "M0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2 | M: 0 | N: 2b | Stage: / | T: 4a": {
        "Tumor Stage": "",
        "TNM: Tumor": "T4a",
        "TNM: Node": "N2b",
        "TNM: Metastasis": "M0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2 | M: 0 | N: 3a | Stage: / | T: 2": {
        "Tumor Stage": "",
        "TNM: Tumor": "T2",
        "TNM: Node": "N3a",
        "TNM: Metastasis": "M0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2 | M: 0 | N: 3a | Stage: / | T: 4a": {
        "Tumor Stage": "",
        "TNM: Tumor": "T4a",
        "TNM: Node": "N3a",
        "TNM: Metastasis": "M0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2 | M: 0 | N: 3a | T: 4a": {
        "Tumor Stage": "",
        "TNM: Tumor": "T4a",
        "TNM: Node": "N3a",
        "TNM: Metastasis": "M0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2 | M: 1 | N: 0 | Stage: / | T: 3": {
        "Tumor Stage": "",
        "TNM: Tumor": "T3",
        "TNM: Node": "N0",
        "TNM: Metastasis": "M1",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2 | M: 1 | N: 1 | Stage: / | T: 3": {
        "Tumor Stage": "",
        "TNM: Tumor": "T3",
        "TNM: Node": "N1",
        "TNM: Metastasis": "M1",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2 | M: 1 | N: 1 | Stage: IV | T: 3": {
        "Tumor Stage": "IV",
        "TNM: Tumor": "T3",
        "TNM: Node": "N1",
        "TNM: Metastasis": "M1",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2 | M: 1 | N: 2 | Stage: / | T: 3": {
        "Tumor Stage": "",
        "TNM: Tumor": "T3",
        "TNM: Node": "N2",
        "TNM: Metastasis": "M1",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2 | M: 1 | N: 2 | Stage: IV | T: 3": {
        "Tumor Stage": "IV",
        "TNM: Tumor": "T3",
        "TNM: Node": "N2",
        "TNM: Metastasis": "M1",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2 | M: 1 | N: 2a | Stage: / | T: 3": {
        "Tumor Stage": "",
        "TNM: Tumor": "T3",
        "TNM: Node": "N2a",
        "TNM: Metastasis": "M1",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2 | M: 1 | N: 2b | Stage: IV | T: 3": {
        "Tumor Stage": "IV",
        "TNM: Tumor": "T3",
        "TNM: Node": "N2b",
        "TNM: Metastasis": "M1",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2 | M: 1 | N: 2b | Stage: IV | T: 4a": {
        "Tumor Stage": "IV",
        "TNM: Tumor": "T4a",
        "TNM: Node": "N2b",
        "TNM: Metastasis": "M1",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2 | M: N.A. | N: 3a | Stage: 0 | T: 3": {
        "Tumor Stage": "0",
        "TNM: Tumor": "T3",
        "TNM: Node": "N3a",
        "TNM: Metastasis": "",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2 | M: x | N: 0 | StageAtDX: IIA | T: 3": {
        "Tumor Stage": "II A",
        "TNM: Tumor": "T3",
        "TNM: Node": "N0",
        "TNM: Metastasis": "Mx",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2 | StageAtDX: II | TNM_Stage: T1N0M0 | cM: 0 | pT: 1": {
        "Tumor Stage": "II",
        "TNM: Tumor": "pT1",
        "TNM: Node": "N0",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2 | StageAtDX: IIIA | TNM_Stage: T2N0M0 | cM: 0 | pT: 2": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "pT2",
        "TNM: Node": "N0",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2 | TNM_Stage: TXN0M0 | cM: 0": {
        "Tumor Stage": "",
        "TNM: Tumor": "Tx",
        "TNM: Node": "N0",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2.0": {
        "Tumor Stage": "",
        "TNM: Tumor": "",
        "TNM: Node": "",
        "TNM: Metastasis": "",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2.0 | Stage: IIB | TNM: T3N1M0": {
        "Tumor Stage": "II B",
        "TNM: Tumor": "T3",
        "TNM: Node": "N1",
        "TNM: Metastasis": "M0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2.0 | Stage: IIIA | TNM: T3N2M0": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "T3",
        "TNM: Node": "N2",
        "TNM: Metastasis": "M0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2.0 | Stage: IIIA | TNM: T4N1M0": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "T4",
        "TNM: Node": "N1",
        "TNM: Metastasis": "M0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2.0 | Stage: IIIB | TNM: T4N2M0": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "T4",
        "TNM: Node": "N2",
        "TNM: Metastasis": "M0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2.0 | Stage: IV | TNM: M1": {
        "Tumor Stage": "IV",
        "TNM: Tumor": "",
        "TNM: Node": "",
        "TNM: Metastasis": "M1",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2.0 | Stage: IV | TNM: T3N2M1": {
        "Tumor Stage": "IV",
        "TNM: Tumor": "T3",
        "TNM: Node": "N2",
        "TNM: Metastasis": "M1",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2.0 | TNM: T3N2": {
        "Tumor Stage": "",
        "TNM: Tumor": "T3",
        "TNM: Node": "N2",
        "TNM: Metastasis": "",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: 2026-02-01 00:00:00 | M: - | N: - | StageAtDX: - | T: -": {
        "Tumor Stage": "",
        "TNM: Tumor": "",
        "TNM: Node": "",
        "TNM: Metastasis": "",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "Grade: 2026-02-01 00:00:00 | StageAtDX: IIIC | TNM_Stage: T4AN3M0 | cM: 0 | pN: 3 | pT: 4a": {
        "Tumor Stage": "III C",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN3",
        "TNM: Metastasis": "cM0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "Grade: 3 | M: - | N: - | StageAtDX: - | T: -": {
        "Tumor Stage": "",
        "TNM: Tumor": "",
        "TNM: Node": "",
        "TNM: Metastasis": "",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | M: / | N: / | Stage: / | T: /": {
        "Tumor Stage": "",
        "TNM: Tumor": "",
        "TNM: Node": "",
        "TNM: Metastasis": "",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | M: / | N: / | Stage: / | T: 2": {
        "Tumor Stage": "",
        "TNM: Tumor": "T2",
        "TNM: Node": "",
        "TNM: Metastasis": "",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | M: 0 | N: 0 | Stage: IIA | T: 2": {
        "Tumor Stage": "II A",
        "TNM: Tumor": "T2",
        "TNM: Node": "N0",
        "TNM: Metastasis": "M0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | M: 0 | N: 0 | Stage: IIB | T: 3": {
        "Tumor Stage": "II B",
        "TNM: Tumor": "T3",
        "TNM: Node": "N0",
        "TNM: Metastasis": "M0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | M: 0 | N: 0 | StageAtDX: IIA | T: 2": {
        "Tumor Stage": "II A",
        "TNM: Tumor": "T2",
        "TNM: Node": "N0",
        "TNM: Metastasis": "M0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | M: 0 | N: 0(sn) | Stage: IA | T: 1c": {
        "Tumor Stage": "I A",
        "TNM: Tumor": "T1c",
        "TNM: Node": "N0(sn)",
        "TNM: Metastasis": "M0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | M: 0 | N: 1 | Stage: IIB | T: 2": {
        "Tumor Stage": "II B",
        "TNM: Tumor": "T2",
        "TNM: Node": "N1",
        "TNM: Metastasis": "M0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | M: 0 | N: 1 | Stage: IIB | T: 2(m)": {
        "Tumor Stage": "II B",
        "TNM: Tumor": "T2(m)",
        "TNM: Node": "N1",
        "TNM: Metastasis": "M0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | M: 0 | N: 1 | Stage: IIIB | T: 4": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "T4",
        "TNM: Node": "N1",
        "TNM: Metastasis": "M0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | M: 0 | N: 1b | Stage: IIIB | T: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "T3",
        "TNM: Node": "N1b",
        "TNM: Metastasis": "M0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | M: 0 | N: 2 | Stage: / | T: 4a": {
        "Tumor Stage": "",
        "TNM: Tumor": "T4a",
        "TNM: Node": "N2",
        "TNM: Metastasis": "M0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | M: 0 | N: 2 | Stage: IIIA | T: 2": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "T2",
        "TNM: Node": "N2",
        "TNM: Metastasis": "M0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | M: 0 | N: 2 | Stage: IIIA | T: 2a": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "T2a",
        "TNM: Node": "N2",
        "TNM: Metastasis": "M0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | M: 0 | N: 2 | Stage: IIIA | T: 3": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "T3",
        "TNM: Node": "N2",
        "TNM: Metastasis": "M0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | M: 0 | N: 2 | Stage: IIIB | T: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "T3",
        "TNM: Node": "N2",
        "TNM: Metastasis": "M0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | M: 0 | N: 2 | Stage: IIIB | T: 4": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "T4",
        "TNM: Node": "N2",
        "TNM: Metastasis": "M0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | M: 0 | N: 3 | Stage: / | T: 2b": {
        "Tumor Stage": "",
        "TNM: Tumor": "T2b",
        "TNM: Node": "N3",
        "TNM: Metastasis": "M0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | M: 0 | N: 3 | Stage: IIIC | T: 1c": {
        "Tumor Stage": "III C",
        "TNM: Tumor": "T1c",
        "TNM: Node": "N3",
        "TNM: Metastasis": "M0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | M: 0 | N: 3 | Stage: IIIC | T: 2": {
        "Tumor Stage": "III C",
        "TNM: Tumor": "T2",
        "TNM: Node": "N3",
        "TNM: Metastasis": "M0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | M: 0 | N: 3 | Stage: IIIC | T: 4": {
        "Tumor Stage": "III C",
        "TNM: Tumor": "T4",
        "TNM: Node": "N3",
        "TNM: Metastasis": "M0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | M: 0 | N: 3a | Stage: / | T: 4a": {
        "Tumor Stage": "",
        "TNM: Tumor": "T4a",
        "TNM: Node": "N3a",
        "TNM: Metastasis": "M0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | M: 1 | N: 2 | Stage: / | T: 3": {
        "Tumor Stage": "",
        "TNM: Tumor": "T3",
        "TNM: Node": "N2",
        "TNM: Metastasis": "M1",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | M: 1 | N: 3a | Stage: / | T: 4a": {
        "Tumor Stage": "",
        "TNM: Tumor": "T4a",
        "TNM: Node": "N3a",
        "TNM: Metastasis": "M1",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | M: 1 | N: 3b | Stage: / | T: 4a": {
        "Tumor Stage": "",
        "TNM: Tumor": "T4a",
        "TNM: Node": "N3b",
        "TNM: Metastasis": "M1",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | M: x | N: 0 | StageAtDX: IIA | T: 3": {
        "Tumor Stage": "II A",
        "TNM: Tumor": "T3",
        "TNM: Node": "N0",
        "TNM: Metastasis": "Mx",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | M: x | N: 3 | StageAtDX: IIIB | T: 2": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "T2",
        "TNM: Node": "N3",
        "TNM: Metastasis": "Mx",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | M: x | N: x | Stage: / | T: 2b": {
        "Tumor Stage": "",
        "TNM: Tumor": "T2b",
        "TNM: Node": "Nx",
        "TNM: Metastasis": "Mx",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | StageAtDX: II | TNM_Stage: T1N0M0 | cM: 0 | pM: x | pN: x | pT: 1": {
        "Tumor Stage": "II",
        "TNM: Tumor": "pT1",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | StageAtDX: II | TNM_Stage: T1N0M0 | cM: 0 | pT: 1": {
        "Tumor Stage": "II",
        "TNM: Tumor": "pT1",
        "TNM: Node": "N0",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | StageAtDX: III | TNM_Stage: T3CN1M0 | cM: 0 | pM: 0 | pN: 1 | pT: 3c": {
        "Tumor Stage": "III",
        "TNM: Tumor": "pT3c",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | StageAtDX: IIIA | TNM_Stage: T2N0M0 | cM: 0 | pM: x | pN: x | pT: 2": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | StageAtDX: IIIA | TNM_Stage: T2N0M0 | cM: 0 | pT: 2": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "pT2",
        "TNM: Node": "N0",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | StageAtDX: IIIA | TNM_Stage: T3N1M0 | cM: 0 | pN: 1 | pT: 3": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | StageAtDX: IIIB | TNM_Stage: T3N0M0 | cM: 0 | pT: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "N0",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | StageAtDX: IIIB | TNM_Stage: T4AN1M0 | cM: 0 | pM: 0 | pN: 1 | pT: 4a": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | StageAtDX: IIIC | TNM_Stage: T3N3M0 | cM: 0 | pM: x | pN: 3 | pT: 3": {
        "Tumor Stage": "III C",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN3",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | StageAtDX: IV | TNM_Stage: T2N0M1 | cM: 1 (Lung) | pM: x | pT: 2": {
        "Tumor Stage": "IV",
        "TNM: Tumor": "pT2",
        "TNM: Node": "N0",
        "TNM: Metastasis": "cM1 (Lung), pMx",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | StageAtDX: IV | TNM_Stage: T2N0M1 | cM: 1 (lung) | pT: 2": {
        "Tumor Stage": "IV",
        "TNM: Tumor": "pT2",
        "TNM: Node": "N0",
        "TNM: Metastasis": "cM1 (lung)",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | StageAtDX: IV | TNM_Stage: T3AN0M1 | cM: 1 (Lung, cervical lymph node) | pM: 0 | pN: x | pT: 3a": {
        "Tumor Stage": "IV",
        "TNM: Tumor": "pT3a",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM1 (Lung, cervical lymph node), pM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | StageAtDX: IV | TNM_Stage: T4N0M0 | cM: 0 | pM: 0 | pN: 0 | pT: 4": {
        "Tumor Stage": "IV",
        "TNM: Tumor": "pT4",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | StageAtDX: IVA | TNM_Stage: T3N3M0 | cM: 0 | pN: 3 | pT: 3": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN3",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3 | StageAtDX: IVA | TNM_Stage: T4N2M0 | cM: 0 | pN: 2 | pT: 4": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT4",
        "TNM: Node": "pN2",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3.0": {
        "Tumor Stage": "",
        "TNM: Tumor": "",
        "TNM: Node": "",
        "TNM: Metastasis": "",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3.0 | Stage: IIIA | TNM: T3N2M0": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "T3",
        "TNM: Node": "N2",
        "TNM: Metastasis": "M0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3.0 | Stage: IIIA | TNM: T3N2Mx": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "T3",
        "TNM: Node": "N2",
        "TNM: Metastasis": "Mx",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3.0 | Stage: IIIB | TNM: T4N2M0": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "T4",
        "TNM: Node": "N2",
        "TNM: Metastasis": "M0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 3.0 | TNM: T3N1M0": {
        "Tumor Stage": "",
        "TNM: Tumor": "T3",
        "TNM: Node": "N1",
        "TNM: Metastasis": "M0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: 4 | M: 1 | N: 2 | Stage: IV | T: 3": {
        "Tumor Stage": "IV",
        "TNM: Tumor": "T3",
        "TNM: Node": "N2",
        "TNM: Metastasis": "M1",
        "Grade": "4",
        "Additional tumor classification": "",
    },
    "Grade: 4 | StageAtDX: III | TNM_Stage: T3A(m)N0M0 | cM: 0 | pM: 1 (Adrenal) | pN: 0 | pT: 3a(m)": {
        "Tumor Stage": "III",
        "TNM: Tumor": "pT3a(m)",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0, pM1 (Adrenal)",
        "Grade": "4",
        "Additional tumor classification": "",
    },
    "Grade: 4 | StageAtDX: III | TNM_Stage: T3AN0M0 | cM: 0 | pM: 0 | pN: 0 | pT: 3a": {
        "Tumor Stage": "III",
        "TNM: Tumor": "pT3a",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "4",
        "Additional tumor classification": "",
    },
    "Grade: 4 | StageAtDX: IV | TNM_Stage: T2AN0M1 | cM: 1 (Lung) | pM: 0 | pN: 0 | pT: 2a": {
        "Tumor Stage": "IV",
        "TNM: Tumor": "pT2a",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM1 (Lung), pM0",
        "Grade": "4",
        "Additional tumor classification": "",
    },
    "Grade: 4 | StageAtDX: IV | TNM_Stage: T2BN0M1 | cM: 1 (Pancreas, Lung) | pM: 1 | pN: 0 | pT: 2b": {
        "Tumor Stage": "IV",
        "TNM: Tumor": "pT2b",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM1 (Pancreas, Lung), pM1",
        "Grade": "4",
        "Additional tumor classification": "",
    },
    "Grade: 4 | StageAtDX: IV | TNM_Stage: T3AN0M1 | cM: 1 (Lung) | pM: 0 | pN: x | pT: 3a": {
        "Tumor Stage": "IV",
        "TNM: Tumor": "pT3a",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM1 (Lung), pM0",
        "Grade": "4",
        "Additional tumor classification": "",
    },
    "Grade: 4 | StageAtDX: IV | TNM_Stage: T3AN1M1 | cM: 1 (Lung) | pM: 0 | pN: 0 | pT: 3a": {
        "Tumor Stage": "IV",
        "TNM: Tumor": "pT3a",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM1 (Lung), pM0",
        "Grade": "4",
        "Additional tumor classification": "",
    },
    "Grade: 4 | StageAtDX: IV | TNM_Stage: T4N0M1 | cM: 1 (Liver) | pM: 1 | pN: 0 | pT: 4": {
        "Tumor Stage": "IV",
        "TNM: Tumor": "pT4",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM1 (Liver), pM1",
        "Grade": "4",
        "Additional tumor classification": "",
    },
    "Grade: 5 | StageAtDX: IV | TNM_Stage: T3AN1M0 | cM: 0 | pM: X | pN: 1 | pT: 3a": {
        "Tumor Stage": "IV",
        "TNM: Tumor": "pT3a",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "",
        "Additional tumor classification": "ISUP Grade 5",
    },
    "Grade: 5 | StageAtDX: IV | TNM_Stage: T3BN1M0 | cM: 0 | pM: X | pN: 1 | pT: 3b": {
        "Tumor Stage": "IV",
        "TNM: Tumor": "pT3b",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "",
        "Additional tumor classification": "ISUP Grade 5",
    },
    "Grade: 5 | StageAtDX: IVA | TNM_Stage: T3AN1M0 | cM: 0 | pM: X | pN: 1 | pT: 3a": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT3a",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "",
        "Additional tumor classification": "ISUP Grade 5",
    },
    "Grade: 5 | StageAtDX: IVA | TNM_Stage: T3BN1M0 | cM: 0 | pM: X | pN: 1 | pT: 3b": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT3b",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "",
        "Additional tumor classification": "ISUP Grade 5",
    },
    "Grade: G2 | Stage_at_DX: 1a": {
        "Tumor Stage": "1 a",
        "TNM: Tumor": "",
        "TNM: Node": "",
        "TNM: Metastasis": "",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: Gleason 4+5=9 | M: 1 | N: 1 | T: 3a": {
        "Tumor Stage": "",
        "TNM: Tumor": "T3a",
        "TNM: Node": "N1",
        "TNM: Metastasis": "M1",
        "Grade": "",
        "Additional tumor classification": "Gleason 4+5=9",
    },
    "Grade: Gleason Grade 3:4 (7) | M: MX | N: N3 | StageAtDX: Not stated | T: T4": {
        "Tumor Stage": "",
        "TNM: Tumor": "T4",
        "TNM: Node": "N3",
        "TNM: Metastasis": "Mx",
        "Grade": "",
        "Additional tumor classification": "Gleason 3+4=7",
    },
    "Grade: Gleason Grade 4:4 (8) | M: Not stated | N: Not stated | StageAtDX: Not stated | T: Not stated": {
        "Tumor Stage": "",
        "TNM: Tumor": "",
        "TNM: Node": "",
        "TNM: Metastasis": "",
        "Grade": "",
        "Additional tumor classification": "Gleason 4+4=8",
    },
    "Grade: Gleason Grade 5:3 (8) | M: Not stated | N: Not stated | StageAtDX: Not stated | T: Not stated": {
        "Tumor Stage": "",
        "TNM: Tumor": "",
        "TNM: Node": "",
        "TNM: Metastasis": "",
        "Grade": "",
        "Additional tumor classification": "Gleason 5+3=8",
    },
    "Grade: Gleason Grade 5:5 (10) | M: Not stated | N: Not stated | StageAtDX: Not stated | T: Not stated": {
        "Tumor Stage": "",
        "TNM: Tumor": "",
        "TNM: Node": "",
        "TNM: Metastasis": "",
        "Grade": "",
        "Additional tumor classification": "Gleason 5+5=10",
    },
    "Grade: Gleason grade: 4+4=8 | M: Mx | N: Nx | StageAtDX: Not stated | T: Tx": {
        "Tumor Stage": "",
        "TNM: Tumor": "Tx",
        "TNM: Node": "Nx",
        "TNM: Metastasis": "Mx",
        "Grade": "",
        "Additional tumor classification": "Gleason 4+4=8",
    },
    "Grade: Gleason grade: 4+5 =9 | M: Mx | N: Nx | StageAtDX: Not stated | T: Tx": {
        "Tumor Stage": "",
        "TNM: Tumor": "Tx",
        "TNM: Node": "Nx",
        "TNM: Metastasis": "Mx",
        "Grade": "",
        "Additional tumor classification": "Gleason 4+5=9",
    },
    "Grade: Gleason grade: 5+4=9 | M: Mx | N: Nx | StageAtDX: Not stated | T: Tx": {
        "Tumor Stage": "",
        "TNM: Tumor": "Tx",
        "TNM: Node": "Nx",
        "TNM: Metastasis": "Mx",
        "Grade": "",
        "Additional tumor classification": "Gleason 5+4=9",
    },
    "Grade: Gleason grade: 5+5=10 | M: Mx | N: Nx | StageAtDX: Not stated | T: Tx": {
        "Tumor Stage": "",
        "TNM: Tumor": "Tx",
        "TNM: Node": "Nx",
        "TNM: Metastasis": "Mx",
        "Grade": "",
        "Additional tumor classification": "Gleason 5+5=10",
    },
    "Grade: High grade | StageAtDX: II | TNM_Stage: T1N0M0 | cM: 0 | pT: 1": {
        "Tumor Stage": "II",
        "TNM: Tumor": "pT1",
        "TNM: Node": "N0",
        "TNM: Metastasis": "cM0",
        "Grade": "",
        "Additional tumor classification": "High grade",
    },
    "Grade: High grade | StageAtDX: IIIA | TNM_Stage: T2N0M0 | cM: 0 | pT: 2": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "pT2",
        "TNM: Node": "N0",
        "TNM: Metastasis": "cM0",
        "Grade": "",
        "Additional tumor classification": "High grade",
    },
    "Grade: High grade | StageAtDX: IIIB | TNM_Stage: T4N0M0 | cM: 0 | pN: 0 | pT: 4": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT4",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0",
        "Grade": "",
        "Additional tumor classification": "High grade",
    },
    "Grade: High | cM: 0": {
        "Tumor Stage": "",
        "TNM: Tumor": "",
        "TNM: Node": "",
        "TNM: Metastasis": "cM0",
        "Grade": "",
        "Additional tumor classification": "High grade",
    },
    "Grade: IV": {
        "Tumor Stage": "",
        "TNM: Tumor": "",
        "TNM: Node": "",
        "TNM: Metastasis": "",
        "Grade": "4",
        "Additional tumor classification": "",
    },
    "Grade: Low | cM: 0": {
        "Tumor Stage": "",
        "TNM: Tumor": "",
        "TNM: Node": "",
        "TNM: Metastasis": "cM0",
        "Grade": "",
        "Additional tumor classification": "Low grade",
    },
    "Grade: M | StageAtDX: I | TNM_Stage: T1N0M0 | cM: 0 | pM: 0 | pN: 0 | pT: 1": {
        "Tumor Stage": "I",
        "TNM: Tumor": "pT1",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IA | TNM_Stage: T1aN0M0 | cM: 0 | pM: 0 | pN: 0 | pT: 1a": {
        "Tumor Stage": "I A",
        "TNM: Tumor": "pT1a",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IA | TNM_Stage: T1aNxM0 | cM: 0 | pM: 0 | pN: x | pT: 1a": {
        "Tumor Stage": "I A",
        "TNM: Tumor": "pT1a",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IA | TNM_Stage: T1aNxM0 | cM: 0 | pN: x | pT: 1a": {
        "Tumor Stage": "I A",
        "TNM: Tumor": "pT1a",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IB | TNM_Stage: T1BN0M0 | cM: 0 | pM: 0 | pN: 0 | pT: 1b": {
        "Tumor Stage": "I B",
        "TNM: Tumor": "pT1b",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IB | TNM_Stage: T1BN1M0 | cM: 0 | pN: 1 | pT: 1b": {
        "Tumor Stage": "I B",
        "TNM: Tumor": "pT1b",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IB | TNM_Stage: T1NxM0 | cM: 0 | pN: x | pT: 1": {
        "Tumor Stage": "I B",
        "TNM: Tumor": "pT1",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IB | TNM_Stage: T1aNxM0 | cM: 0 | pN: x | pT: 1a": {
        "Tumor Stage": "I B",
        "TNM: Tumor": "pT1a",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IB | TNM_Stage: T1bN0M0 | cM: 0 | pM: 0 | pN: 0 | pT: 1b": {
        "Tumor Stage": "I B",
        "TNM: Tumor": "pT1b",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IB | TNM_Stage: T1bNxM0 | cM: 0 | pM: 0 | pN: x | pT: 1b": {
        "Tumor Stage": "I B",
        "TNM: Tumor": "pT1b",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IB | TNM_Stage: T1bNxM0 | cM: 0 | pN: x | pT: 1b": {
        "Tumor Stage": "I B",
        "TNM: Tumor": "pT1b",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: II | TNM_Stage: T1bNxM0 | cM: 0 | pM: 0 | pN: x | pT: 1b": {
        "Tumor Stage": "II",
        "TNM: Tumor": "pT1b",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: II | TNM_Stage: T2AN0M0 | cM: 0 | pM: 0 | pN: 1 | pT: 2a": {
        "Tumor Stage": "II",
        "TNM: Tumor": "pT2a",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: II | TNM_Stage: T2N0M0 | cM: 0 | pM: 0 | pN: 0 | pT: 2": {
        "Tumor Stage": "II",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: II | TNM_Stage: T2N0M0 | cM: 0 | pN: 0 | pT: 2": {
        "Tumor Stage": "II",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: II | TNM_Stage: T2NxM0 | cM: 0 | pM: 0 | pN: x | pT: 2": {
        "Tumor Stage": "II",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: II | TNM_Stage: T2NxM0 | cM: 0 | pM: x | pN: x | pT: 2": {
        "Tumor Stage": "II",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: II | TNM_Stage: T2NxM0 | cM: 0 | pN: x | pT: 2": {
        "Tumor Stage": "II",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: II | TNM_Stage: T4AN0M0 | cM: 0 | pN: 0 | pT: 4a": {
        "Tumor Stage": "II",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IIA | TNM_Stage: T3N0M0 | cM: 0 | pN: 0 | pT: 3": {
        "Tumor Stage": "II A",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IIB | TNM_Stage: T1BN1M0 | cM: 0 | pN: 1 | pT: 1b": {
        "Tumor Stage": "II B",
        "TNM: Tumor": "pT1b",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IIB | TNM_Stage: T3N0M0 | cM: 0 | pM: x | pN: 0 | pT: 3": {
        "Tumor Stage": "II B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IIB | TNM_Stage: T3N1M0 | cM: 0 | pN: 1 | pT: 3": {
        "Tumor Stage": "II B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: III | TNM_Stage: T1N1M0 | cM: 0 | pN: 1 | pT: 1": {
        "Tumor Stage": "III",
        "TNM: Tumor": "pT1",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: III | TNM_Stage: T2N1M0 | cM: 0 | pN: 1 | pT: 2": {
        "Tumor Stage": "III",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: III | TNM_Stage: T3N1M0 | cM: 0 | pN: 1 | pT: 3": {
        "Tumor Stage": "III",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IIIA | TNM_Stage: T1N1AM0 | cM: 0 | pN: 1a | pT: 1": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "pT1",
        "TNM: Node": "pN1a",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IIIA | TNM_Stage: T1N2AM0 | cM: 0 | pN: 2a | pT: 1": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "pT1",
        "TNM: Node": "pN2a",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IIIA | TNM_Stage: T2N1BM0 | cM: 0 | pN: 1b | pT: 2": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN1b",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IIIA | TNM_Stage: T2N2AM0 | cM: 0 | pM: x | pN: 2a | pT: 2": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN2a",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IIIA | TNM_Stage: T3N0M0 | cM: 0 | pM: 0 | pN: X | pT: 3": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IIIA | TNM_Stage: T3aNxM0 | cM: 0 | pN: x | pT: 3a": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "pT3a",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IIIA | TNM_Stage: T4AN1M0 | cM: x | pN: 1 | pT: 4a": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cMx",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IIIA | TNM_Stage: T4AN2M0 | cM: 0 | pN: 2 | pT: 4a": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN2",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IIIB | TNM_Stage: T1BN1M0 | cM: 0 | pM: 0 | pN: X | pT: 1b": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT1b",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IIIB | TNM_Stage: T2N2AM0 | cM: 0 | pN: 2a | pT: 2": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN2a",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IIIB | TNM_Stage: T2N2BM0 | cM: 0 | pN: 2b | pT: 2": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN2b",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IIIB | TNM_Stage: T3N1AM0 | cM: 0 | pM: x | pN: 1a | pT: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1a",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IIIB | TNM_Stage: T3N1AM0 | cM: 0 | pN: 1a | pT: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1a",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IIIB | TNM_Stage: T3N1BM0 | cM: 0 | pN: 1b | pT: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1b",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IIIB | TNM_Stage: T3N1BM0 | cM: 0/x | pN: 1b | pT: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1b",
        "TNM: Metastasis": "cM0/x",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IIIB | TNM_Stage: T3N1CM0 | cM: 0 | pM: x | pN: 1c | pT: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1c",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IIIB | TNM_Stage: T3N1CM0 | cM: 0 | pN: 1c | pT: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1c",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IIIB | TNM_Stage: T3N1M0 | cM: 0 | pM: 0 | pN: 1 | pT: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IIIB | TNM_Stage: T3N2AM0 | cM: 0 | pM: x | pN: 1a | pT: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1a",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IIIB | TNM_Stage: T3N2M0 | cM: 0 | pN: 2 | pT: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN2",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IIIB | TNM_Stage: T3N3AM0 | cM: 0 | pM: 0 | pN: 3a | pT: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN3a",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IIIB | TNM_Stage: T4AN1BM0 | cM: 0 | pM: x | pN: 1b | pT: 4a": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN1b",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IIIB | TNM_Stage: T4AN1BM0 | cM: 0 | pN: 1b | pT: 4a": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN1b",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IIIB | TNM_Stage: T4N0M0 | cM: 0 | pM: 0 | pN: 0 | pT: 4": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT4",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IIIB | TNM_Stage: T4N1BM0 | cM: 0 | pN: 1b | pT: 4a": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN1b",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IIIB | TNM_Stage: T4aN1AM0 | cM: 0 | pM: x | pN: 1a | pT: 4a": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN1a",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IIIC | TNM_Stage: T2BN1M0 | cM: 0 | pM: 0 | pN: 1 | pT: 2b": {
        "Tumor Stage": "III C",
        "TNM: Tumor": "pT2b",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IIIC | TNM_Stage: T3N1M0 | cM: 0 | pM: 0 | pN: 1 | pT: 3": {
        "Tumor Stage": "III C",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IIIC | TNM_Stage: T4N2M0 (clinical stage at dx) | cM: 0 | pN: 1a | pT: 2": {
        "Tumor Stage": "III C",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN1a",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IVA | TNM_Stage: T2BN2M0 | cM: 0 | pM: 0 | pN: 2 | pT: 2b": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT2b",
        "TNM: Node": "pN2",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IVA | TNM_Stage: T2N2BM0 | cM: 0 | pN: 2b | pT: 2": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN2b",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IVA | TNM_Stage: T2N2bM0 | cM: 0 | pM: X | pN: 2b | pT: 2": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN2b",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IVA | TNM_Stage: T2N2bM0 | cM: 0 | pM: x | pN: 2b | pT: 2": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN2b",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IVA | TNM_Stage: T2N2bM0 | cM: 0 | pN: 2b | pT: 2": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN2b",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IVA | TNM_Stage: T3N1AM1A | cM: 1a (liver) | pM: x | pN: 1a | pT: 3": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1a",
        "TNM: Metastasis": "cM1a (liver), pMx",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IVA | TNM_Stage: T3N1AM1A | cM: 1a (liver) | pN: 1a | pT: 3": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1a",
        "TNM: Metastasis": "cM1a (liver)",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IVA | TNM_Stage: T3N2CM0 | cM: 0": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "T3",
        "TNM: Node": "N2c",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IVA | TNM_Stage: T3N2bM0 | cM: 0 | pN: 2b | pT: 3": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN2b",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IVA | TNM_Stage: T4AN1AM1A | cM: 1a (liver) | pM: 1a (liver) | pN: 1a | pT: 4a": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN1a",
        "TNM: Metastasis": "cM1a (liver), pM1a (liver)",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IVA | TNM_Stage: T4AN2M0 | cM: 0 | pM: 0 | pN: 2 | pT: 4a": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN2",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IVA | TNM_Stage: T4a(m)N2bM0 | cM: 0 | pN: 2b | pT: 4a(m)": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT4a(m)",
        "TNM: Node": "pN2b",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IVA | TNM_Stage: T4aN1M0 | cM: 0 | pM: x | pN: 1 | pT: 4a": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IVA | TNM_Stage: T4aN2bM0 | cM: 0 | pN: 2b | pT: 4a": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN2b",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IVA | TNM_Stage: T4aN2cM0 | cM: 0 | pN: 2c | pT: 4a": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN2c",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IVA | TNM_Stage: T4aN3M0 | cM: 0 | pN: 3 | pT: 4a": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN3",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IVB | TNM_Stage: T2N3bM0 | cM: 0 | pN: 3b | pT: 2": {
        "Tumor Stage": "IV B",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN3b",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IVB | TNM_Stage: T4AN3BM0 | cM: 0 | pM: x | pN: 3b | pT: 4a": {
        "Tumor Stage": "IV B",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN3b",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IVB | TNM_Stage: T4aN3bM0 | cM: 0 | pM: x | pN: 3b | pT: 4a": {
        "Tumor Stage": "IV B",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN3b",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IVC | TNM_Stage: T3N2AM1C | cM: 0 | pM: 1c (omentum) | pN: 2a | pT: 3": {
        "Tumor Stage": "IV C",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN2a",
        "TNM: Metastasis": "cM0, pM1c (omentum)",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M | StageAtDX: IVC | TNM_Stage: T4AN2BM1C | cM: 1c (peritoneum, ovaries) | pM: 1c (peritoneum, ovaries) | pN: 2b | pT: 4a": {
        "Tumor Stage": "IV C",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN2b",
        "TNM: Metastasis": "cM1c (peritoneum, ovaries), pM1c (peritoneum, ovaries)",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: M/P | StageAtDX: IIB | TNM_Stage: T2BN0M0 | cM: 0 | pM: 0 | pN: 0 | pT: 2b": {
        "Tumor Stage": "II B",
        "TNM: Tumor": "pT2b",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "2.5",
        "Additional tumor classification": "",
    },
    "Grade: M/P | StageAtDX: IIB | TNM_Stage: T2N1M0 | cM: 0 | pM: 0 | pN: 1 | pT: 2": {
        "Tumor Stage": "II B",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "2.5",
        "Additional tumor classification": "",
    },
    "Grade: M/P | StageAtDX: IIB | TNM_Stage: T3N1M0 | cM: 0 | pN: 1 | pT: 3": {
        "Tumor Stage": "II B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0",
        "Grade": "2.5",
        "Additional tumor classification": "",
    },
    "Grade: M/P | StageAtDX: III | TNM_Stage: T2N2M0 | cM: 0 | pN: 2 | pT: 2": {
        "Tumor Stage": "III",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN2",
        "TNM: Metastasis": "cM0",
        "Grade": "2.5",
        "Additional tumor classification": "",
    },
    "Grade: M/P | StageAtDX: IIIA | TNM_Stage: T3N2M0 | cM: 0 | pM: x | pN: 2 | pT: 3": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN2",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "2.5",
        "Additional tumor classification": "",
    },
    "Grade: M/P | StageAtDX: IIIA | TNM_Stage: T3N2M0 | cM: 0 | pN: 2 | pT: 3": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN2",
        "TNM: Metastasis": "cM0",
        "Grade": "2.5",
        "Additional tumor classification": "",
    },
    "Grade: M/P | StageAtDX: IIIA | TNM_Stage: T4AN2M0 | cM: 0 | pM: x | pN: 2 | pT: 4a": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN2",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "2.5",
        "Additional tumor classification": "",
    },
    "Grade: M/P | StageAtDX: IIIB | TNM_Stage: T4AN3AM0 | cM: 0 | pM: x | pN: 3a | pT: 4a": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN3a",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "2.5",
        "Additional tumor classification": "",
    },
    "Grade: M/P | StageAtDX: IIIC | TNM_Stage: T3N3M0 | cM: 0 | pN: 3 | pT: 3": {
        "Tumor Stage": "III C",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN3",
        "TNM: Metastasis": "cM0",
        "Grade": "2.5",
        "Additional tumor classification": "",
    },
    "Grade: M/P | StageAtDX: IVA | TNM_Stage: T4aN2cM0 | cM: 0 | pN: 2c | pT: 4a": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN2c",
        "TNM: Metastasis": "cM0",
        "Grade": "2.5",
        "Additional tumor classification": "",
    },
    "Grade: M/W | StageAtDX: IB | TNM_Stage: T2N0M0 | cM: 0 | pN: 0 | pT: 2": {
        "Tumor Stage": "I B",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0",
        "Grade": "1.5",
        "Additional tumor classification": "",
    },
    "Grade: M/W | StageAtDX: IIA | TNM_Stage: T3N0M0 | cM: 0 | pM: 0 | pN: 0 | pT: 3": {
        "Tumor Stage": "II A",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "1.5",
        "Additional tumor classification": "",
    },
    "Grade: M/W | StageAtDX: III | TNM_Stage: T2N2M0 | cM: 0 | pN: 2 | pT: 2": {
        "Tumor Stage": "III",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN2",
        "TNM: Metastasis": "cM0",
        "Grade": "1.5",
        "Additional tumor classification": "",
    },
    "Grade: M/W | StageAtDX: IIIA | TNM_Stage: T2N1AM0 | cM: 0 | pN: 1a | pT: 2": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN1a",
        "TNM: Metastasis": "cM0",
        "Grade": "1.5",
        "Additional tumor classification": "",
    },
    "Grade: M/W | StageAtDX: IIIB | TNM_Stage: T3N1AM0 | cM: 0 | pM: x | pN: 1a | pT: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1a",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "1.5",
        "Additional tumor classification": "",
    },
    "Grade: M/W | StageAtDX: IIIB | TNM_Stage: T3N1AM0 | cM: 0 | pN: 1a | pT: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1a",
        "TNM: Metastasis": "cM0",
        "Grade": "1.5",
        "Additional tumor classification": "",
    },
    "Grade: M/W | StageAtDX: IIIB | TNM_Stage: T3N1BM0 | cM: 0 | pN: 1b | pT: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1b",
        "TNM: Metastasis": "cM0",
        "Grade": "1.5",
        "Additional tumor classification": "",
    },
    "Grade: M/W | StageAtDX: IIIB | TNM_Stage: T3N1CM0 | cM: 0 | pN: 1c | pT: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1c",
        "TNM: Metastasis": "cM0",
        "Grade": "1.5",
        "Additional tumor classification": "",
    },
    "Grade: M/W | StageAtDX: IIIB | TNM_Stage: T3N2BM0 | cM: x | pN: 2b | pT: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN2b",
        "TNM: Metastasis": "cMx",
        "Grade": "1.5",
        "Additional tumor classification": "",
    },
    "Grade: M/W | StageAtDX: IVA | TNM_Stage: T4AN1AM1A | cM: 1a (liver) | pN: 1a | pT: 4a": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN1a",
        "TNM: Metastasis": "cM1a (liver)",
        "Grade": "1.5",
        "Additional tumor classification": "",
    },
    "Grade: M/W | StageAtDX: IVC | TNM_Stage: T3N2AM1C | cM: 1c (liver, peritoneum) | pN: 2a | pT: 3": {
        "Tumor Stage": "IV C",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN2a",
        "TNM: Metastasis": "cM1c (liver, peritoneum)",
        "Grade": "1.5",
        "Additional tumor classification": "",
    },
    "Grade: N.A | M: N.A | N: N.A | Stage: N.A | T: N.A": {
        "Tumor Stage": "",
        "TNM: Tumor": "",
        "TNM: Node": "",
        "TNM: Metastasis": "",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "Grade: N/A | Stage_at_DX: 1c": {
        "Tumor Stage": "1 c",
        "TNM: Tumor": "",
        "TNM: Node": "",
        "TNM: Metastasis": "",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "Grade: N/A | Stage_at_DX: 2b": {
        "Tumor Stage": "2 b",
        "TNM: Tumor": "",
        "TNM: Node": "",
        "TNM: Metastasis": "",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "Grade: Not stated | M: Mx | N: Nx | T: Tx": {
        "Tumor Stage": "",
        "TNM: Tumor": "Tx",
        "TNM: Node": "Nx",
        "TNM: Metastasis": "Mx",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: I | TNM_Stage: T1N1M0 | cM: 0 | pN: 1 | pT: 1": {
        "Tumor Stage": "I",
        "TNM: Tumor": "pT1",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: I | TNM_Stage: T2N1M0 | cM: 0 | pN: 1 | pT: 2": {
        "Tumor Stage": "I",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: I | TNM_Stage: T2N1M0 | cM: 0 | pN: x | pT: 2": {
        "Tumor Stage": "I",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IB | TNM_Stage: T1bN0M0 | cM: 0 | pM: 0 | pN: 0 | pT: 1b": {
        "Tumor Stage": "I B",
        "TNM: Tumor": "pT1b",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: II | TNM_Stage: T2AN0M0 | cM: 0 | pM: 0 | pN: 0 | pT: 2a": {
        "Tumor Stage": "II",
        "TNM: Tumor": "pT2a",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: II | TNM_Stage: T2BN0M0 | cM: 0 | pM: 0 | pN: 0 | pT: 2b": {
        "Tumor Stage": "II",
        "TNM: Tumor": "pT2b",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: II | TNM_Stage: T2N0M0 | cM: 0 | pM: 0 | pN: 0 | pT: 2": {
        "Tumor Stage": "II",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: II | TNM_Stage: T2NxM0 | cM: 0 | pM: 0 | pN: x | pT: 2": {
        "Tumor Stage": "II",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: II | TNM_Stage: T2NxM0 | cM: 0 | pN: x | pT: 2": {
        "Tumor Stage": "II",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIA | TNM_Stage: T2AN0M0 | cM: 0 | pM: 0 | pN: 0 | pT: 2a": {
        "Tumor Stage": "II A",
        "TNM: Tumor": "pT2a",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIA | TNM_Stage: T2N0M0 | cM: 0 | pM: 0 | pN: 0 | pT: 2": {
        "Tumor Stage": "II A",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIA | TNM_Stage: T2N1M0 | cM: 0 | pN: 1 | pT: 2": {
        "Tumor Stage": "II A",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIA | TNM_Stage: T3N0M0 | cM: 0 | pM: x | pN: 0 | pT: 3": {
        "Tumor Stage": "II A",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIA | TNM_Stage: T3N0M0 | cM: 0 | pN: 0 | pT: 3": {
        "Tumor Stage": "II A",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIB | TNM_Stage: T2BN0M0 | cM: 0 | pM: 0 | pN: 0 | pT: 2b": {
        "Tumor Stage": "II B",
        "TNM: Tumor": "pT2b",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIB | TNM_Stage: T3N0M0 | cM: 0 | pM: 0 | pN: 0 | pT: 3": {
        "Tumor Stage": "II B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIB | TNM_Stage: T3N1M0 | cM: 0 | pN: 1 | pT: 3": {
        "Tumor Stage": "II B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: III | TNM_Stage: T2N1M0 | cM: 0 | pM: x | pN: 1 | pT: 2": {
        "Tumor Stage": "III",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: III | TNM_Stage: T2N2M0 | cM: 0 | pN: 2 | pT: 2": {
        "Tumor Stage": "III",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN2",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: III | TNM_Stage: T3N1M0 | cM: 0 | pM: x | pN: 1 | pT: 3": {
        "Tumor Stage": "III",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: III | TNM_Stage: T3N2M0 | cM: 0 | pM: 0 | pN: 2 | pT: 3": {
        "Tumor Stage": "III",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN2",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIIA | TNM_Stage: T3N0M0 | cM: 0 | pM: 0 | pN: 0 | pT: 3": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIIA | TNM_Stage: T3N2M0 | cM: 0 | pN: 2 | pT: 3": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN2",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIIA | TNM_Stage: T3NxM0 | cM: 0 | pM: 0 | pN: x | pT: 3": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIIA | TNM_Stage: T3bN0M0 | cM: 0 | pN: 0 | pT: 3b": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "pT3b",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIIA | TNM_Stage: T3bNxM0 | cM: 0 | pN: x | pT: 3b": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "pT3b",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIIA | TNM_Stage: T4AN2M0 | cM: 0 | pN: 2 | pT: 4a": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN2",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIIB | TNM_Stage: T2BN1M0 | cM: 0 | pM: 0 | pN: 1 | pT: 2b": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT2b",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIIB | TNM_Stage: T2BN1M0 | cM: 0 | pN: 1 | pT: 2b": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT2b",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIIB | TNM_Stage: T2N1M0 | cM: 0 | pM: 0 | pN: X | pT: 2": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIIB | TNM_Stage: T3BN1M0 | cM: 0 | pM: 0 | pN: 1 | pT: 3b": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT3b",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIIB | TNM_Stage: T3N1AM0 | cM: 0 | pN: 1a | pT: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1a",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIIB | TNM_Stage: T3N1BM0 | cM: 0 | pN: 1b | pT: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1b",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIIB | TNM_Stage: T3N1M0 | cM: 0 | pM: 0 | pN: 1 | pT: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIIB | TNM_Stage: T3N1M0 | pM: x | pN: 1 | pT: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "pMx",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIIB | TNM_Stage: T3N3AM0 | cM: x | pM: x | pN: 3a | pT: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN3a",
        "TNM: Metastasis": "cMx, pMx",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIIB | TNM_Stage: T4A3AM0 | cM: 0 | pN: 3a | pT: 4a": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN3a",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIIB | TNM_Stage: T4AN1AM0 | cM: 0 | pM: x | pN: 1a | pT: 4a": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN1a",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIIB | TNM_Stage: T4AN3AM0 | cM: 0 | pN: 3a | pT: 4a": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN3a",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIIB | TNM_Stage: T4NxM0 | cM: 0 | pN: x | pT: 4": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT4",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIIC | TNM_Stage: T3N1M0 | cM: 0 | pM: 0 | pN: 1 | pT: 3": {
        "Tumor Stage": "III C",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIIC | TNM_Stage: T3N2BM0 | cM: 0 | pN: 2b | pT: 3": {
        "Tumor Stage": "III C",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN2b",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIIC | TNM_Stage: T4AN2BM0 | cM: 0 | pM: x | pN: 2b | pT: 4a": {
        "Tumor Stage": "III C",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN2b",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIIC | TNM_Stage: T4AN2BM0 | cM: 0 | pN: 2b | pT: 4a": {
        "Tumor Stage": "III C",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN2b",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIIC | TNM_Stage: T4AN2BM0 | cM: x | pN: 2b | pT: 4a": {
        "Tumor Stage": "III C",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN2b",
        "TNM: Metastasis": "cMx",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIIC | TNM_Stage: T4AN3AM0 | cM: 0 | pM: x | pN: 3a | pT: 4a": {
        "Tumor Stage": "III C",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN3a",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIIC | TNM_Stage: T4AN3AM0 | cM: 0 | pN: 3a | pT: 4a": {
        "Tumor Stage": "III C",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN3a",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIIC | TNM_Stage: T4AN3BM0 | cM: 0 | pN: 3b | pT: 4a": {
        "Tumor Stage": "III C",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN3b",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIIC | TNM_Stage: T4BN2AM0 | cM: 1 (?peritoneum) | pM: 0 | pN: 2a | pT: 4b": {
        "Tumor Stage": "III C",
        "TNM: Tumor": "pT4b",
        "TNM: Node": "pN2a",
        "TNM: Metastasis": "cM1 (?peritoneum), pM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIIC | TNM_Stage: T4BN2BM0 | cM: 0 | pN: 2b | pT: 4b": {
        "Tumor Stage": "III C",
        "TNM: Tumor": "pT4b",
        "TNM: Node": "pN2b",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IIIC | TNM_Stage: T4BN2BM0 | cM: 1a (ovary/fallopian tube) | pM: M0 - direct invasion of ovary/fallopian tube | pN: 2b | pT: 4b": {
        "Tumor Stage": "III C",
        "TNM: Tumor": "pT4b",
        "TNM: Node": "pN2b",
        "TNM: Metastasis": "cM1a (ovary/fallopian tube), pM0 - direct invasion of ovary/fallopian tube",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IV | TNM_Stage: T4AN3M1 (surrounding fat, peritoneum) | cM: 0 | pM: 1 | pN: 3 | pT: 4a": {
        "Tumor Stage": "IV",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN3",
        "TNM: Metastasis": "cM0, pM1",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IVA | TNM_Stage: T1(mi)N2bM0 | cM: 0 | pN: 2b | pT: 1(mi)": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT1(mi)",
        "TNM: Node": "pN2b",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IVA | TNM_Stage: T2BN2M0 | cM: 0 | pM: 0 | pN: 2 | pT: 2b": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT2b",
        "TNM: Node": "pN2",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IVA | TNM_Stage: T2N2BM0 | cM: 0 | pM: x | pN: 2b | pT: 2": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN2b",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IVA | TNM_Stage: T2N2CM0 | cM: 0 | pN: 2c | pT: 2": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN2c",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IVA | TNM_Stage: T2N2bM0 | cM: 0 | pM: 0 | pN: 2b | pT: 2": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN2b",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IVA | TNM_Stage: T2N2bM0 | cM: 0 | pM: X | pN: 2b | pT: 2": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN2b",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IVA | TNM_Stage: T2N2cM0 | cM: 0 | pM: X | pN: 2c | pT: 2": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN2c",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IVA | TNM_Stage: T2N2cM0 | cM: 0 | pN: 2c | pT: 2": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN2c",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IVA | TNM_Stage: T3N2M0 | cM: 0 | pM: 0 | pN: 2 | pT: 3": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN2",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IVA | TNM_Stage: T3N2aM0 | cM: 0 | pM: 0 | pN: 2a | pT: 3": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN2a",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IVA | TNM_Stage: T3N2bM0 | cM: 0 | pN: 2b | pT: 3": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN2b",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IVA | TNM_Stage: T3N3M0 | cM: 0 | pM: x | pN: 3 | pT: 3": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN3",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IVA | TNM_Stage: T3N3M0 | cM: 0 | pN: 3 | pT: 3": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN3",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IVA | TNM_Stage: T4AN2BM0 | cM: 0 | pM: x | pN: 2b | pT: 4a": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN2b",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IVA | TNM_Stage: T4AN2BM0 | cM: 0 | pN: 2b | pT: 4a": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN2b",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IVA | TNM_Stage: T4AN2BM1A | cM: 1A (distant lymph nodes) | pN: 2b | pT: 4a": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN2b",
        "TNM: Metastasis": "cM1A (distant lymph nodes)",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IVA | TNM_Stage: T4N2bM0 | cM: 0 | pM: X | pN: 2b | pT: 4": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT4",
        "TNM: Node": "pN2b",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IVA | TNM_Stage: T4aN2aM0 | cM: 0 | pN: 2a | pT: 4a": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN2a",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IVA | TNM_Stage: T4aN2cM0 | cM: 0 | pN: 2c | pT: 4a": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN2c",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IVB | TNM_Stage: T2N3bM0 | cM: 0 | pM: x | pN: 3b | pT: 2": {
        "Tumor Stage": "IV B",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN3b",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IVB | TNM_Stage: T2N3bM0 | cM: 0 | pN: 3b | pT: 2": {
        "Tumor Stage": "IV B",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN3b",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IVB | TNM_Stage: T3N0M1 | cM: 1 (liver) | pM: X | pN: X | pT: 3": {
        "Tumor Stage": "IV B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM1 (liver), pMx",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IVB | TNM_Stage: T3N3bM0 | cM: 0 | pN: 3b | pT: 3": {
        "Tumor Stage": "IV B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN3b",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IVB | TNM_Stage: T4aN3bM0 | cM: 0 | pN: 3b | pT: 4a": {
        "Tumor Stage": "IV B",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN3b",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IVB | TNM_Stage: T4bN3bM0 | cM: 0 | pN: 3b | pT: 4b": {
        "Tumor Stage": "IV B",
        "TNM: Tumor": "pT4b",
        "TNM: Node": "pN3b",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IVC | TNM_Stage: T4AN2BM1C | cM: 0 | pM: 1c (omentum) | pN: 2b | pT: 4a": {
        "Tumor Stage": "IV C",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN2b",
        "TNM: Metastasis": "cM0, pM1c (omentum)",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IVC | TNM_Stage: T4AN2BM1C | cM: 1c (peritoneum) | pM: x | pN: 2b | pT: 4a": {
        "Tumor Stage": "IV C",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN2b",
        "TNM: Metastasis": "cM1c (peritoneum), pMx",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IVC | TNM_Stage: T4BN1BM1C | cM: 1c (peritoneum) | pM: 1c (peritoneum) | pN: 1b | pT: 4b": {
        "Tumor Stage": "IV C",
        "TNM: Tumor": "pT4b",
        "TNM: Node": "pN1b",
        "TNM: Metastasis": "cM1c (peritoneum), pM1c (peritoneum)",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IVC | TNM_Stage: T4BN2BM1C | cM: 1c (peritoneum, liver, ileocolic nodes) | pM: x | pN: 2b | pT: 4b": {
        "Tumor Stage": "IV C",
        "TNM: Tumor": "pT4b",
        "TNM: Node": "pN2b",
        "TNM: Metastasis": "cM1c (peritoneum, liver, ileocolic nodes), pMx",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: P | StageAtDX: IVC | TNM_Stage: T4aN2bM1 | cM: 1 (Lung)": {
        "Tumor Stage": "IV C",
        "TNM: Tumor": "T4a",
        "TNM: Node": "N2b",
        "TNM: Metastasis": "cM1 (Lung)",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: T3N0M0 | StageAtDX: II | TNM_Stage: P | cM: 0 | pN: 0 | pT: 3": {
        "Tumor Stage": "II",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0",
        "Grade": "3",
        "Additional tumor classification": "",
    },
    "Grade: T3N1M0 | StageAtDX: IIIB | TNM_Stage: M | cM: 0 | pN: 1 | pT: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: T3N3M0 | StageAtDX: IVA | TNM_Stage: M | cM: 0 | pN: 3 | pT: 3": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN3",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: U&P | StageAtDX: IIIB | TNM_Stage: T4BN1M0 | cM: 0 | pN: 1 | pT: 4b": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT4b",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0",
        "Grade": "3.5",
        "Additional tumor classification": "",
    },
    "Grade: W | StageAtDX: I | TNM_Stage: T1N0M0 | cM: 0 | pN: 0 | pT: 1": {
        "Tumor Stage": "I",
        "TNM: Tumor": "pT1",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0",
        "Grade": "1",
        "Additional tumor classification": "",
    },
    "Grade: W | StageAtDX: I | TNM_Stage: T1N0M0 | cM: 0 | pN: X | pT: 1": {
        "Tumor Stage": "I",
        "TNM: Tumor": "pT1",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM0",
        "Grade": "1",
        "Additional tumor classification": "",
    },
    "Grade: W | StageAtDX: IA | TNM_Stage: T1AN0M0 | cM: 0 | pM: 0 | pN: 0 | pT: 1a": {
        "Tumor Stage": "I A",
        "TNM: Tumor": "pT1a",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "1",
        "Additional tumor classification": "",
    },
    "Grade: W | StageAtDX: IB | TNM_Stage: T1aNxM0 | cM: 0 | pN: x | pT: 1a": {
        "Tumor Stage": "I B",
        "TNM: Tumor": "pT1a",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM0",
        "Grade": "1",
        "Additional tumor classification": "",
    },
    "Grade: W | StageAtDX: IB | TNM_Stage: T1bNxM0 | cM: 0 | pN: x | pT: 1b": {
        "Tumor Stage": "I B",
        "TNM: Tumor": "pT1b",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM0",
        "Grade": "1",
        "Additional tumor classification": "",
    },
    "Grade: W | StageAtDX: II | TNM_Stage: T2N0M0 | cM: 0 | pM: 0 | pN: 0 | pT: 2": {
        "Tumor Stage": "II",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "1",
        "Additional tumor classification": "",
    },
    "Grade: W | StageAtDX: II | TNM_Stage: T2N0M0 | cM: 0 | pM: X | pN: 0 | pT: 2": {
        "Tumor Stage": "II",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "1",
        "Additional tumor classification": "",
    },
    "Grade: W | StageAtDX: II | TNM_Stage: T2N0M0 | cM: 0 | pN: 0 | pT: 2": {
        "Tumor Stage": "II",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0",
        "Grade": "1",
        "Additional tumor classification": "",
    },
    "Grade: W | StageAtDX: II | TNM_Stage: T2N0M0 | cM: 0 | pN: X | pT: 2": {
        "Tumor Stage": "II",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM0",
        "Grade": "1",
        "Additional tumor classification": "",
    },
    "Grade: W | StageAtDX: II | TNM_Stage: T3N0M0 | cM: 0 | pN: 0 | pT: 3": {
        "Tumor Stage": "II",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0",
        "Grade": "1",
        "Additional tumor classification": "",
    },
    "Grade: W | StageAtDX: IIA | TNM_Stage: T1BN2M0 | cM: 0 | pN: 2 | pT: 1b": {
        "Tumor Stage": "II A",
        "TNM: Tumor": "pT1b",
        "TNM: Node": "pN2",
        "TNM: Metastasis": "cM0",
        "Grade": "1",
        "Additional tumor classification": "",
    },
    "Grade: W | StageAtDX: IIA | TNM_Stage: T3N0M0 | cM: 0 | pN: 0 | pT: 3": {
        "Tumor Stage": "II A",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0",
        "Grade": "1",
        "Additional tumor classification": "",
    },
    "Grade: W | StageAtDX: III | TNM_Stage: T1N1M0 | cM: 0 | pN: 1 | pT: 1": {
        "Tumor Stage": "III",
        "TNM: Tumor": "pT1",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0",
        "Grade": "1",
        "Additional tumor classification": "",
    },
    "Grade: W | StageAtDX: III | TNM_Stage: T2N1M0 | cM: 0 | pN: 1 | pT: 2": {
        "Tumor Stage": "III",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0",
        "Grade": "1",
        "Additional tumor classification": "",
    },
    "Grade: W | StageAtDX: IIIB | TNM_Stage: T3N2AM0 | cM: 0 | pN: 2a | pT: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN2a",
        "TNM: Metastasis": "cM0",
        "Grade": "1",
        "Additional tumor classification": "",
    },
    "Grade: W | StageAtDX: IIIB | TNM_Stage: T4AN1BM0 | cM: 0 | pN: 1b | pT: 4a": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN1b",
        "TNM: Metastasis": "cM0",
        "Grade": "1",
        "Additional tumor classification": "",
    },
    "Grade: W | StageAtDX: IIIC | TNM_Stage: T3N2BM0 | cM: 0 | pN: 2b | pT: 3": {
        "Tumor Stage": "III C",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN2b",
        "TNM: Metastasis": "cM0",
        "Grade": "1",
        "Additional tumor classification": "",
    },
    "Grade: W | StageAtDX: IVA | TNM_Stage: T4aN1M0 | cM: 0 | pN: 1 | pT: 4a": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0",
        "Grade": "1",
        "Additional tumor classification": "",
    },
    "Grade: W | StageAtDX: IVA | TNM_Stage: T4aN2b(mi)M0 | cM: 0 | pM: x | pN: 2b(mi) | pT: 4a": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN2b(mi)",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "1",
        "Additional tumor classification": "",
    },
    "Grade: W | TNM_Stage: TXN0M0 | cM: 0 | pN: 0 | pT: X": {
        "Tumor Stage": "",
        "TNM: Tumor": "pTx",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0",
        "Grade": "1",
        "Additional tumor classification": "",
    },
    "Grade: W+M | StageAtDX: I | TNM_Stage: T1NxM0 | cM: 0 | pM: x | pN: x | pT: 1": {
        "Tumor Stage": "I",
        "TNM: Tumor": "pT1",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "1.5",
        "Additional tumor classification": "",
    },
    "Grade: W/M | StageAtDX: IIIA | TNM_Stage: T2N1AM0 | cM: x | pM: x | pN: 1a | pT: 2": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN1a",
        "TNM: Metastasis": "cMx, pMx",
        "Grade": "1.5",
        "Additional tumor classification": "",
    },
    "Grade: W/M | StageAtDX: IIIA | TNM_Stage: T3N1M0 | cM: 0 | pN: 1 | pT: 3": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0",
        "Grade": "1.5",
        "Additional tumor classification": "",
    },
    "Grade: W/M | StageAtDX: IIIB | TNM_Stage: T3N1AM0 | cM: 0 | pM: x | pN: 1a | pT: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1a",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "1.5",
        "Additional tumor classification": "",
    },
    "Grade: W/M | StageAtDX: IIIB | TNM_Stage: T3N1AM0 | cM: 0 | pN: 1a | pT: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1a",
        "TNM: Metastasis": "cM0",
        "Grade": "1.5",
        "Additional tumor classification": "",
    },
    "Grade: W/M | StageAtDX: IIIB | TNM_Stage: T3N1BM0 | cM: 0 | pM: x | pN: 1b | pT: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1b",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "1.5",
        "Additional tumor classification": "",
    },
    "Grade: W/M | StageAtDX: IIIB | TNM_Stage: T3N1CM0 | cM: 0 | pM: x | pN: 1c | pT: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1c",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "1.5",
        "Additional tumor classification": "",
    },
    "Grade: W/M | StageAtDX: IIIB | TNM_Stage: T3N1M0 | cM: 0 | pM: x | pN: 1 | pT: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "1.5",
        "Additional tumor classification": "",
    },
    "Grade: W/M | StageAtDX: IIIB | TNM_Stage: T3N2BM0 | cM: 0 | pN: 2b | pT: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN2b",
        "TNM: Metastasis": "cM0",
        "Grade": "1.5",
        "Additional tumor classification": "",
    },
    "Grade: W/M | StageAtDX: IIIC | TNM_Stage: T3N2BM0 | cM: 0 | pM: x | pN: 2b | pT: 3": {
        "Tumor Stage": "III C",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN2b",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "1.5",
        "Additional tumor classification": "",
    },
    "Grade: W/M | StageAtDX: IVA | TNM_Stage: T3N1BM1A | cM: 1A (lung) | pN: 1b | pT: 3": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1b",
        "TNM: Metastasis": "cM1A (lung)",
        "Grade": "1.5",
        "Additional tumor classification": "",
    },
    "Grade: W/M | StageAtDX: IVA | TNM_Stage: T3N2AM1A | cM: 1a (liver) | pN: 2a | pT: 3": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN2a",
        "TNM: Metastasis": "cM1a (liver)",
        "Grade": "1.5",
        "Additional tumor classification": "",
    },
    "Grade: W/M | StageAtDX: IVA | TNM_Stage: T3N2BM1A | cM: 1a (liver) | pM: x | pN: 2b | pT: 3": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN2b",
        "TNM: Metastasis": "cM1a (liver), pMx",
        "Grade": "1.5",
        "Additional tumor classification": "",
    },
    "Grade: W/P | StageAtDX: IIA | TNM_Stage: T2N1M0 | cM: 0 | pN: 1 | pT: 2": {
        "Tumor Stage": "II A",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "Grade: W/P | StageAtDX: IIIC | TNM_Stage: T4AN3AM0 | cM: 0 | pM: x | pN: 3a | pT: 4a": {
        "Tumor Stage": "III C",
        "TNM: Tumor": "pT4a",
        "TNM: Node": "pN3a",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "2",
        "Additional tumor classification": "",
    },
    "M: 1 | N: 1 | StageAtDX: cT3a N1 M1 | T: 3a": {
        "Tumor Stage": "cT3a N1 M1",
        "TNM: Tumor": "T3a",
        "TNM: Node": "N1",
        "TNM: Metastasis": "M1",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "M: 1 | N: 1 | StageAtDX: cT3b N1 M1 | T: 3b": {
        "Tumor Stage": "cT3b N1 M1",
        "TNM: Tumor": "T3b",
        "TNM: Node": "N1",
        "TNM: Metastasis": "M1",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "M: x | N: 1 | StageAtDX: cTx N1 Mx | T: x": {
        "Tumor Stage": "cTx N1 Mx",
        "TNM: Tumor": "Tx",
        "TNM: Node": "N1",
        "TNM: Metastasis": "Mx",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "M: x | N: x | StageAtDX: cT3 Nx Mx | T: 3": {
        "Tumor Stage": "cT3 Nx Mx",
        "TNM: Tumor": "T3",
        "TNM: Node": "Nx",
        "TNM: Metastasis": "Mx",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "M: x | N: x | StageAtDX: cT4 Nx Mx | T: 4": {
        "Tumor Stage": "cT4 Nx Mx",
        "TNM: Tumor": "T4",
        "TNM: Node": "Nx",
        "TNM: Metastasis": "Mx",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "M: x | N: x | StageAtDX: pT3 Nx Mx | T: 3": {
        "Tumor Stage": "pT3 Nx Mx",
        "TNM: Tumor": "T3",
        "TNM: Node": "Nx",
        "TNM: Metastasis": "Mx",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "Stage At Dx: T2N2cM0": {
        "Tumor Stage": "T2N2cM0",
        "TNM: Tumor": "T2",
        "TNM: Node": "N2c",
        "TNM: Metastasis": "M0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "Stage: III": {
        "Tumor Stage": "III",
        "TNM: Tumor": "",
        "TNM: Node": "",
        "TNM: Metastasis": "",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "Stage: IIIA | TNM: T3N2M0": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "T3",
        "TNM: Node": "N2",
        "TNM: Metastasis": "M0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "Stage: IIIA | TNM: T3N2Mx": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "T3",
        "TNM: Node": "N2",
        "TNM: Metastasis": "Mx",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "Stage: IV": {
        "Tumor Stage": "IV",
        "TNM: Tumor": "",
        "TNM: Node": "",
        "TNM: Metastasis": "",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "Stage: IV-A": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "",
        "TNM: Node": "",
        "TNM: Metastasis": "",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "Stage: T1": {
        "Tumor Stage": "",
        "TNM: Tumor": "T1",
        "TNM: Node": "",
        "TNM: Metastasis": "",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "Stage: T1N0M0": {
        "Tumor Stage": "",
        "TNM: Tumor": "T1",
        "TNM: Node": "N0",
        "TNM: Metastasis": "M0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "Stage: T2": {
        "Tumor Stage": "",
        "TNM: Tumor": "T2",
        "TNM: Node": "",
        "TNM: Metastasis": "",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "Stage: T2N1M0": {
        "Tumor Stage": "",
        "TNM: Tumor": "T2",
        "TNM: Node": "N1",
        "TNM: Metastasis": "M0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "Stage: T2NIM1": {
        "Tumor Stage": "",
        "TNM: Tumor": "T2",
        "TNM: Node": "N1",
        "TNM: Metastasis": "M1",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "Stage: T3": {
        "Tumor Stage": "",
        "TNM: Tumor": "T3",
        "TNM: Node": "",
        "TNM: Metastasis": "",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "Stage: T4NxMx": {
        "Tumor Stage": "",
        "TNM: Tumor": "T4",
        "TNM: Node": "Nx",
        "TNM: Metastasis": "Mx",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "StageAtDX: !A | TNM_Stage: T1N0M0 | cM: 0 | pT: 1": {
        "Tumor Stage": "!A",
        "TNM: Tumor": "pT1",
        "TNM: Node": "N0",
        "TNM: Metastasis": "cM0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "StageAtDX: I | TNM_Stage: T1BN0M0 | cM: 0 | pM: 0 | pN: 0 | pT: 1b": {
        "Tumor Stage": "I",
        "TNM: Tumor": "pT1b",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "StageAtDX: I | TNM_Stage: T1BN0M0 | cM: 0 | pM: 0 | pN: 0 | pT: 1bm": {
        "Tumor Stage": "I",
        "TNM: Tumor": "pT1bm",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "StageAtDX: I | TNM_Stage: T1BN0M0 | cM: 0 | pM: 0 | pN: x | pT: 1b": {
        "Tumor Stage": "I",
        "TNM: Tumor": "pT1b",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "StageAtDX: I | TNM_Stage: T1BN1BM0 | cM: 0 | pM: 0 | pN: 1b | pT: 1bm": {
        "Tumor Stage": "I",
        "TNM: Tumor": "pT1bm",
        "TNM: Node": "pN1b",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "StageAtDX: I | TNM_Stage: T1N1M0 | cM: 0 | pM: x | pN: 1 | pT: 1": {
        "Tumor Stage": "I",
        "TNM: Tumor": "pT1",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0, pMx",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "StageAtDX: I | TNM_Stage: T1N1M0 | cM: 0 | pN: 1 | pT: 1": {
        "Tumor Stage": "I",
        "TNM: Tumor": "pT1",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "StageAtDX: I | TNM_Stage: T2N0M0 | cM: 0 | pM: 0 | pN: 0 | pT: 2": {
        "Tumor Stage": "I",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "StageAtDX: I | TNM_Stage: T2N1AM0 | cM: 0 | pM: 0 | pN: 1a | pT: 2m": {
        "Tumor Stage": "I",
        "TNM: Tumor": "pT2m",
        "TNM: Node": "pN1a",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "StageAtDX: I | TNM_Stage: T2N1M0 | cM: 0 | pM: 0 | pN: x | pT: 2": {
        "Tumor Stage": "I",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "StageAtDX: I | TNM_Stage: T2N1M0 | cM: 0 | pN: 1 | pT: 2": {
        "Tumor Stage": "I",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "StageAtDX: I | TNM_Stage: T3BN0M0 | cM: 0 | pM: 0 | pN: x | pT: 3b": {
        "Tumor Stage": "I",
        "TNM: Tumor": "pT3b",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "StageAtDX: I | TNM_Stage: T3BN1AM0 | cM: 0 | pM: 0 | pN: 1a | pT: 3bm": {
        "Tumor Stage": "I",
        "TNM: Tumor": "pT3bm",
        "TNM: Node": "pN1a",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "StageAtDX: IA | TNM_Stage: T1N0M0 | cM: 0 | pT: 1": {
        "Tumor Stage": "I A",
        "TNM: Tumor": "pT1",
        "TNM: Node": "N0",
        "TNM: Metastasis": "cM0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "StageAtDX: IB | TNM_Stage: T2N0M0 | cM: 0 | pN: 0 | pT: 2": {
        "Tumor Stage": "I B",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "StageAtDX: IB | TNM_Stage: T2N0M0 | cM: 0 | pT: 2": {
        "Tumor Stage": "I B",
        "TNM: Tumor": "pT2",
        "TNM: Node": "N0",
        "TNM: Metastasis": "cM0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "StageAtDX: II | TNM_Stage: T2N0M0 | cM: 0 | pM: 0 | pN: 0 | pT: 2": {
        "Tumor Stage": "II",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "StageAtDX: II | TNM_Stage: T2N2M0 | cM: 0 | pN: 2 | pT: 2": {
        "Tumor Stage": "II",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN2",
        "TNM: Metastasis": "cM0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "StageAtDX: II | TNM_Stage: T3BN0M0 | cM: 0 | pM: 0 | pN: 0 | pT: 3bm": {
        "Tumor Stage": "II",
        "TNM: Tumor": "pT3bm",
        "TNM: Node": "pN0",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "StageAtDX: II | TNM_Stage: T3N1M0 | cM: 0 | pN: 1 | pT: 3": {
        "Tumor Stage": "II",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "StageAtDX: II | cM: 0": {
        "Tumor Stage": "II",
        "TNM: Tumor": "",
        "TNM: Node": "",
        "TNM: Metastasis": "cM0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "StageAtDX: III | TNM_Stage: T3AN1M0 | cM: 0 | pM: 0 | pN: 1 | pT: 3a(m)": {
        "Tumor Stage": "III",
        "TNM: Tumor": "pT3a(m)",
        "TNM: Node": "pN1",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "StageAtDX: III | TNM_Stage: T4N2M0 | cM: 0 | pN: 2 | pT: 4": {
        "Tumor Stage": "III",
        "TNM: Tumor": "pT4",
        "TNM: Node": "pN2",
        "TNM: Metastasis": "cM0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "StageAtDX: IIIA | TNM_Stage: T2N0M0 | cM: 0 | pT: 2": {
        "Tumor Stage": "III A",
        "TNM: Tumor": "pT2",
        "TNM: Node": "N0",
        "TNM: Metastasis": "cM0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "StageAtDX: IIIB | TNM_Stage: T2N2BM0 | cM: 0 | pN: X | pT: 2": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pNx",
        "TNM: Metastasis": "cM0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "StageAtDX: IIIB | TNM_Stage: T3N1BM0 | cM: 0 | pN: 1b | pT: 3": {
        "Tumor Stage": "III B",
        "TNM: Tumor": "pT3",
        "TNM: Node": "pN1b",
        "TNM: Metastasis": "cM0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "StageAtDX: IIIC | TNM_Stage: T4AN3M0 | cM: 0 | pN: 3 | pT: 4A": {
        "Tumor Stage": "III C",
        "TNM: Tumor": "pT4A",
        "TNM: Node": "pN3",
        "TNM: Metastasis": "cM0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "StageAtDX: IS | TNM_Stage: T1N0M0 | cM: 0 | pT: 1": {
        "Tumor Stage": "I S",
        "TNM: Tumor": "pT1",
        "TNM: Node": "N0",
        "TNM: Metastasis": "cM0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "StageAtDX: IS | TNM_Stage: T2N0M0 | cM: 0 | pT: 2": {
        "Tumor Stage": "I S",
        "TNM: Tumor": "pT2",
        "TNM: Node": "N0",
        "TNM: Metastasis": "cM0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "StageAtDX: IV | TNM_Stage: T2N0M1 | cM: 1 (Bone) | pT: 2": {
        "Tumor Stage": "IV",
        "TNM: Tumor": "pT2",
        "TNM: Node": "N0",
        "TNM: Metastasis": "cM1 (Bone)",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "StageAtDX: IVA | TNM_Stage: T2N2aM0 | cM: 0 | pN: 2a | pT: 2": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT2",
        "TNM: Node": "pN2a",
        "TNM: Metastasis": "cM0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "StageAtDX: IVA | TNM_Stage: T3BN1BM0 | cM: 0 | pM: 0 | pN: 1b | pT: 3b": {
        "Tumor Stage": "IV A",
        "TNM: Tumor": "pT3b",
        "TNM: Node": "pN1b",
        "TNM: Metastasis": "cM0, pM0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "TNM: T2N1M0": {
        "Tumor Stage": "",
        "TNM: Tumor": "T2",
        "TNM: Node": "N1",
        "TNM: Metastasis": "M0",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "TNM: T3N1M1": {
        "Tumor Stage": "",
        "TNM: Tumor": "T3",
        "TNM: Node": "N1",
        "TNM: Metastasis": "M1",
        "Grade": "",
        "Additional tumor classification": "",
    },
    "WHO_Grade: IV": {
        "Tumor Stage": "",
        "TNM: Tumor": "",
        "TNM: Node": "",
        "TNM: Metastasis": "",
        "Grade": "4",
        "Additional tumor classification": "WHO Grade: IV",
    },
    "WHO_Grade_Primary: II": {
        "Tumor Stage": "",
        "TNM: Tumor": "",
        "TNM: Node": "",
        "TNM: Metastasis": "",
        "Grade": "2",
        "Additional tumor classification": "WHO Grade Primary: II",
    },
    "WHO_Grade_Primary: III": {
        "Tumor Stage": "",
        "TNM: Tumor": "",
        "TNM: Node": "",
        "TNM: Metastasis": "",
        "Grade": "3",
        "Additional tumor classification": "WHO Grade Primary: III",
    },
    "WHO_Grade_Primary: IV": {
        "Tumor Stage": "",
        "TNM: Tumor": "",
        "TNM: Node": "",
        "TNM: Metastasis": "",
        "Grade": "4",
        "Additional tumor classification": "WHO Grade Primary: IV",
    },
}


def clean_tnm_parameter(val, required_char):
    if pd.isna(val) or not str(val).strip():
        return ""

    val_str = str(val).strip()

    # If multiple values exist (e.g. "cM0, pM0" or "cM1a (liver), pMx")
    # Priority: search for pX first (pathological), then cX (clinical), then any X
    patterns = [
        rf"\b(y|p|c|r|a|u)?{required_char}[a-z0-9()]+",  # Match stage code like pT1a, cM0, pMx
        rf"\b{required_char}[a-z0-9()]+",  # Match simple T1, N0, M1
    ]

    for pattern in patterns:
        matches = re.findall(pattern, val_str, flags=re.IGNORECASE)
        if matches:
            # Reconstruct complete match from regex groups if needed
            full_match = re.search(
                rf"\b((?:y|p|c|r|a|u)?{required_char}[a-z0-9()]*)\b",
                val_str,
                re.IGNORECASE,
            )
            if full_match:
                match_val = full_match.group(1)
                # Ensure the core required letter (T/N/M) is uppercase
                return re.sub(
                    rf"({required_char})",
                    lambda m: m.group(1).upper(),
                    match_val,
                    flags=re.IGNORECASE,
                )

    # Fallback: If value is just a number/letter missing the leading character (e.g., '1a' for Tumor)
    clean_val = re.sub(r"[^a-zA-Z0-9()]", "", val_str)
    if clean_val and not clean_val.upper().startswith(
        ("T", "N", "M", "PT", "CT", "YT", "RT", "PN", "CN", "PM", "CM")
    ):
        return f"{required_char.upper()}{clean_val}"

    return "" if val_str.lower() in ["-", "/", "nr", "na", "not stated"] else val_str


# Execute staging aggregation, dictionary extraction, column renaming, and TNM cleaning in correct order
for df_name, df in tables_mapped.items():
    existing_cols = [c for c in tumor_stage_columns if c in df.columns]
    if existing_cols:
        formatted_df = (
            df[existing_cols].fillna("").replace("NR", "", regex=False).astype(str)
        )
        for c in existing_cols:
            formatted_df[c] = formatted_df[c].apply(
                lambda v: f"{c}: {v.strip()}" if v.strip() else ""
            )

        df["Tumor Stage Grade Combined"] = formatted_df.agg(" | ".join, axis=1)
        df["Tumor Stage Grade Combined"] = (
            df["Tumor Stage Grade Combined"]
            .str.split(r"\s*\|\s*")
            .apply(
                lambda vals: " | ".join(
                    dict.fromkeys(v.strip() for v in vals if v.strip())
                )
            )
        )
        df = df.drop(columns=existing_cols)

        # Extract Staging, TNM components, Grade, and Classification from Combined Column
        df["Tumor Stage"] = df["Tumor Stage Grade Combined"].map(
            lambda x: staging_mapping.get(x, {}).get("Tumor Stage", "")
        )
        df["TNM: Tumor"] = df["Tumor Stage Grade Combined"].map(
            lambda x: staging_mapping.get(x, {}).get("TNM: Tumor", "")
        )
        df["TNM: Node"] = df["Tumor Stage Grade Combined"].map(
            lambda x: staging_mapping.get(x, {}).get("TNM: Node", "")
        )
        df["TNM: Metastasis"] = df["Tumor Stage Grade Combined"].map(
            lambda x: staging_mapping.get(x, {}).get("TNM: Metastasis", "")
        )
        df["Grade"] = df["Tumor Stage Grade Combined"].map(
            lambda x: staging_mapping.get(x, {}).get("Grade", "")
        )
        df["Additional tumor classification"] = df["Tumor Stage Grade Combined"].map(
            lambda x: staging_mapping.get(x, {}).get(
                "Additional tumor classification", ""
            )
        )

        df = df.rename(columns=column_mapping["Staging_Grading_and_Scores"])

    # Clean and standardize TNM parameters AFTER extraction from staging_mapping
    if "TNM: Tumor" in df.columns:
        df["TNM: Tumor"] = df["TNM: Tumor"].apply(lambda x: clean_tnm_parameter(x, "T"))
    if "TNM: Node" in df.columns:
        df["TNM: Node"] = df["TNM: Node"].apply(lambda x: clean_tnm_parameter(x, "N"))
    if "TNM: Metastasis" in df.columns:
        df["TNM: Metastasis"] = df["TNM: Metastasis"].apply(
            lambda x: clean_tnm_parameter(x, "M")
        )

    tables_mapped[df_name] = df

# 6. Specimen, Sample & Tissue QC
tissue_description_cols = [
    "Cancer_Status",
    "New_Primary_Or_Recurrence_SpecimenNature",
    "Block_Type",
    "Specimen_Nature",
    "Specimen_Type",
]
vendor_tissue_size_cols = [
    "DescriptionTumour_MacroscopicTumourType_GrossDescription",
    "Divided",
    "QC_Results_Tumour_Size",
]
vendor_necrosis_cols = ["QC_Results_Necrotic_Tissue_Perc", "Tumour_Necrosis"]
vendor_tumor_cols = ["QC_Results_TumourTissue_Perc", "TumourInvasion_Extent"]
sample_type_cols = ["Sample_Format", "Sample_Type", "Specimen1_Type", "Specimen_1_Type"]

for df_name, df in tables_mapped.items():
    df["Tissue description"] = combine_columns(df, tissue_description_cols)
    df["Vendor tissue size (mm)"] = combine_columns(df, vendor_tissue_size_cols)
    df["Vendor Necrosis %"] = combine_columns(df, vendor_necrosis_cols)
    df["Vendor Tumor %"] = combine_columns(df, vendor_tumor_cols)
    df["Sample Type"] = combine_columns(df, sample_type_cols)

    src_cols_to_drop = [
        c
        for c in tissue_description_cols
        + vendor_tissue_size_cols
        + vendor_necrosis_cols
        + vendor_tumor_cols
        + sample_type_cols
        if c in df.columns
    ]
    df = df.drop(columns=src_cols_to_drop)
    tables_mapped[df_name] = df.rename(columns=column_mapping["Specimen_Sample_and_QC"])

# 7. Biomarkers, Genetics & IHC
test_results_genomics_cols = [
    "BRAF_Mutation",
    "CDKN2A_Loss",
    "Chromosome_19Q13_Loss",
    "Chromosome_1P36_Loss",
    "EGFR_Amplification",
    "EGFR_Mutation",
    "Histone_H3F3A_Mutation",
    "Histone_HIST1H3B_Mutation",
    "IDH2_Mutation",
    "KRAS_Mutation",
    "MGMT_Promoter_Methylation",
    "NGS_IDH1_Mutation",
    "NRAS_Mutation",
    "TERT_Promoter_Mutation",
]
for df_name, df in tables_mapped.items():
    df["Test results (Genomics)"] = combine_with_column_name(
        df, test_results_genomics_cols
    )
    df = df.drop(columns=[c for c in test_results_genomics_cols if c in df.columns])
    tables_mapped[df_name] = df.rename(
        columns=column_mapping["Biomarkers_Genetics_and_IHC"]
    )


# Function to standardize Estrogen and Progesterone Receptors
def standardize_receptor(val):
    if pd.isna(val) or str(val).strip() in ["", "/", "nan", "None", "NR", "ND"]:
        return ""
    val_str = str(val).strip().lower()

    # Check numeric percentages
    if val_str.isdigit():
        return "pos" if int(val_str) > 0 else "neg"

    if any(pos in val_str for pos in ["pos", "+", "positive"]):
        return "pos"
    elif any(neg in val_str for neg in ["neg", "-", "negative", "0"]):
        return "neg"

    return ""


# Function to standardize HER2 values
def standardize_her2(val):
    if pd.isna(val) or str(val).strip() in ["", "/", "nan", "None", "NR", "ND"]:
        return ""
    val_str = str(val).strip().upper()

    if val_str in ["3", "3+", "POSITIVE 3+"]:
        return "positive 3+"
    elif val_str in ["2", "2+", "2A", "POSITIVE 2+"]:
        return "positive 2+"
    elif val_str in ["1", "1+", "NEGATIVE 1+"]:
        return "negative 1+"
    elif val_str in ["0", "NEGATIVE 0"]:
        return "negative 0"
    elif "NEG" in val_str:
        return "negative"
    elif "POS" in val_str:
        return "positive 3+"

    return ""


# Apply biomarker standardization to all dataframes
for df_name, df in tables_mapped.items():
    if "Estrogen Receptor" in df.columns:
        df["Estrogen Receptor"] = df["Estrogen Receptor"].apply(standardize_receptor)
    if "Progesteron Receptor" in df.columns:
        df["Progesteron Receptor"] = df["Progesteron Receptor"].apply(
            standardize_receptor
        )
    if "HER2" in df.columns:
        df["HER2"] = df["HER2"].apply(standardize_her2)

    tables_mapped[df_name] = df

# 8. Treatments, Surgery, Relapse, Dates & Outcomes
followup_cols = [
    "AdjuvantTx1",
    "AdjuvantTx1_Agent",
    "AdjuvantTx2",
    "AdjuvantTx2_Agent",
    "AdjuvantTx3",
    "AdjuvantTx3_Agent",
    "AdjuvantTx4",
    "AdjuvantTx4_Agent",
    "Adjuvant_ChemoTx_1",
    "Adjuvant_ChemoTx_2",
    "Adjuvant_ChemoTx_3",
    "Adjuvant_ChemoTx_Regimen1",
    "Adjuvant_ChemoTx_Regimen_1",
    "Adjuvant_ChemoTx_Regimen_2",
    "Adjuvant_Tx_Regimen1",
    "Adjuvant_Tx_Regimen2",
    "Immuno_Tx",
    "Lines_Of_Therapy",
    "Other_Therapy",
    "Radiation1_Type",
    "Radiation2_Type",
    "Radiation_1_Type",
    "Radiation_2_Type",
    "Radiotherapy",
    "Rituximab",
    "SubsequentTx1",
    "SubsequentTx1_Agent",
    "SubsequentTx2",
    "SubsequentTx2_Agent",
    "SubsequentTx3",
    "SubsequentTx3_Agent",
    "SubsequentTx4",
    "SubsequentTx4_Agent",
    "SubsequentTx5",
    "SubsequentTx5_Agent",
    "Therapy_1",
    "Therapy_2",
    "Other_Relapse_1",
    "Other_Relapse_1_Sites",
    "Other_Relapse_2",
    "Other_Relapse_2_Sites",
    "Other_Relapse_3",
    "Other_Relapse_3_Sites",
    "Regional_LymphNodes_Status",
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
    "TumourResponse_AdjuvantTx1",
    "TumourResponse_AdjuvantTx2",
    "TumourResponse_AdjuvantTx3",
    "TumourResponse_AdjuvantTx4",
    "TumourResponse_NeoadjuvantTx1",
    "TumourResponse_NeoadjuvantTx2",
    "TumourResponse_NeoadjuvantTx3",
    "TumourResponse_SubsequentTx1",
    "TumourResponse_SubsequentTx2",
    "TumourResponse_SubsequentTx3",
    "TumourResponse_SubsequentTx4",
    "TumourResponse_SubsequentTx5",
    "Adjuvant_ChemoTx_1_End_Date",
    "Adjuvant_ChemoTx_1_Start_Date",
    "Adjuvant_ChemoTx_2_End_Date",
    "Adjuvant_ChemoTx_2_Start_Date",
    "Adjuvant_ChemoTx_3_End_Date",
    "Adjuvant_ChemoTx_3_Start_Date",
    "Adjuvant_ChemoTx_Regimen_1_Finish_Date",
    "Adjuvant_ChemoTx_Regimen_1_Start_Date",
    "CurrentStatus",
    "Current_Status",
    "Date_Last_Follow_Up",
    "Date_Last_Known_Alive",
    "Date_Last_Seen",
    "Date_Of_Last_Contact",
    "Date_Of_Met_Disease",
    "Date_Of_Met_Disease_1",
    "Date_Of_Met_Disease_2",
    "Date_Of_Relapse",
    "Estimated_Treatment_Dates",
    "Event_RecurrenceProgression",
    "First_Relapse_Date",
    "Follow_Up_1",
    "Follow_Up_2",
    "Follow_Up_Result",
    "Immuno_Tx_End_Date",
    "Immuno_Tx_Start_Date",
    "Last_Follow_Up",
    "Radiation_1_End_Date",
    "Radiation_1_Start_Date",
    "Radiation_2_End_Date",
    "Radiation_2_Start_Date",
    "Start_Treatment_1",
]

previous_treatment_cols = [
    "NeoadjuvantTx1",
    "NeoadjuvantTx1_Agent",
    "NeoadjuvantTx2",
    "NeoadjuvantTx2_Agent",
    "NeoadjuvantTx3",
    "NeoadjuvantTx3_Agent",
]
cause_of_death_cols = [
    "CauseOfDeath",
    "Cause_Of_Death",
    "Cause_of_Death",
    "DeathTumourRelated",
    "Event_CancerDeath",
    "Event_DeathAllCauses",
    "Event_RecurrenceCancerDeath",
]
date_of_diagnosis_cols = [
    "DateDX",
    "Date_Of_Clinical_DX",
    "Date_Of_Diagnosis",
    "Date_Of_Histological_Diagnosis",
    "Date_of_Diagnosis",
    "Month_Year_Initial_Diagnosis_PrimaryCancer",
]
collection_date_cols = [
    "Collection_Date",
    "Date_Of_Collection",
    "Date_Of_Surgery",
    "Date_of_Surgery",
    "Month_Year_Surgery_SampleCollection",
    "SurgeryDate",
    "Year_Of_Collection",
]
date_of_death_cols = ["Date_Of_Death", "Date_of_Death", "Death_Date"]

# 8. Treatments, Surgery, Relapse, Dates & Outcomes
for df_name, df in tables_mapped.items():
    # 1. Use combine_with_column_name ONLY for complex therapy/treatment fields
    df["Follow-up data"] = combine_with_column_name(df, followup_cols)
    df["Previous Treatment"] = combine_with_column_name(df, previous_treatment_cols)
    df["Cause of Death"] = combine_columns(
        df, cause_of_death_cols
    )  # Clean values without prefix

    # 2. Use combine_columns (values only) for Date fields to avoid prefixes
    df["Date of Diagnosis"] = combine_columns(df, date_of_diagnosis_cols)
    df["Collection Date"] = combine_columns(df, collection_date_cols)
    df["Date of Death"] = combine_columns(df, date_of_death_cols)

    # 3. Deduplicate date values if multiple raw date columns were joined (e.g., "2016-10-01 | 2016-10-01")
    for date_col in ["Date of Diagnosis", "Collection Date", "Date of Death"]:
        if date_col in df.columns:
            df[date_col] = (
                df[date_col]
                .str.split(r"\s*\|\s*")
                .apply(
                    lambda vals: " | ".join(
                        dict.fromkeys(v.strip() for v in vals if v.strip())
                    )
                )
            )

    # 4. Drop raw source columns
    all_dates_treatments_cols = (
        followup_cols
        + previous_treatment_cols
        + cause_of_death_cols
        + date_of_diagnosis_cols
        + collection_date_cols
        + date_of_death_cols
    )
    df = df.drop(columns=[c for c in all_dates_treatments_cols if c in df.columns])

    # 5. Apply category renames
    df = df.rename(columns=column_mapping["Treatments_and_Procedures"])
    df = df.rename(columns=column_mapping["Disease_Progression_and_Metastasis"])
    tables_mapped[df_name] = df.rename(columns=column_mapping["Dates_and_Outcomes"])

# 9. Unorganized / Administrative
comments_cols = ["Comments", "Comments_ClinicalTeam", "NHSB_Comments", "Note", "Others"]
for df_name, df in tables_mapped.items():
    df["Comments"] = combine_columns(df, comments_cols)
    df = df.drop(columns=[c for c in comments_cols if c in df.columns])
    tables_mapped[df_name] = df.rename(
        columns=column_mapping["Administrative_and_Notes"]
    )

# 10. Clean duplicate columns and nan/unnamed columns across all dataframes
for df_name, df in tables_mapped.items():
    # Drop NA / Unnamed columns
    df = df.loc[:, df.columns.notna()]
    df = df.loc[:, ~df.columns.str.startswith("Unnamed:")]

    # Resolve duplicate column names by keeping the first non-null value per row across duplicates
    if df.columns.has_duplicates:
        df = df.T.groupby(level=0).first().T

    tables_mapped[df_name] = df

# 11. Format the MSI Status column
for df_name, df in tables_mapped.items():
    if "MSI Status" in df.columns:
        df["MSI Status"] = df["MSI Status"].replace(
            {
                "normal": "MSS",
                "abnormal": "MSI-H",
            }
        )
        tables_mapped[df_name] = df

# Check if any expected columns are missing
cols_to_have = [
    "Internal ID",
    "ExpMED ID",
    "Vendor",
    "Site",
    "Vendor Order ID",
    "Vendor Sample ID",
    "Patient ID",
    "Sample Type",
    "Matched sample available",
    "Duplicate Identifier",
    "Organ_Combined",
    "Organ",
    "Organ sublocalization",
    "Species",
    "Tumor indication",
    "Vendor Disease Term",
    "Histological subtype",
    "Morphology Number",
    "Collection Date",
    "Country of Origin",
    "Gender",
    "Ethnicity",
    "Tissue description",
    "Tumor Stage",
    "TNM: Tumor",
    "TNM: Node",
    "TNM: Metastasis",
    "TNM: Nodes positive",
    "TNM: total nodes",
    "TNM: Veins",
    "TNM: Lymphatic Vessels",
    "TNM: radicality",
    "TNM Edition",
    "Grade",
    "Additional tumor classification",
    "Estrogen Receptor",
    "Progesteron Receptor",
    "HER2",
    "Tumor %",
    "Necrosis %",
    "Vendor Tumor %",
    "Vendor Necrosis %",
    "Vendor tissue size (mm)",
    "QC Status",
    "Used for TMA",
    "Number of slides",
    "Shipped to",
    "Postmortem ischemia time (min)",
    "Ischemia time (min)",
    "Tobacco History",
    "Smoking cessation date",
    "Smoking history (Descr.)",
    "Alcohol History (Descr.)",
    "Test results (Serology)",
    "Test results (Genomics)",
    "In-house sequencing",
    "Patient Age at collection",
    "MSI Status",
    "Height (cm)",
    "Weight (kg)",
    "Cause of death",
    "Date of Death",
    "Type of Surgery",
    "Previous Treatment",
    "Medications",
    "Comorbidities",
    "Follow-up data",
    "Consent type",
    "Comments",
]


cols_with_restrictions = [
    "Sample Type",
    "Organ",
    "Tumor indication",
    "Collection Date",
    "Gender",
    "Ethnicity",
    "Tissue description",
    "Tumor Stage",
    "TNM: Tumor",
    "TNM: Node",
    "TNM: Metastasis",
    "Grade",
    "Additional tumor classification",
    "Estrogen Receptor",
    "Progesteron Receptor",
    "HER2",
    "Tumor %",
    "Necrosis %",
    "Vendor Tumor %",
    "Vendor Necrosis %",
    "Postmortem ischemia time (min)",
    "Ischemia time (min)",
    "Tobacco History",
    "Smoking cessation date",
    "Patient Age at collection",
    "MSI Status",
    "Height (cm)",
    "Weight (kg)",
    "Date of Death",
    "Type of Surgery",
    "Consent type",
]

# Clean the Vendor Tumor % column
for df_name, df in tables_mapped.items():
    if "Vendor Tumor %" in df.columns:
        df["Vendor Tumor %"] = (
            df["Vendor Tumor %"].astype(str).str.replace("\n", " ", regex=False)
        )
        tables_mapped[df_name] = df

# only keep the columns that are in cols_to_have
for df_name, df in tables_mapped.items():
    df = df[[col for col in cols_to_have if col in df.columns]]
    tables_mapped[df_name] = df

# check the unique values of the cols_to_have in each dataframe
# for df_name, df in tables_mapped.items():
#     print(df.shape[1])
#     for col in cols_with_restrictions:
#         if col in df.columns:
#             unique_values = df[col].dropna().unique()
#             print(f"{df_name} - {col}: {unique_values}")
#     print("=" * 50)

# Add the Center of Origin column to all dataframes
for df_name, df in tables_mapped.items():
    if "leeds" in df_name.lower() or "ovarian" in df_name.lower():
        df["Center of Origin"] = "Leeds"
    elif "hgugm" in df_name.lower():
        df["Center of Origin"] = "HGUGM"
    elif "hannover" in df_name.lower():
        df["Center of Origin"] = "Hannover"
    elif "chu_nice" in df_name.lower():
        df["Center of Origin"] = "CHU Nice"
    elif "nhsb" in df_name.lower():
        df["Center of Origin"] = "NHSB"
    elif "iom" in df_name.lower():
        df["Center of Origin"] = "IOM"

# Combine all follow up dataframes into a single dataframe
combined_followup_df = pd.concat(
    [df for df_name, df in tables_mapped.items() if "fu" in df_name], ignore_index=True
)

# Combine all basic dataframes into a single dataframe
combined_basic_df = pd.concat(
    [df for df_name, df in tables_mapped.items() if "basic" in df_name],
    ignore_index=True,
)

# Combine all supplementary dataframes into a single dataframe
combined_supplementary_df = pd.concat(
    [df for df_name, df in tables_mapped.items() if "supplementary" in df_name],
    ignore_index=True,
)

# Arrange the columns in the follow up dataset
cols_order_followup = [
    "Internal ID",
    "Vendor Order ID",
    "Vendor Sample ID",
    "Patient ID",
    "Center of Origin",
    "Gender",
    "Ethnicity",
    "Tobacco History",
    "Smoking history (Descr.)",
    "Comorbidities",
    "Sample Type",
    "Tissue description",
    "Vendor tissue size (mm)",
    "Organ_Combined",
    "Organ",
    "Organ sublocalization",
    "Tumor indication",
    "Vendor Disease Term",
    "Collection Date",
    "Tumor Stage",
    "TNM Edition",
    "TNM: Tumor",
    "TNM: Node",
    "TNM: Metastasis",
    "TNM: Nodes positive",
    "TNM: total nodes",
    "Grade",
    "Additional tumor classification",
    "Vendor Tumor %",
    "Vendor Necrosis %",
    "Estrogen Receptor",
    "Progesteron Receptor",
    "HER2",
    "Test results (Genomics)",
    "Previous Treatment",
    "Medications",
    "Follow-up data",
    "Date of Death",
]
combined_followup_df = combined_followup_df[cols_order_followup]

# Arrange the columns in the supplementary dataset
cols_order_supplementary = [
    "Internal ID",
    "Vendor Order ID",
    "Vendor Sample ID",
    "Patient ID",
    "Center of Origin",
    "Ethnicity",
    "Sample Type",
    "Tissue description",
    "Vendor tissue size (mm)",
    "Organ_Combined",
    "Organ",
    "Organ sublocalization",
    "Collection Date",
    "Tumor Stage",
    "TNM: Tumor",
    "TNM: Node",
    "TNM: Metastasis",
    "Grade",
    "Additional tumor classification",
    "Vendor Tumor %",
    "Vendor Necrosis %",
    "MSI Status",
    "Test results (Genomics)",
    "Previous Treatment",
    "Follow-up data",
    "Date of Death",
]
combined_supplementary_df = combined_supplementary_df[cols_order_supplementary]

# Arrange the columns in the basic dataset
cols_order_basic = cols_order_basic = [
    "Internal ID",
    "Vendor Order ID",
    "Vendor Sample ID",
    "Patient ID",
    "Center of Origin",
    "Gender",
    "Ethnicity",
    "Tobacco History",
    "Smoking history (Descr.)",
    "Sample Type",
    "Tissue description",
    "Vendor tissue size (mm)",
    "Organ_Combined",
    "Organ",
    "Organ sublocalization",
    "Tumor indication",
    "Vendor Disease Term",
    "Collection Date",
    "Tumor Stage",
    "TNM Edition",
    "TNM: Tumor",
    "TNM: Node",
    "TNM: Metastasis",
    "Grade",
    "Additional tumor classification",
    "Vendor Tumor %",
    "Vendor Necrosis %",
    "Test results (Genomics)",
    "Previous Treatment",
    "Follow-up data",
    "Date of Death",
]
combined_basic_df = combined_basic_df[cols_order_basic]


# Final formatting steps
combined_followup_df["Organ sublocalization"] = combined_followup_df[
    "Organ sublocalization"
].str.replace(r"Hgugm", "", regex=True)

# Check which column contains values with a "|"
for col in combined_followup_df.columns:
    if "|" in combined_followup_df[col].astype(str).unique():
        print(f"{col}: {combined_followup_df[col].astype(str).unique()}")

# Fill the NA Gender values from Ovarian 500 with "Female"
combined_followup_df.loc[
    combined_followup_df["Internal ID"].isin([8330, 8331, 8332]), "Gender"
] = "Female"

dfs = [combined_followup_df, combined_basic_df, combined_supplementary_df]

for i in range(len(dfs)):
    df = dfs[i]

    # Convert empty strings and whitespace-only strings to NaN
    df.replace(r"^\s*$", np.nan, regex=True, inplace=True)

    # Drop columns that are completely empty (all NaN)
    df.dropna(axis=1, how="all", inplace=True)

    if "Tissue description" in df.columns:
        df.loc[df["Tissue description"] == "FFPE", "Tissue description"] = ""
        df.loc[df["Tissue description"] == "Primary", "Tissue description"] = (
            "Primary tumor"
        )

    if "Sample Type" in df.columns:
        df["Sample Type"] = "FFPE Tissue"

    # Format specific Vendor Disease Term values and corresponding Additional tumor classification
    ids_adenocarcinoma = [
        8098,
        8099,
        8104,
        8108,
        8112,
        8113,
        8114,
        8116,
        8144,
        8145,
        8149,
        8127,
        8168,
    ]
    ids_pd_adenocarcinoma = [8106, 8115, 8125, 8126]
    df.loc[df["Internal ID"].isin(ids_adenocarcinoma), "Vendor Disease Term"] = (
        "Adenocarcinoma"
    )
    df.loc[df["Internal ID"].isin(ids_pd_adenocarcinoma), "Vendor Disease Term"] = (
        "Poorly Differentiated Adenocarcinoma"
    )
    df.loc[df["Internal ID"] == 8103, "Vendor Disease Term"] = (
        "Moderate to Poorly Differentiated Adenocarcinoma"
    )
    df.loc[df["Internal ID"] == 8107, "Vendor Disease Term"] = (
        "Diffuse Adenocarcinoma with Features of Hormonal Manipulation"
    )
    df.loc[df["Internal ID"] == 8109, "Vendor Disease Term"] = (
        "Focal Adenocarcinoma with Hormonal Effects"
    )
    df.loc[df["Internal ID"] == 8110, "Vendor Disease Term"] = ""

    df.loc[df["Internal ID"] == 8098, "Additional tumor classification"] = (
        "Gleason 5+4=9"
    )
    df.loc[df["Internal ID"] == 8099, "Additional tumor classification"] = (
        "Gleason 4+5=9"
    )
    df.loc[df["Internal ID"] == 8103, "Additional tumor classification"] = (
        "Gleason 3+5=8"
    )
    df.loc[df["Internal ID"] == 8104, "Additional tumor classification"] = (
        "Gleason 4+5=9"
    )
    df.loc[df["Internal ID"] == 8106, "Additional tumor classification"] = (
        "Gleason 5+4=9"
    )
    df.loc[df["Internal ID"] == 8108, "Additional tumor classification"] = (
        "Gleason 4+3=7"
    )
    df.loc[df["Internal ID"] == 8109, "Additional tumor classification"] = (
        "Gleason 4+5=9"
    )
    df.loc[df["Internal ID"] == 8112, "Additional tumor classification"] = (
        "Gleason 4+4=8"
    )
    df.loc[df["Internal ID"] == 8113, "Additional tumor classification"] = (
        "Gleason 5+5=10"
    )
    df.loc[df["Internal ID"] == 8114, "Additional tumor classification"] = (
        "Gleason 5+4=9"
    )
    df.loc[df["Internal ID"] == 8116, "Additional tumor classification"] = (
        "Gleason 4+5=9"
    )
    df.loc[df["Internal ID"] == 8144, "Additional tumor classification"] = (
        "Gleason 4+5=9"
    )
    df.loc[df["Internal ID"] == 8145, "Additional tumor classification"] = (
        "Gleason 3+4=7"
    )
    df.loc[df["Internal ID"] == 8149, "Additional tumor classification"] = (
        "Gleason 5+5=10"
    )
    df.loc[df["Internal ID"] == 8125, "Additional tumor classification"] = (
        "Gleason 4+4=8"
    )
    df.loc[df["Internal ID"] == 8126, "Additional tumor classification"] = (
        "Gleason 5+3=8"
    )
    df.loc[df["Internal ID"] == 8127, "Additional tumor classification"] = (
        "Gleason 3+4=7"
    )
    df.loc[df["Internal ID"] == 8168, "Additional tumor classification"] = (
        "Gleason 5+5=10"
    )

    # Format the tissue description for Internal ID 8110
    df.loc[df["Internal ID"] == 8110, "Tissue description"] = "Metastasis"

    # Convert empty strings and whitespace-only strings to NaN
    df.replace(r"^\s*$", np.nan, regex=True, inplace=True)

    # Drop columns that are completely empty (all NaN)
    df.dropna(axis=1, how="all", inplace=True)

    # combine the follow up and supplementary data
    df1 = combined_followup_df.set_index("Internal ID")
    df2 = combined_basic_df.set_index("Internal ID")
    follow_up_basic = pd.concat([df1, df2], axis=0, ignore_index=False)
    df3 = combined_supplementary_df.set_index("Internal ID")
    merged_df = follow_up_basic.combine_first(df3).reset_index()

    # Replace the NR in the Gender column with blanks
    merged_df.loc[merged_df["Gender"] == "NR"] = None

    # Format the ethnicity column
    merged_df["Ethnicity"] = merged_df["Ethnicity"].replace(
        "CAUCASIAN", "Caucasian/White"
    )
    merged_df["Ethnicity"] = merged_df["Ethnicity"].replace(
        "Caucasian", "Caucasian/White"
    )
    for word in ["NR", "Not specified", "ND", pd.NA]:
        merged_df["Ethnicity"] = merged_df["Ethnicity"].replace(word, None)
    merged_df["Ethnicity"] = merged_df["Ethnicity"].replace(
        "White British", "Caucasian/White"
    )
    merged_df["Ethnicity"] = merged_df["Ethnicity"].replace(
        "Any other White Background", "Caucasian/White"
    )
    merged_df["Ethnicity"] = merged_df["Ethnicity"].replace("White", "Caucasian/White")

    # Format the Tissue description column
    merged_df["Tissue description"] = merged_df["Tissue description"].replace(
        "Metastatic", "Metastasis"
    )
    merged_df["Tissue description"] = merged_df["Tissue description"].fillna("Tumor")

    # Complete the Organ column
    for idx, row in merged_df.iterrows():
        tissue_desc = row["Tissue description"]
        sublocalization = row["Organ sublocalization"]
        indication = row["Tumor indication"]
        disease_term = row["Vendor Disease Term"]
        internal_ID = row["Internal ID"]
        organ = row["Organ"]

        if sublocalization in [
            "Tongue",
            "Larynx",
            "Laryngopharynx",
            "Brain",
            "Stomach",
            "Gallbladder",
            "Liver",
            "Prostate",
            "Parathyroid",
            "Skeletal Muscle",
        ]:
            merged_df.loc[idx, "Organ"] = sublocalization
        elif (
            tissue_desc in ["Colon Pri. And Met.", "Colon Pri."]
            or indication == "Colorectal Cancer (CRC)"
        ):
            merged_df.loc[idx, "Organ"] = "Colon"
            merged_df.loc[idx, "Tumor indication"] = "Colorectal Cancer (CRC)"
        elif (
            tissue_desc == "Gastric Pri. and Nodal Met."
            or indication == "Gastric Cancer"
            or internal_ID in [8396, 8409, 8410]
        ):
            merged_df.loc[idx, "Organ"] = "Stomach"
            merged_df.loc[idx, "Tumor indication"] = "Gastric Cancer"
        elif disease_term == "Metachronous Colon Adenocarcinoma with Liver Metastasis":
            merged_df.loc[idx, "Organ"] = "Colon"
        elif (
            indication
            in ["Pancreatic ductal adenocarcinoma (PDAC)", "Pancreatic cancer (other)"]
            or tissue_desc == "Pancreas Pri."
        ):
            merged_df.loc[idx, "Organ"] = "Pancreas"
            merged_df.loc[idx, "Tumor indication"] = (
                "Pancreatic ductal adenocarcinoma (PDAC)"
            )
        elif tissue_desc == "Oesophagus Pri.":
            merged_df.loc[idx, "Organ"] = "Esophagus"
            merged_df.loc[idx, "Tumor indication"] = (
                "Esophageal Squamous Cell Carcinoma (ESCC)"
            )
        elif indication == "Hepatocellular Carcinoma (HCC)":
            merged_df.loc[idx, "Organ"] = "Liver"
        elif (
            tissue_desc == "Lymph Node Pri."
            or "Node" in tissue_desc
            or disease_term in ["Fl", "Dlbcl"]
        ):
            merged_df.loc[idx, "Organ"] = "Lymph node"
            merged_df.loc[idx, "Tumor indication"] = "Lymphoma"
        elif tissue_desc in ["Small Bowel", "Small Intestine"]:
            merged_df.loc[idx, "Organ"] = "Small intestine"
        elif "Caecum" in tissue_desc:
            merged_df.loc[idx, "Organ"] = "Colon"
        elif tissue_desc == "Small Bowel And Colon":
            merged_df.loc[idx, "Organ"] = "Multiorgan"
        elif "Tonsillar Mass" in tissue_desc or "Tonsil" in tissue_desc:
            merged_df.loc[idx, "Organ"] = "Tonsil"
        elif indication == "Breast Cancer":
            merged_df.loc[idx, "Organ"] = "Breast"
        elif indication == "Prostate Cancer":
            merged_df.loc[idx, "Organ"] = "Prostate"
        elif indication == "Lung Cancer (NSCLC)":
            merged_df.loc[idx, "Organ"] = "Lung"
            merged_df.loc[idx, "Tumor indication"] = "Lung Cancer (NSCLC)"
        elif indication == "Bladder Cancer/Urothelial Carcinoma":
            merged_df.loc[idx, "Organ"] = "Bladder"
        elif indication == "Ovarian Cancer":
            merged_df.loc[idx, "Organ"] = "Ovary"
        elif internal_ID in [
            8109,
            8111,
            8112,
            8113,
            8114,
            8115,
            8116,
            8125,
            8126,
            8127,
            8168,
        ]:
            merged_df.loc[idx, "Organ"] = "Prostate"
            merged_df.loc[idx, "Tumor indication"] = "Prostate Cancer"
        elif internal_ID == 8339:
            merged_df.loc[idx, "Organ"] = "Lung"
            merged_df.loc[idx, "Tumor indication"] = "Lung Cancer (NSCLC)"
        elif disease_term == "Clear Cell Carcinoma":
            merged_df.loc[idx, "Tumor indication"] = "Renal Cell Carcinoma"
        elif disease_term == "Parathyroid Carcinoma":
            merged_df.loc[idx, "Tumor indication"] = "Thyroid Cancer"

# Apply to date columns
date_columns = ["Collection Date", "Date of Death"]
for date in date_columns:
    merged_df[date] = process_dates(merged_df[date])


# Make final replacements to specific columns
merged_df["Tissue description"] = merged_df["Tissue description"].str.replace(
    r".*Pri\..*", "Primary tumor", regex=True
)
merged_df.loc[
    ~merged_df["Tissue description"].isin(
        ["Primary tumor", "Metastasis", "Tumor", "Normal"]
    ),
    "Tissue description",
] = "Tumor"
merged_df.drop(columns="Organ_Combined", inplace=True)
merged_df["Tumor indication"] = merged_df["Tumor indication"].str.replace("N/A", "")
merged_df["Vendor Disease Term"] = merged_df["Vendor Disease Term"].str.replace(
    "NR", ""
)
merged_df["Tumor Stage"] = merged_df["Tumor Stage"].str.replace("!A", "I A")
patterns = [
    "T2N2cM0",
    "cT3a N1 M1",
    "pT3 Nx Mx",
    "cT3b N1 M1",
    "cT4 Nx Mx",
    "cT3 Nx Mx",
    "cTx N1 Mx",
]
merged_df["Tumor Stage"] = merged_df["Tumor Stage"].str.replace(
    "|".join(map(re.escape, patterns)), "", regex=True
)
merged_df["Tumor Stage"] = merged_df["Tumor Stage"].str.replace("2 b", "II B")
merged_df["Tumor Stage"] = merged_df["Tumor Stage"].str.replace("1 a", "I A")
merged_df["Tumor Stage"] = merged_df["Tumor Stage"].str.replace("1 c", "I C")
merged_df["Tumor Stage"] = merged_df["Tumor Stage"].str.replace(" (at least)", "")
stage_cols = ["TNM: Node", "TNM: Tumor", "TNM: Metastasis"]
for col in stage_cols:
    # Add closing parenthesis where already opened
    merged_df[col] = merged_df[col].str.replace(r"\((m|mi|sN)$", r"(\1)", regex=True)

    # Convert trailing m without parentheses
    merged_df[col] = merged_df[col].str.replace(r"(?<!\()m$", "(m)", regex=True)

    # Convert trailing mi without parentheses
    merged_df[col] = merged_df[col].str.replace(r"(?<!\()mi$", "(mi)", regex=True)

    # Convert trailing sN without parentheses
    merged_df[col] = merged_df[col].str.replace(r"(?<!\()sN$", "(sN)", regex=True)
cols_to_replace_NA = ["TNM: Nodes positive", "TNM: total nodes", "Vendor Tumor %"]
for col in cols_to_replace_NA:
    merged_df[col] = merged_df[col].str.replace(r"N.A.*", "", regex=True)
patterns = [
    "L1 - Lymphatic invasion V1 - Microscopic venous invasion PN1 - Perineural invasion",
    "V1 - Microscopic venous invasion",
    "Extraprostatic extension: present",
]
merged_df["Vendor Tumor %"] = merged_df["Vendor Tumor %"].str.replace(
    "|".join(map(re.escape, patterns)), "", regex=True
)
merged_df["Test results (Genomics)"] = merged_df["Test results (Genomics)"].str.replace(
    "EGFR_Mutation: - | KRAS_Mutation: -", "", regex=True
)
