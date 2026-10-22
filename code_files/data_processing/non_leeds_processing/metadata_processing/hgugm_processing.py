# -----------------------------------------------------------------------------
# Setup & Imports
# -----------------------------------------------------------------------------

# Import the required libraries
import sys
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# Define the current directory and project root
current_dir = Path(__file__).resolve()
project_root = next(
    (p for p in Path(".").resolve().parents if (p / "pyproject.toml").exists()),
    Path(".").resolve(),
)

if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from code_files.utilities.io.database import create_gcloud_engine
from code_files.utilities.processing.cleaners import title_case

pd.set_option("display.max_columns", None)
pd.set_option("display.max_rows", None)

env_path = project_root / "code_files" / "utilities" / "io" / ".env"
engine = create_gcloud_engine(path=env_path)

BASE_DIR = Path("C:/Users/medaamm/Box/Data file (AA)/HGUGM/Source Files")

# -----------------------------------------------------------------------------
# Sheet-Specific Rename Dictionaries (Run FIRST to avoid collisions)
# -----------------------------------------------------------------------------
SHEET_SPECIFIC_MAPS = {
    "Ovarian_2014": {
        "Start Date.3": "Adjuvant_ChemoTx_3_Start_Date",
        "Finish Date.4": "Adjuvant_ChemoTx_3_End_Date",
        "Start Date.4": "Adjuvant_ChemoTx_4_Start_Date",
        "Finish Date.5": "Adjuvant_ChemoTx_4_End_Date",
        "Start Date.5": "Immuno_Tx_Start_Date",
        "Finish Date.6": "Immuno_Tx_End_Date",
        "Start Date.6": "Metastatic_Chemo_Targeted_Tx_1_Start_Date",
        "Finish Date.7": "Metastatic_Chemo_Targeted_Tx_1_End_Date",
        "Start Date.7": "Metastatic_Chemo_Targeted_Tx_2_Start_Date",
        "Finish Date.8": "Metastatic_Chemo_Targeted_Tx_2_End_Date",
        "Start Date.8": "Metastatic_Chemo_Targeted_Tx_3_Start_Date",
        "Finish Date.9": "Metastatic_Chemo_Targeted_Tx_3_End_Date",
    },
    "HNSCC_2014": {
        "Start Date.3": "Metastatic_Chemo_Targeted_Tx_1_Start_Date",
        "Finish Date.4": "Metastatic_Chemo_Targeted_Tx_1_End_Date",
        "Start Date.4": "Metastatic_Chemo_Targeted_Tx_2_Start_Date",
        "Finish Date.5": "Metastatic_Chemo_Targeted_Tx_2_End_Date",
    },
    "HCC_FollowUp": {
        "Start Date.3": "Metastatic_Chemo_Targeted_Tx_1_Start_Date",
        "Finish Date.4": "Metastatic_Chemo_Targeted_Tx_1_End_Date",
    },
    "Prostate_Cancer": {
        "Start Date.3": "Metastatic_Radio_Targeted_Tx_1_Start_Date",
        "Finish Date.4": "Metastatic_Radio_Targeted_Tx_1_End_Date",
        "Start Date.4": "Metastatic_Chemo_Targeted_Tx_1_Start_Date",
        "Finish Date.5": "Metastatic_Chemo_Targeted_Tx_1_End_Date",
        "Start Date.5": "Metastatic_Chemo_Targeted_Tx_2_Start_Date",
        "Finish Date.6": "Metastatic_Chemo_Targeted_Tx_2_End_Date",
        "Start Date.6": "Immuno_Met_Tx_Start_Date",
        "Finish Date.7": "Immuno_Met_Tx_End_Date",
        "Start Date.7": "Other_Therapy_Start_Date",
        "Finish Date.8": "Other_Therapy_End_Date",
    },
    "Renal_Carcinoma": {
        "Start Date.3": "Metastatic_Radio_Targeted_Tx_1_Start_Date",
        "Finish Date.4": "Metastatic_Radio_Targeted_Tx_1_End_Date",
        "Start Date.4": "Metastatic_Chemo_Targeted_Tx_1_Start_Date",
        "Finish Date.5": "Metastatic_Chemo_Targeted_Tx_1_End_Date",
        "Start Date.5": "Metastatic_Chemo_Targeted_Tx_2_Start_Date",
        "Finish Date.6": "Metastatic_Chemo_Targeted_Tx_2_End_Date",
        "Start Date.6": "Immuno_Met_Tx_Start_Date",
        "Finish Date.7": "Immuno_Met_Tx_End_Date",
    },
    "Gastric_Cancer": {
        "Start Date.3": "Adjuvant_ChemoTx_3_Start_Date",
        "Finish Date.4": "Adjuvant_ChemoTx_3_End_Date",
        "Start Date.4": "Immuno_Tx_Start_Date",
        "Finish Date.5": "Immuno_Tx_End_Date",
        "Start Date.5": "Metastatic_Chemo_Targeted_Tx_1_Start_Date",
        "Finish Date.6": "Metastatic_Chemo_Targeted_Tx_1_End_Date",
        "Start Date.6": "Metastatic_Chemo_Targeted_Tx_2_Start_Date",
        "Finish Date.7": "Metastatic_Chemo_Targeted_Tx_2_End_Date",
    },
    "Bladder_Cancer": {
        "Start Date.3": "Immuno_Tx_Start_Date",
        "Finish Date.4": "Immuno_Tx_End_Date",
    },
}

# -----------------------------------------------------------------------------
# Master Global Mapping Dictionary
# -----------------------------------------------------------------------------
GLOBAL_COLUMN_MAP = {
    # Identifiers
    "Unique ID (Series)": "Unique_ID_Series",
    "Unique ID#": "Sample_ID",
    "#Id": "Sample_ID",
    "# Id": "Sample_ID",
    "Sample Id": "Sample_ID",
    "Sample Id#": "Sample_ID",
    "SAMPLE ID#": "Sample_ID",
    "Atcg #": "Sample_ID",
    "ATCG #": "Sample_ID",
    "Patient Id": "Patient_ID",
    "Patient ID": "Patient_ID",
    "Spot #": "Spot_Number",
    # Demographics & Clinical Diagnosis
    "sex": "Gender",
    "Sex": "Gender",
    "Age range": "Age_Range",
    "Age At Diagnosis (Years)": "Age",
    "Age at Diagnosis (years)": "Age",
    "race": "Ethnicity",
    "Race": "Ethnicity",
    "Clinical diagnosis": "Clinical_Diagnosis",
    "Clinical Diagnose": "Clinical_Diagnosis",
    "Diagnosis": "Pathological_Diagnosis",
    "Pathological Diagnosis": "Pathological_Diagnosis",
    "Pathological Diagnose": "Pathological_Diagnosis",
    "Histological Diagnosis": "Histological_Diagnosis",
    "Histological diagnosis": "Histological_Diagnosis",
    "smoker": "Smoker",
    "Smoker": "Smoker",
    "Tumor Grade": "Grade",
    "Grade": "Grade",
    "GRADE": "Grade",
    "Nuclear Grade": "Nuclear_Grade",
    "NUCLEAR GRADE": "Nuclear_Grade",
    " Stage ": "Stage",
    "Stage": "Stage",
    "Stage At Dx": "Stage",
    "Stage At Dx ": "Stage",
    "Stage at DX ": "Stage",
    "Tnm": "TNM",
    "TNM": "TNM",
    "Gleason": "Gleason",
    "Score": "Gleason_Score",
    # Blocks / Mets
    "NODAL METS": "Nodal_Mets",
    "Nodal Mets": "Nodal_Mets",
    "Block A": "Block_A",
    "BLOCK A": "Block_A",
    "Block B": "Block_B",
    "BLOCK B": "Block_B",
    "Block C": "Block_C",
    "BLOCK C": "Block_C",
    "Block D": "Block_D",
    "BLOCK D": "Block_D",
    "Block E": "Block_E",
    "BLOCK E": "Block_E",
    "Block F": "Block_F",
    "BLOCK F": "Block_F",
    "Met Block 1": "Met_Block_1",
    "Met Block 2": "Met_Block_2",
    "Met Block 3": "Met_Block_3",
    # Surgery & Specimen
    "surgical specimen": "Surgical_Specimen",
    "Surgical Specimen": "Surgical_Specimen",
    "Surgery & Site": "Surgical_Specimen",
    "Surgical Date": "Date_Of_Surgery",
    "Date Of Surgery                                 (Mm/Yy)": "Date_Of_Surgery",
    "Date of Surgery                                 (mm/yy)": "Date_Of_Surgery",
    "Date Of Surgery (Mm/Yy)": "Date_Of_Surgery",
    "Date of Surgery (mm/yy)": "Date_Of_Surgery",
    "Specimen 1 Collection Date  (Mm/Yy)": "Specimen_Collection_Date",
    "Specimen 1 Collection Date  (mm/yy)": "Specimen_Collection_Date",
    "Specimen 1 Collection Date (Mm/Yy)": "Specimen_Collection_Date",
    "Specimen 1 Collection Date (mm/yy)": "Specimen_Collection_Date",
    "Specimen 1 Type": "Specimen_Type",
    "Site Of Diagnosis": "Site_Of_Diagnosis",
    "Site of Diagnosis": "Site_Of_Diagnosis",
    # History & Status
    "previous treatment": "Previous_Treatment",
    "Previous Treatment": "Previous_Treatment",
    "clinical response": "Previous_Treatment_ClinicalResponse",
    "Clinical Response": "Clinical_Response",
    "survival": "Survival",
    "Survival": "Survival",
    "Census": "Current_Status",
    "Exitus": "Current_Status",
    "Current Status": "Current_Status",
    "Date Last Seen": "Date_Last_Seen",
    "Death Date": "Death_Date",
    "Cause Of Death": "Cause_Of_Death",
    "Cause of Death": "Cause_Of_Death",
    "Concurrent Medical Conditions": "Concurrent_Medical_Conditions",
    "Date Of Dx (Mm/Yy)": "Date_Of_DX",
    "Date of DX (mm/yy)": "Date_Of_DX",
    "Date": "Date_Of_DX",
    # Relapses & Metastases
    "Date Of Met Disease": "Date_Of_Met_Disease_1",
    "Date of Met Disease": "Date_Of_Met_Disease_1",
    "Met Disease": "Met_Disease_1",
    "Other Relapse 1": "Other_Relapse_1",
    "Sites": "Other_Relapse_1_Sites",
    "Other Relapse 2": "Other_Relapse_2",
    "Sites2": "Other_Relapse_2_Sites",
    "Other Relapse 3": "Other_Relapse_3",
    "Sites3": "Other_Relapse_3_Sites",
    "Other Relapse 4": "Other_Relapse_4",
    "Sites4": "Other_Relapse_4_Sites",
    # Therapies
    "Chemotherapy": "Any_Chemotherapy",
    "Quimiothetapy": "Any_Chemotherapy",
    "1st line": "1st_Line_Chemotherapy",
    "1ª Line": "1st_Line_Chemotherapy",
    "Treatment": "1st_Line_Chemotherapy_Agent",
    "2º line": "2nd_Line_Chemotherapy",
    "2º Line": "2nd_Line_Chemotherapy",
    "Treatment.1": "2nd_Line_Chemotherapy_Agent",
    "3º line": "3rd_Line_Chemotherapy",
    "3º Line": "3rd_Line_Chemotherapy",
    "Treatment.2": "3rd_Line_Chemotherapy_Agent",
    "Neoadjuvant": "Any_Neoadjuvant_Chemotherapy",
    "Qt Neoadyuv": "Neoadjuvant_Chemotherapy",
    "QT neoadyuv": "Any_Neoadjuvant_Chemotherapy",
    "Treatment.3": "Neoadjuvant_Therapy_Agent",
    "Palliative QT": "Any_Palliative_Chemotherapy",
    "Qt Palliative": "Palliative_Chemotherapy",
    "Treatment.4": "Palliative_Chemotherapy_Agent",
    "sensitizing QT": "Any_Sensitizing_Chemotherapy",
    "Qt Sensibilizante": "Sensitizing_Chemotherapy",
    "Treatment.5": "Sensitizing_Chemotherapy_Agent",
    "DRUG": "Drug",
    "Adjuvant Chemotx Regimen 1": "Adjuvant_ChemoTx_1",
    "Adjuvant ChemoTx Regimen 1": "Adjuvant_ChemoTx_1",
    "Adjuvant Chemotx Regimen 2": "Adjuvant_ChemoTx_2",
    "Adjuvant ChemoTx Regimen 2": "Adjuvant_ChemoTx_2",
    "Adjuvant Chemotx Regimen 3": "Adjuvant_ChemoTx_3",
    "Adjuvant ChemoTx Regimen 3": "Adjuvant_ChemoTx_3",
    "Adjuvant Chemotx Regimen 4": "Adjuvant_ChemoTx_4",
    "Start Date": "Radiation_1_Start_Date",
    "Finish Date": "Radiation_1_End_Date",
    "Start Date.1": "Adjuvant_ChemoTx_1_Start_Date",
    "Finish Date.2": "Adjuvant_ChemoTx_1_End_Date",
    "Radiation 1: Type:": "Radiation_1_Type",
    "Radiation 1: Type": "Radiation_1_Type",
    "Radiation 2: Type": "Radiation_2_Type",
    "Start Date 2": "Radiation_2_Start_Date",
    "Finish Date.1": "Radiation_2_End_Date",
    "Immuno Tx": "Immuno_Tx",
    "Start Date.2": "Immuno_Tx_Start_Date",
    "Finish Date.3": "Immuno_Tx_End_Date",
    "Metastatic Chemo/Targeted                      Tx Regimen 1": (
        "Metastatic_Chemo_Targeted_Tx_1"
    ),
    "Metastatic Chemo/Targeted                      Tx Regimen 2": (
        "Metastatic_Chemo_Targeted_Tx_2"
    ),
    "Metastatic Chemo/Targeted                      Tx Regimen 3": (
        "Metastatic_Chemo_Targeted_Tx_3"
    ),
    "Metastatic Radio/Targeted                      Tx Regimen 1": (
        "Metastatic_Radio_Targeted_Tx_1"
    ),
    "Metastatic Radiation/Targeted                      Tx Regimen 1": (
        "Metastatic_Radio_Targeted_Tx_1"
    ),
    "Immuno Met Tx": "Immuno_Met_Tx",
    "Other Therapy": "Other_Therapy",
}

VALUE_MAPPINGS = {
    "Pathological_Diagnosis": {
        "1": "Squamous Cell Carcinoma",
        "2": "Adenocarcinoma",
        "3": "Bronchioalveolar Carcinoma",
        "4": "Small Cell Lung Cancer",
        "5": "Large Cell Undifferentiated Carcinoma",
        "6": "Adenosquamous Carcinoma",
        "7": "Mixed (Small Cell/Non-Small Cell) Carcinoma",
        "8": "Other Mixed Carcinoma",
        "9": "Other (Not Specified)",
    },
    "Gender": {"1": "Male", "2": "Female"},
    "Ethnicity": {"1": "White", "2": "Black", "3": "Asian"},
    "Clinical_Diagnosis": {
        "1": "Lung Neoplasia",
        "2": "Infectious",
        "3": "Other",
    },
    "Smoker": {"1": "Yes", "2": "No"},
    "Previous_Treatment": {
        "1": "No treatment prior to surgery",
        "2": "Neoadjuvant Chemotherapy",
        "3": "Neoadjuvant Radiotherapy",
    },
    "Clinical_Response": {
        "1": "No Response",
        "2": "Partial Response",
        "3": "Complete Response",
    },
    "Previous_Treatment_ClinicalResponse": {
        "1": "No Response",
        "2": "Partial Response",
        "3": "Complete Response",
    },
    "Surgical_Specimen": {
        "1": "Lobectomy",
        "2": "Bilobectomy",
        "3": "Pneumonectomy",
        "4": "Other",
    },
    "Census": {
        "1": "Alive",
        "2": "Death (Unknown Cause)",
        "3": "Tumoural Death",
        "4": "Non-Tumoural Death",
    },
    "Current_Status": {
        "1": "Alive or Non-Tumoural Death",
        "2": "Tumoural Death",
        "3": "Unknown Cause",
    },
    "Any_Chemotherapy": {"1": "No", "2": "Yes", "3": "Unknown"},
    "1st_Line_Chemotherapy": {"1": "No", "2": "Yes", "3": "Unknown"},
    "2nd_Line_Chemotherapy": {"1": "No", "2": "Yes", "3": "Unknown"},
    "3rd_Line_Chemotherapy": {"1": "No", "2": "Yes", "3": "Unknown"},
    "Any_Neoadjuvant_Chemotherapy": {"1": "No", "2": "Yes", "3": "Unknown"},
    "Any_Palliative_Chemotherapy": {"1": "No", "2": "Yes", "3": "Unknown"},
    "Any_Sensitizing_Chemotherapy": {"1": "No", "2": "Yes", "3": "Unknown"},
}


# -----------------------------------------------------------------------------
# Clean & Read Generic Helper Function
# -----------------------------------------------------------------------------
def load_and_clean_excel(
    file_path: Path,
    sheet_name: str,
    output_key: str,
    header: int,
    drop_dx: bool = True,
    apply_title_case: bool = False,
) -> pd.DataFrame:
    """Generic reader pipeline to parse, rename, replace mapped values, and clean rows."""
    df = pd.read_excel(
        file_path, sheet_name=sheet_name, header=header, dtype=object
    ).dropna(how="all")

    if apply_title_case:
        df = title_case(df)

    # 1. Sheet-specific renames FIRST to prevent duplicate suffix collisions
    if output_key in SHEET_SPECIFIC_MAPS:
        df = df.rename(columns=SHEET_SPECIFIC_MAPS[output_key])

    # 2. General global column mapping SECOND
    df = df.rename(columns=GLOBAL_COLUMN_MAP)

    # 3. Clean column names
    df.columns = df.columns.astype(str).str.strip()
    df = df.loc[:, ~df.columns.str.contains("^Unnamed", case=False, na=False)]
    df = df.loc[:, df.columns.notnull() & (df.columns != "")]

    # 4. Perform value mappings
    for col, mapping in VALUE_MAPPINGS.items():
        if col in df.columns:
            str_map = {str(k): v for k, v in mapping.items()}
            df[col] = df[col].astype(str).replace(str_map)

    # 5. Replace invalid/missing strings with NaN
    clean_cols = [
        c
        for c in [
            "Gender",
            "Ethnicity",
            "Stage",
            "Grade",
            "T",
            "N",
            "M",
            "Specimen_Type",
        ]
        if c in df.columns
    ]
    for c in clean_cols:
        df[c] = df[c].replace(["\u3000", "NR", "nan", np.nan], pd.NA)

    if drop_dx and "Date_Of_DX" in df.columns:
        df = df.drop(columns=["Date_Of_DX"])

    # 6. Safety check: keep unique column names only
    if df.columns.duplicated().any():
        df = df.loc[:, ~df.columns.duplicated()]

    return df


# -----------------------------------------------------------------------------
# 1. Baseline: NSCLC Prognosis (Combine 3 Sheets)
# -----------------------------------------------------------------------------
prognosis_file = (
    BASE_DIR
    / "69571118 NSCLC prognosis array with treatment data HGUGM-4 corrected.xlsx"
)
nsclc_sheets = [
    ("Data sheet (1167)", 15),
    ("Data sheet(1248)", 13),
    ("Data sheet1249", 13),
]

hgugm_nsclc_prognosis_combined = pd.concat(
    [
        load_and_clean_excel(
            prognosis_file,
            sheet,
            "NSCLC_Prognosis",
            hdr,
            drop_dx=False,
            apply_title_case=False,
        )
        for sheet, hdr in nsclc_sheets
    ],
    ignore_index=True,
)

# Convert numeric metrics
for col in ["Reference", "Age", "Year_Diagnosis", "Month_Diagnosis"]:
    if col in hgugm_nsclc_prognosis_combined.columns:
        hgugm_nsclc_prognosis_combined[col] = pd.to_numeric(
            hgugm_nsclc_prognosis_combined[col], errors="coerce"
        ).astype("Int64")

# -----------------------------------------------------------------------------
# 2. HGUGM 2014 Follow-Up Projects
# -----------------------------------------------------------------------------
f2014 = BASE_DIR / "HGUGM 2014 projects follow up data (final).xlsx"
hgugm_2014_pancreas = load_and_clean_excel(
    f2014, "Pancreatic cancer", "Pancreas_2014", header=3, apply_title_case=True
)
hgugm_2014_nsclc = load_and_clean_excel(
    f2014, "NSCLC", "NSCLC_2014", header=2, apply_title_case=True
)
hgugm_2014_ovarian = load_and_clean_excel(
    f2014, "Ovarian cancer", "Ovarian_2014", header=3, apply_title_case=True
)
hgugm_2014_hnscc = load_and_clean_excel(
    f2014, "Head & neck SCC", "HNSCC_2014", header=2, apply_title_case=True
)

# -----------------------------------------------------------------------------
# 3. HCC Follow-Up Data
# -----------------------------------------------------------------------------
f_hcc = BASE_DIR / "HCC follow up data HGUGM.xlsx"
hcc_follow_up = load_and_clean_excel(
    f_hcc, "Hoja1", "HCC_FollowUp", header=0, apply_title_case=True
)

# -----------------------------------------------------------------------------
# 4. Prostate Cancer
# -----------------------------------------------------------------------------
f_prostate = BASE_DIR / "HGUGM prostate cancer FFPE blocks - follow up data.xlsx"
hgugm_prostate_cancer = load_and_clean_excel(
    f_prostate, "source", "Prostate_Cancer", header=0, apply_title_case=True
).iloc[::2, :]

# -----------------------------------------------------------------------------
# 5. Renal Carcinoma
# -----------------------------------------------------------------------------
f_renal = BASE_DIR / "HGUGM RENAL CARCINOMA FOLLOW-UP DATA-1.xlsx"
hgugm_renal_carcinoma = load_and_clean_excel(
    f_renal, "Hoja1", "Renal_Carcinoma", header=0, apply_title_case=True
)

# -----------------------------------------------------------------------------
# 6. Gastric Cancer
# -----------------------------------------------------------------------------
f_gastric = BASE_DIR / "HGUGM GASTRIC CANCER FOLLOW UP DATA.xlsx"
hgugm_gastric_cancer = load_and_clean_excel(
    f_gastric, "Hoja1", "Gastric_Cancer", header=0
)

# -----------------------------------------------------------------------------
# 7. Bladder Cancer
# -----------------------------------------------------------------------------
f_bladder = BASE_DIR / "HGUGM bladder cancer follow up data.xlsx"
hgugm_bladder_cancer = load_and_clean_excel(
    f_bladder, "Hoja1", "Bladder_Cancer", header=0
)

# -----------------------------------------------------------------------------
# Global Priority Column Ordering Definition
# -----------------------------------------------------------------------------
dataframes = {
    "NSCLC_Prognosis": (
        hgugm_nsclc_prognosis_combined,
        ["Sample_ID", "Gender", "Ethnicity", "Age"],
    ),
    "Pancreas_2014": (
        hgugm_2014_pancreas,
        ["Sample_ID", "Gender", "Age"],
    ),
    "Ovarian_2014": (
        hgugm_2014_ovarian,
        ["Sample_ID", "Gender", "Ethnicity", "Age"],
    ),
    "NSCLC_2014": (
        hgugm_2014_nsclc,
        ["Sample_ID", "Gender", "Ethnicity", "Age"],
    ),
    "HNSCC_2014": (
        hgugm_2014_hnscc,
        ["Sample_ID", "Gender", "Ethnicity", "Age"],
    ),
    "HCC_FollowUp": (
        hcc_follow_up,
        ["Sample_ID", "Nodal_Mets", "Block_A", "Gender", "Ethnicity", "Age"],
    ),
    "Prostate_Cancer": (
        hgugm_prostate_cancer,
        ["Sample_ID", "Gender", "Ethnicity", "Age"],
    ),
    "Renal_Carcinoma": (
        hgugm_renal_carcinoma,
        [
            "Sample_ID",
            "Block_A",
            "Block_B",
            "Block_C",
            "Block_D",
            "Block_E",
            "Block_F",
            "Gender",
            "Ethnicity",
            "Age",
        ],
    ),
    "Gastric_Cancer": (
        hgugm_gastric_cancer,
        [
            "Sample_ID",
            "Nodal_Mets",
            "Block_A",
            "Block_B",
            "Block_C",
            "Block_D",
            "Block_E",
            "Block_F",
            "Met_Block_1",
            "Met_Block_2",
            "Met_Block_3",
            "Gender",
            "Ethnicity",
            "Age",
        ],
    ),
    "Bladder_Cancer": (
        hgugm_bladder_cancer,
        [
            "Sample_ID",
            "Block_A",
            "Block_B",
            "Block_C",
            "Block_D",
            "Block_E",
            "Patient_ID",
            "Gender",
            "Ethnicity",
            "Age",
        ],
    ),
}

final_dfs = {}

for sheet_name, (df, priority_cols) in dataframes.items():
    # Priority re-ordering
    existing_priority = [c for c in priority_cols if c in df.columns]
    remaining_cols = [c for c in df.columns if c not in existing_priority]
    ordered_df = df[existing_priority + remaining_cols]

    # Clean whitespace from column headers
    ordered_df.columns = ordered_df.columns.astype(str).str.strip()

    # Filter out empty/unnamed columns
    ordered_df = ordered_df.loc[
        :, ~ordered_df.columns.str.contains("^Unnamed", case=False, na=False)
    ]
    ordered_df = ordered_df.loc[
        :, ordered_df.columns.notnull() & (ordered_df.columns != "")
    ]

    # Convert empty/blank text cells to NaN
    ordered_df = ordered_df.replace(r"^\s*$", np.nan, regex=True)
    ordered_df = ordered_df.replace(["NR", "nan", "None", "\u3000"], np.nan)

    # Drop columns that have NO valid values across all rows
    ordered_df = ordered_df.dropna(how="all", axis=1)

    final_dfs[sheet_name] = ordered_df

# -----------------------------------------------------------------------------
# Write Clean Output to Excel
# -----------------------------------------------------------------------------
excel_output_path = Path(
    "C:/Users/medaamm/OneDrive - University of Leeds/UOL Data Manager/database_management/data_files/non_leeds_data/metadata/hgugm_inventories.xlsx"
)
excel_output_path.parent.mkdir(parents=True, exist_ok=True)

with pd.ExcelWriter(excel_output_path, engine="openpyxl") as writer:
    for sheet_name, df in final_dfs.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        worksheet = writer.sheets[sheet_name]

        # Column widths
        width = 25
        for col_num in range(1, len(df.columns) + 1):
            col_letter = get_column_letter(col_num)
            worksheet.column_dimensions[col_letter].width = width

        # Table styling
        max_row, max_col = df.shape
        if max_row > 0 and max_col > 0:
            max_col_letter = get_column_letter(max_col)
            tab = Table(
                displayName=sheet_name.replace(" ", "_"),
                ref=f"A1:{max_col_letter}{max_row + 1}",
            )
            style = TableStyleInfo(
                name="TableStyleMedium9",
                showFirstColumn=False,
                showLastColumn=False,
                showRowStripes=True,
                showColumnStripes=True,
            )
            tab.tableStyleInfo = style
            worksheet.add_table(tab)

# -----------------------------------------------------------------------------
# Upload to Google Cloud Engine
# -----------------------------------------------------------------------------
for sheet_name, df in final_dfs.items():
    df.to_sql(
        name=sheet_name,
        con=engine,
        schema="hgugm",
        if_exists="replace",
        index=False,
    )
    print(f"Uploaded {sheet_name} to GCP successfully.")

# -----------------------------------------------------------------------------
# Create a dictionary of Sample ID and file name
# -----------------------------------------------------------------------------
# In[]:
sampleID_filename_dict = {}
for df in [
    hgugm_nsclc_prognosis_combined,
    hgugm_2014_pancreas,
    hgugm_2014_hnscc,
    hgugm_2014_nsclc,
    hgugm_2014_ovarian,
    hgugm_bladder_cancer,
    hgugm_gastric_cancer,
    hgugm_prostate_cancer,
    hgugm_renal_carcinoma,
    hcc_follow_up,
]:
    # map each sample id to the file name (sheet name) it came from
    for sheet_name, (df, _) in dataframes.items():
        if df.equals(df):
            for id in df["Sample_ID"]:
                sampleID_filename_dict[id] = sheet_name

# -----------------------------------------------------------------------------
# Upload the dictionary to GCP as a table
# -----------------------------------------------------------------------------
sampleID_filename_df = pd.DataFrame(
    list(sampleID_filename_dict.items()), columns=["Sample_ID", "File_Name"]
)
sampleID_filename_df.to_sql(
    name="sample_id_filename_map",
    con=engine,
    schema="hgugm",
    if_exists="replace",
    index=False,
)
print("Uploaded sample ID to filename mapping to GCP successfully.")

# %%
