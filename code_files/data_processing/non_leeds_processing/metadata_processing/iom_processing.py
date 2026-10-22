# -----------------------------------------------------------------------------
# Setup & Imports
# -----------------------------------------------------------------------------

# Import the required libraries
import sys
from pathlib import Path

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

pd.set_option("display.max_columns", None)
pd.set_option("display.max_rows", None)

env_path = project_root / "code_files" / "utilities" / "io" / ".env"
engine = create_gcloud_engine(path=env_path)

BASE_DIR = Path("C:/Users/medaamm/Box/Data file (AA)/IOM-Aviano/Source Files")

# -----------------------------------------------------------------------------
# Sheet-Specific Rename Dictionaries (Run FIRST to avoid collisions)
# -----------------------------------------------------------------------------

colnames_cll = [
    "Sample_ID",
    "Core_Reference",
    "Age",
    "Gender",
    "Diagnosis",
    "Date_Of_Diagnosis",
    "Site",
    "Stage",
    "Therapy_1",
    "N_Cycles",
    "Response_Therapy_1",
    "Relapse_Therapy_1",
    "Therapy_2",
    "Response_Therapy_2",
    "Relapse_Therapy_2",
    "Therapy_3",
    "Response_Therapy_3",
    "Relapse_Therapy_3",
    "Therapy_4",
    "Response_Therapy_4",
    "Relapse_Therapy_4",
    "Therapy_5",
]

# -----------------------------------------------------------------------------
# Preparing the DataFrames and Collecting Column Names
# -----------------------------------------------------------------------------

# Import all the dataframes
iom_multi_tumor_array_450 = pd.read_excel(
    r"C:/Users/medaamm/Box/Data file (AA)/IOM-Aviano/Source Files/69571165 (TA1787, 1788) Multi tumor array 450 with histopath data CHECKED (14.11.03).xlsx",
    sheet_name="Data sheet Histopath Data",
    header=10,
)

iom_multi_tumor_array_120 = pd.read_excel(
    r"C:/Users/medaamm/Box/Data file (AA)/IOM-Aviano/Source Files/69572065-4340 Multi tumor array 120 (TA4340).xlsx",
    sheet_name="Sheet1",
    header=11,
)

iom_crc_data_histopath = pd.read_excel(
    r"C:/Users/medaamm/Box/Data file (AA)/IOM-Aviano/Source Files/IOM & Aviano follow up data/IOM CRC Data With Follow up (TMA) with STAGE.xlsx",
    sheet_name="Histopath data",
    header=0,
    index_col=0,
).reset_index()
iom_crc_data_follow_up = pd.read_excel(
    r"C:/Users/medaamm/Box/Data file (AA)/IOM-Aviano/Source Files/IOM & Aviano follow up data/IOM CRC Data With Follow up (TMA) with STAGE.xlsx",
    sheet_name="Follow up data",
    header=4,
    index_col=0,
).reset_index()

atcg_1065 = pd.read_excel(
    r"C:/Users/medaamm/Box/Data file (AA)/IOM-Aviano/Source Files/ATCG TMA data sheets 1068 1072 1076 1109.xlsx",
    sheet_name="ATCG-1065 (9-9-13)",
    header=1,
)
atcg_1067 = pd.read_excel(
    r"C:/Users/medaamm/Box/Data file (AA)/IOM-Aviano/Source Files/ATCG TMA data sheets 1068 1072 1076 1109.xlsx",
    sheet_name="ACTG-1067 (9-9-13)",
    header=1,
)
atcg_1068 = pd.read_excel(
    r"C:/Users/medaamm/Box/Data file (AA)/IOM-Aviano/Source Files/ATCG TMA data sheets 1068 1072 1076 1109.xlsx",
    sheet_name="ATCG-1068 (9-9-13)",
    header=1,
)
adn_100 = pd.read_excel(
    r"C:/Users/medaamm/Box/Data file (AA)/IOM-Aviano/Source Files/ATCG TMA data sheets 1068 1072 1076 1109.xlsx",
    sheet_name="ADN 100 (9-9-13)",
    header=1,
)
atcg_1072 = pd.read_excel(
    r"C:/Users/medaamm/Box/Data file (AA)/IOM-Aviano/Source Files/ATCG TMA data sheets 1068 1072 1076 1109.xlsx",
    sheet_name="ATCG-1072 (9-9-13)",
    header=1,
)
atcg_1076 = pd.read_excel(
    r"C:/Users/medaamm/Box/Data file (AA)/IOM-Aviano/Source Files/ATCG TMA data sheets 1068 1072 1076 1109.xlsx",
    sheet_name="10-ATCG-1076 (9-9-13)",
    header=1,
)
atcg_1109 = pd.read_excel(
    "C:/Users/medaamm/Box/Data file (AA)/IOM-Aviano/Source Files/ATCG TMA data sheets 1068 1072 1076 1109.xlsx",
    sheet_name=" ATCG-1109 (9-9-13)",
    header=1,
)

tmo_bladder = pd.read_excel(
    r"C:/Users/medaamm/Box/Data file (AA)/IOM-Aviano/Source Files/TMO samples for sequencing (KR) list for selection.xlsx",
    sheet_name="bladder",
    header=0,
)
tmo_breast = pd.read_excel(
    "C:/Users/medaamm/Box/Data file (AA)/IOM-Aviano/Source Files/TMO samples for sequencing (KR) list for selection.xlsx",
    sheet_name="breast",
    header=0,
)
tmo_colorectal = pd.read_excel(
    r"C:/Users/medaamm/Box/Data file (AA)/IOM-Aviano/Source Files/TMO samples for sequencing (KR) list for selection.xlsx",
    sheet_name="colorectal",
    header=0,
)
tmo_glioma = pd.read_excel(
    r"C:/Users/medaamm/Box/Data file (AA)/IOM-Aviano/Source Files/TMO samples for sequencing (KR) list for selection.xlsx",
    sheet_name="glioma",
    header=0,
)
tmo_cutaneous = pd.read_excel(
    r"C:/Users/medaamm/Box/Data file (AA)/IOM-Aviano/Source Files/TMO samples for sequencing (KR) list for selection.xlsx",
    sheet_name="cutaneous",
    header=0,
)
tmo_prostate = pd.read_excel(
    r"C:/Users/medaamm/Box/Data file (AA)/IOM-Aviano/Source Files/TMO samples for sequencing (KR) list for selection.xlsx",
    sheet_name="prostate",
    header=0,
)
tmo_papillarythyroid = pd.read_excel(
    r"C:/Users/medaamm/Box/Data file (AA)/IOM-Aviano/Source Files/TMO samples for sequencing (KR) list for selection.xlsx",
    sheet_name="papillary thyroid",
    header=0,
)

TA_1081 = pd.read_excel(
    r"C:/Users/medaamm/Box/Data file (AA)/IOM-Aviano/Source Files/69571081 CRC primary, non-nodal mets (non donor matched) & normal (non-malignant) colon with histopath data (TA1081).xlsx",
    sheet_name="TA1081",
    header=12,
)

iom_crc_KR = pd.read_excel(
    r"C:/Users/medaamm/Box/Data file (AA)/IOM-Aviano/Source Files/IOM CRC KR-02 - KR-80 UKE histopath & KRAS data.xlsx",
    sheet_name="Data sheet",
    header=0,
)

iom_nhl_dlbcl = pd.read_excel(
    r"C:/Users/medaamm/Box/Data file (AA)/IOM-Aviano/Source Files/69571071 69571075 NHL TMA follow up data (revised).xlsx",
    sheet_name="DLBCL",
    header=9,
)
iom_nhl_fl = pd.read_excel(
    r"C:/Users/medaamm/Box/Data file (AA)/IOM-Aviano/Source Files/69571071 69571075 NHL TMA follow up data (revised).xlsx",
    sheet_name="FL",
    header=10,
)
iom_nhl_cll = pd.read_excel(
    r"C:/Users/medaamm/Box/Data file (AA)/IOM-Aviano/Source Files/69571071 69571075 NHL TMA follow up data (revised).xlsx",
    sheet_name="CLL",
    header=10,
)

nhl_tma_2_UK = pd.read_excel(
    r"C:/Users/medaamm/Box/Data file (AA)/IOM-Aviano/Source Files/69572567 NHL TMA 2 (TA2567).xlsx",
    sheet_name="Data sheet - UK",
    header=10,
)
nhl_tma_2_Italy = pd.read_excel(
    r"C:/Users/medaamm/Box/Data file (AA)/IOM-Aviano/Source Files/69572567 NHL TMA 2 (TA2567).xlsx",
    sheet_name="Data sheet - Italy",
    header=11,
)
iom_TS21_1721 = pd.read_excel(
    r"C:/Users/medaamm/Box/Data file (AA)/IOM-Aviano/Source Files/IOM & Aviano follow up data/2021 IOM follow up data TS-21.xlsx",
    sheet_name="17-21",
    header=4,
)
iom_TS21_1821 = pd.read_excel(
    r"C:/Users/medaamm/Box/Data file (AA)/IOM-Aviano/Source Files/IOM & Aviano follow up data/2021 IOM follow up data TS-21.xlsx",
    sheet_name="18-21",
    header=4,
)
iom_TS21_1921 = pd.read_excel(
    r"C:/Users/medaamm/Box/Data file (AA)/IOM-Aviano/Source Files/IOM & Aviano follow up data/2021 IOM follow up data TS-21.xlsx",
    sheet_name="19-21",
    header=4,
)
iom_TS21_2021 = pd.read_excel(
    r"C:/Users/medaamm/Box/Data file (AA)/IOM-Aviano/Source Files/IOM & Aviano follow up data/2021 IOM follow up data TS-21.xlsx",
    sheet_name="20-21",
    header=4,
)
iom_TS21_2121 = pd.read_excel(
    r"C:/Users/medaamm/Box/Data file (AA)/IOM-Aviano/Source Files/IOM & Aviano follow up data/2021 IOM follow up data TS-21.xlsx",
    sheet_name="21-21",
    header=4,
)
iom_TS1819 = pd.read_excel(
    r"C:/Users/medaamm/Box/Data file (AA)/IOM-Aviano/Source Files/IOM & Aviano follow up data/IOM 2019 samples follow up data TS-18 & TS-19.xlsx",
    sheet_name="IOM Samples",
    header=0,
)


# Define an empty list for the column names
all_columns = []

dataframes = {
    "iom_multi_tumor_array_450": iom_multi_tumor_array_450,
    "iom_multi_tumor_array_120": iom_multi_tumor_array_120,
    "iom_crc_data_histopath": iom_crc_data_histopath,
    "iom_crc_data_follow_up": iom_crc_data_follow_up,
    "atcg_1065": atcg_1065,
    "atcg_1067": atcg_1067,
    "atcg_1068": atcg_1068,
    "adn_100": adn_100,
    "atcg_1072": atcg_1072,
    "atcg_1076": atcg_1076,
    "atcg_1109": atcg_1109,
    "tmo_bladder": tmo_bladder,
    "tmo_breast": tmo_breast,
    "tmo_colorectal": tmo_colorectal,
    "tmo_glioma": tmo_glioma,
    "tmo_cutaneous": tmo_cutaneous,
    "tmo_prostate": tmo_prostate,
    "tmo_papillarythyroid": tmo_papillarythyroid,
    "TA_1081": TA_1081,
    "iom_crc_KR": iom_crc_KR,
    "iom_nhl_dlbcl": iom_nhl_dlbcl,
    "iom_nhl_fl": iom_nhl_fl,
    "iom_nhl_cll": iom_nhl_cll,
    "nhl_tma_2_UK": nhl_tma_2_UK,
    "nhl_tma_2_Italy": nhl_tma_2_Italy,
    "iom_TS21_1721": iom_TS21_1721,
    "iom_TS21_1821": iom_TS21_1821,
    "iom_TS21_1921": iom_TS21_1921,
    "iom_TS21_2021": iom_TS21_2021,
    "iom_TS21_2121": iom_TS21_2121,
    "iom_TS1819": iom_TS1819,
}

for name, df in dataframes.items():
    # Remove unnamed columns
    df = df.loc[:, ~df.columns.str.contains("^Unnamed", case=False, na=False)]
    # Convert all columns to title case to standardise
    df.columns = [col.title() for col in df.columns]
    # Append the column names to the list
    for col in df.columns:
        all_columns.append(col)

# Strip spaces from all column names
all_columns = [col.strip() for col in all_columns]

# Convert the list of columns to a set
all_columns = set(all_columns)


# -----------------------------------------------------------------------------
# Master Global Mapping Dictionary
# -----------------------------------------------------------------------------

GLOBAL_COLUMN_MAP = {
    "#": "Sample_ID",
    "Id Tristar": "Sample_ID",
    "# Id Prog.": "Sample_ID",
    "# Samples": "Sample_ID",
    "Accession Number": "Accession_Number",
    "Age At Specimen Collection": "Age_At_Specimen_Collection",
    "Alcohol History": "Alcohol_History",
    "Anno.": "Anno.",
    "Bcl2": "BCL2",
    "Bcl6": "BCL6",
    "Bl": "Block_Number",
    "Bone Marrow": "Bone_Marrow",
    "Bulky": "Bulky",
    "Cd10": "CD10",
    "Cd20": "CD20",
    "Cd79A": "CD79A",
    "Collection Matrix": "Collection_Matrix",
    "Comment - Specimen": "Comment_Specimen",
    "Comments": "Comments",
    "Core Id": "Core_ID",
    "Core Ref": "Core_Reference",
    "Current Label": "Current_Label",
    "Date Last Follow Up": "Date_Last_Follow_Up",
    "Year Of Collection": "Year_Of_Collection",
    "Date Of Collection": "Date_Of_Collection",
    "Date Of Death": "Date_Of_Death",
    "Date Of Diagnosis": "Date_Of_Diagnosis",
    "Date Of Relapse": "Date_Of_Relapse",
    "Death Due To Lymphoma?": "Death_Due_To_Lymphoma",
    "Dfs": "DFS",
    "Diagnosi": "Diagnosis",
    "Diagnosis": "Diagnosis",
    "Diagnosis - Specimen": "Diagnosis_Specimen",
    "Diagnosis - Subject": "Diagnosis_Subject",
    "Divided": "Divided",
    "Er": "ER",
    "Ethnicity": "Ethnicity",
    "Extranodal Localization": "Extranodal_Localization",
    "Follow Up 1": "Follow_Up_1",
    "Follow Up 2": "Follow_Up_2",
    "Follow-Up Months": "Follow_Up_Months",
    "Follow-Up Result": "Follow_Up_Result",
    "Gender": "Gender",
    "Gr.": "Grade",
    "Grade": "Grade",
    "Grading": "Grade",
    "H2": "H2",
    "Her2": "HER2",
    "Her2 (Ihc)": "HER2_IHC",
    "Id": "Block_ID",
    "Id #": "Sample_ID",
    "Id # Iom-": "Sample_ID",
    "Id Bl": "Block_Number",
    "Id Origine": "Origin",
    "Id Prog": "Sample_ID",
    "Ihc Molecular Data": "IHC_Molecular_Data",
    "Ii Terapy": "Therapy",
    "International Prognostic Index": "International_Prognostic_Index",
    "Isto.": "Isto.",
    "Ki": "KI67",
    "Ki 67": "KI67",
    "Ki-67": "KI67",
    "Last Follow Up": "Last_Follow_Up",
    "M": "M",
    "Maintainance Therapy With Rituximab ?": "Maintenance_Therapy_With_Rituximab",
    "Meds Rx 1": "Meds_Rx_1",
    "Metastatic Or Primary": "Metastatic_Or_Primary",
    "Mum1": "MUM1",
    "Mut [Kras]": "Mutation_KRAS",
    "Mutation Data": "Mutation_Data",
    "N": "N",
    "N Cycles": "N_Cycles",
    "N. Of Cycles": "N_Cycles",
    "Note": "Note",
    "Organ": "Organ",
    "Origin": "Origin",
    "Original Id": "Original_ID",
    "Originating Site - Specimen": "Originating_Site_Specimen",
    "Os": "OS",
    "Others": "Others",
    "Paired Metastasis": "Paired_Metastasis",
    "Paired Normal": "Paired_Normal",
    "Paired Tumor": "Paired_Tumor",
    "Pet At Diagnosis": "PET_At_Diagnosis",
    "Pr": "PR",
    "Presence Follow Up Data": "Presence_Follow_Up_Data",
    "Presence Metastatic Samples": "Presence_Metastatic_Samples",
    "Project": "Project",
    "Race": "Race",
    "Radiotherapy": "Radiotherapy",
    "Relapse": "Relapse",
    "Relapse.1": "Relapse_1",
    "Relapse.2": "Relapse_2",
    "Response": "Response",
    "Response To Ii Line": "Response_To_Second_Line",
    "Response To Therapy": "Response_To_Therapy",
    "Response.1": "Response_1",
    "Response.2": "Response_2",
    "Rituximab": "Rituximab",
    "Rnascope": "RNAscope",
    "Sample Id#": "Sample_ID",
    "Second Treatment": "Second_Treatment",
    "Sex": "Gender",
    "Site": "Site",
    "Site Of Mets.": "Mets_Site",
    "Site Primitive": "Primary_Site",
    "Smoking History": "Smoking_History",
    "Specimen Biopsy Method": "Specimen_Biopsy_Method",
    "Specimen Collection Date": "Specimen_Collection_Date",
    "Specimen Collection Method": "Specimen_Collection_Method",
    "Specimen Source": "Specimen_Source",
    "Specimen Source Detail": "Specimen_Source_Detail",
    "Specimen Type": "Specimen_Type",
    "Stage": "Stage",
    "Start Treatment": "Start_Treatment",
    "Start Treatment 1": "Start_Treatment_1",
    "Status [Kras]": "Status_KRAS",
    "Study Id": "Study_ID",
    "Sub-Type": "Sub_Type",
    "Subject Id": "Subject_ID",
    "Symptoms": "Symptoms",
    "Survival Month": "Survival_Months",
    "T": "T",
    "Therapy": "Therapy",
    "Therapy 1": "Therapy_1",
    "Therapy 2": "Therapy_2",
    "Therapy.1": "Therapy_1",
    "Therapy.2": "Therapy_2",
    "Tissue Type": "Tissue_Type",
    "Transplantation": "Transplantation",
    "Treatment": "Treatment",
    "Treatment (Schedula)": "Treatment_Schedule",
    "Trk Shipment": "TRK_Shipment",
    "Trk Spedizione": "TRK_Spedizione",
    "Trk-Incoming": "TRK_Incoming",
    "Tumor Grade": "Tumor_Grade",
    "Type": "Type",
    "Volume": "Volume",
    "Volume Uom": "Volume_UOM",
}

# -----------------------------------------------------------------------------
# Clean & Read Generic Helper Function
# -----------------------------------------------------------------------------


def load_and_clean_excel(dataframes):
    """Generic helper function to process and rename the dataframes."""
    cleaned_dataframes = {}

    for name, df in dataframes.items():
        # Remove unnamed columns
        df = df.loc[:, ~df.columns.str.contains("^Unnamed", case=False, na=False)]

        # Standardize case and strip whitespace to match dictionary keys
        df.columns = [str(col).title().strip() for col in df.columns]

        # Apply sheet-specific renaming first (e.g., CLL dataset)
        if "cll" in name.lower():
            df.columns = colnames_cll

        # Apply global column mapping for the others
        else:
            df = df.rename(columns=GLOBAL_COLUMN_MAP)

        # Remove all spaces from the ID columns values
        id_cols = ["Sample_ID", "Block_Number", "Block_ID", "Original_ID"]
        for col in id_cols:
            if col in df.columns:
                df[col] = df[col].astype(str).str.replace(r"\s+", "", regex=True)

        # Store the updated DataFrame back into the dictionary
        cleaned_dataframes[name] = df

    return cleaned_dataframes


# Execute the function to replace the dictionary with the cleaned DataFrames
dataframes = load_and_clean_excel(dataframes)
# -----------------------------------------------------------------------------
# Write Clean Output to Excel
# -----------------------------------------------------------------------------
excel_output_path = Path(
    "C:/Users/medaamm/OneDrive - University of Leeds/UOL Data Manager/database_management/data_files/non_leeds_data/metadata/iom_inventories.xlsx"
)
excel_output_path.parent.mkdir(parents=True, exist_ok=True)

with pd.ExcelWriter(excel_output_path, engine="openpyxl") as writer:
    for sheet_name, df in dataframes.items():
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
for sheet_name, df in dataframes.items():
    df.to_sql(
        name=sheet_name,
        con=engine,
        schema="iom_aviano",
        if_exists="replace",
        index=False,
    )
    print(f"Uploaded {sheet_name} to GCP successfully.")

# -----------------------------------------------------------------------------
# Create a dictionary of Sample ID and file name
# -----------------------------------------------------------------------------

sampleID_filename_dict = {}
for name, df in dataframes.items():
    id_cols = ["Sample_ID", "Block_Number", "Block_ID", "Original_ID"]
    for col in id_cols:
        if col in df.columns:
            for val in df[col].dropna().unique():
                sampleID_filename_dict[str(val)] = name

# -----------------------------------------------------------------------------
# Upload the dictionary to GCP as a table
# -----------------------------------------------------------------------------
sampleID_filename_df = pd.DataFrame(
    list(sampleID_filename_dict.items()), columns=["Sample_ID", "File_Name"]
)
sampleID_filename_df.to_sql(
    name="sample_id_filename_map",
    con=engine,
    schema="iom_aviano",
    if_exists="replace",
    index=False,
)
print("Uploaded sample ID to filename mapping to GCP successfully.")
