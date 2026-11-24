# Import the required libraries
import sys
import warnings
from pathlib import Path
from sqlalchemy.types import Date
import pandas as pd

# Disable UserWarnings and FutureWarnings
warnings.filterwarnings("ignore", category=UserWarning)
warnings.filterwarnings("ignore", category=FutureWarning)

# Add the project directory to the system path
current_dir = Path(__file__).resolve()
project_root = next(
    (p for p in Path(".").resolve().parents if (p / "pyproject.toml").exists()),
    Path(".").resolve(),
)
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))
code_files = project_root / "code_files"
non_leeds_metadata = project_root / "data_files" / "non_leeds_data" / "metadata"

# Import the required functions from the files
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

# -----------------------------------------------------------------------------------------------------------------------------------------------------

# import the NHSB data files
hrpc = pd.read_excel(
    non_leeds_metadata / "69573029 HRPC with follow up & sequential PSA data.xlsx",
    header=11,
    sheet_name="Follow up data",
)
# remove the extra rows
hrpc = hrpc.iloc[:89, :]
# remove the blank rows
hrpc = hrpc.dropna(how="all")


nhsb_505_prostate = pd.read_excel(
    non_leeds_metadata
    / "NHSBAC ACP 89 NHSB 505 Prostate Full Follow Up COMPLETE_To Release.xlsx",
    header=0,
    sheet_name="Sheet1",
)
# remove the extra rows
nhsb_505_prostate = nhsb_505_prostate.iloc[:15, :]
# remove the blank rows
nhsb_505_prostate = nhsb_505_prostate.dropna(how="all")


nhsb_261_met_prostate = pd.read_excel(
    non_leeds_metadata
    / "NHSBAC ACP89 NHSB261 Met prostate ca full follow up - to release.xlsx",
    header=0,
    sheet_name="Full Follow Up",
)
# remove the blank rows
nhsb_261_met_prostate = nhsb_261_met_prostate.dropna(how="all")


nhsb_tri_samples = pd.read_excel(
    non_leeds_metadata / "NHSBAC Tri_Samples_092018 histopath data.xlsx",
    header=0,
    sheet_name="Samples",
)
nhsb_tri_samples = nhsb_tri_samples.loc[
    nhsb_tri_samples.Origin.isin(
        [
            "NHSB",
            "HGUGM,BMS",
            "HGUGM,Bicester",
            "Bicester",
            "HGUGM",
            "HGUGM,BMS,Bicester",
            "BMS,Bicester",
            "BMS",
            "LBIH",
            "CARDIFF",
            "Korea",
        ]
    )
]

# ---------------------------------------------------------------------------------------------------------------------------

colnames_list = []
cols_to_drop = [
    "Collection Date.1",
    "Collection Date2",
    "Diagnosis.1",
    "Diagnosis3",
    "Geographic Origin.1",
    "Geographic Origin4",
    "Surgery, Site And Side5",
]
column_name_map = {
    "Adjuvant Chemotx Regimen 1": "Adjuvant_ChemoTx_Regimen1",
    "Adjuvant Chemotx Regimen 2": "Adjuvant_ChemoTx_Regimen2",
    "Adjuvant Treatment Regimen 1": "Adjuvant_Tx_Regimen1",
    "Adjuvant Treatment Regimen 2": "Adjuvant_Tx_Regimen2",
    "Adjuvant/ Neo-Adjuvant Chemotx Regimen 1": "Adjuvant_Neoadjuvant_ChemoTx_Regimen1",
    "Adjuvant/Neo-Adjuvant Chemotx Regimen 2": "Adjuvant_Neoadjuvant_ChemoTx_Regimen2",
    "Age": "AgeAtDX",
    "Age At Diagnosis": "AgeAtDX",
    "Ajcc Staging M": "M",
    "Ajcc Staging N": "N",
    "Ajcc Staging T": "T",
    "Alcohol History": "Alcohol_History",
    "Alcohol Status": "Alcohol_Status",
    "Cancer status (added)": "Cancer_Status",
    "Case number (added)": "Case_Number",
    "Cause Of Death": "Cause_Of_Death",
    "Collection Date": "Collection_Date",
    "Collection Time": "Collection_Time",
    "Contact Person": "Contact_Person",
    "Date last known alive (added)": "Date_Last_Known_Alive",
    "Death Date": "Death_Date",
    "Diagnosis": "Diagnosis",
    "EGFR Mutation": "EGFR_Mutation",
    "Estimated treatment dates (added)": "Estimated_Treatment_Dates",
    "Ethnicity": "Ethnicity",
    "Expiration Date": "Expiration_Date",
    "Gender": "Gender",
    "Geographic Origin": "Geographic_Origin",
    "Grade": "Grade",
    "Histopathologic Data/ Subtype": "Histopathologic_Subtype",
    "Immuno Met Tx": "Immuno_Met_Tx",
    "Immuno Tx (Regimen, type, duration)": "Immuno_Tx",
    "Immuno-histopathologic Data": "Immuno_Histopathologic_Data",
    "Immunohistopathologic Data": "Immuno_Histopathologic_Data",
    "KRAS Mutation": "KRAS_Mutation",
    "M": "M",
    "Molecular Data": "Molecular_Date",
    "N": "N",
    "NHSB comments (added)": "NHSB_Comments",
    "Origin": "Origin",
    "Original ID": "Original_ID",
    "Other Therapy": "Other_Therapy",
    "Other Therapy / Hormone Treatment": "Other_Therapy",
    "Patient ID": "Patient_ID",
    "Period between disease progression in months (added)": "Period_Between_Disease_Progression_Months",
    "Prefix": "Prefix",
    "Radiation 1: Type": "Radiation1_Type",
    "Radiation 2: Type": "Radiation2_Type",
    "Sample Date": "Collection_Date",
    "Sample ID": "Sample_ID",
    "Sample Type": "Sample_Type",
    "Set Availability": "Set_Availability",
    "Smoking History": "Smoking_History",
    "Smoking Status": "Smoking_Status",
    "Specimen 1 Collection Date": "Collection_Date",
    "Specimen 1 Type": "Sample_Type",
    "Stage": "StageAtDX",
    "Stage At Diagnosis": "StageAtDX",
    "Study Name": "Study_Name",
    "Sub-Type": "Histopathologic_Subtype",
    "Surgery, Site And Side": "Surgery_Site_Side",
    "Surgery, Site And Side/Other info": "Surgery_Site_Side",
    "Surgery, Site and Side": "Surgery_Site_Side",
    "Surgery, Site and Side/Other info": "Surgery_Site_Side",
    "Survival status (added)": "Survival_Status",
    "T": "T",
    "TriStar ID": "Sample_ID",
    "Tumour Grade (added)": "Grade",
    "Type": "Sample_Type",
    "mCRPC or mHSPC (added)": "mCRPC_mHSPC",
}

dfs = [hrpc, nhsb_505_prostate, nhsb_261_met_prostate, nhsb_tri_samples]

for i, df in enumerate(dfs):
    df = df.drop(columns=cols_to_drop, errors="ignore")
    df = df.rename(columns=column_name_map)
    df = df.loc[:, ~df.columns.duplicated(keep="first")]
    dfs[i] = df
    colnames_list.append(df.columns.to_list())

hrpc, nhsb_505_prostate, nhsb_261_met_prostate, nhsb_tri_samples = dfs

# ---------------------------------------------------------------------------------------------------------------------------

colnames = {col for lst in colnames_list for col in lst}
col_order = [
# Identifiers
'Prefix',
'Case_Number',
'Study_Name',
'Origin',
'Original_ID',
'Patient_ID',
'Sample_ID',

# Sample and Tumour Details
'Set_Availability',
'Sample_Type',
'Cancer_Status',
'Collection_Date',
'Collection_Time',
'Contact_Person',
'Diagnosis',
'Expiration_Date',
'Grade',
'Histopathologic_Subtype',
'Immuno_Histopathologic_Data',
'T', 'N', 'M',
'StageAtDX',
'Surgery_Site_Side',

# Demographic Information & Lifestyle
'AgeAtDX',
'Alcohol_History',
'Alcohol_Status',
'Gender',
'Geographic_Origin',
'Ethnicity',
'Smoking_History',
'Smoking_Status',

# Treatment Information
'Estimated_Treatment_Dates',
'Adjuvant_Neoadjuvant_ChemoTx_Regimen1',
'Adjuvant_Neoadjuvant_ChemoTx_Regimen2',
'Adjuvant_Tx_Regimen1',
'Adjuvant_Tx_Regimen2',
'Adjuvant_ChemoTx_Regimen1',
'Adjuvant_ChemoTx_Regimen2',
'Immuno_Met_Tx',
'Immuno_Tx',
'Radiation1_Type',
'Radiation2_Type',
'Other_Therapy',
'Period_Between_Disease_Progression_Months',
'Survival_Status',

# Census
'Cause_Of_Death',
'Date_Last_Known_Alive',
'Death_Date',

# Molecular Information
'EGFR_Mutation',
'KRAS_Mutation',
'Molecular_Date',
'NHSB_Comments',
'mCRPC_mHSPC'
]

dfs = [df[[col for col in col_order if col in df.columns]] for df in dfs]



# ------------------------------------------------------------------------------------------------------------------------------
table_names = ["hrpc", "nhsb_505_prostate", "nhsb_261_met_prostate", "nhsb_tri_samples"]

for name, df in zip(table_names, dfs):
    # Identify all datetime columns (naive or timezone-aware)
    date_cols = list(df.select_dtypes(include=["datetime", "datetimetz"]).columns)
    
    # Keep as datetime64[ns], remove timezone, and zero out the time component
    for col in date_cols:
        df[col] = pd.to_datetime(df[col]).dt.tz_localize(None).dt.normalize()

    # Clean remaining object columns (explicitly skipping date_cols)
    for col in df.select_dtypes(include=["object"]).columns:
        if col in date_cols:
            continue
        if not df[col].apply(lambda x: isinstance(x, (str, type(None)))).all():
            df[col] = df[col].astype(str)

    # Map date columns to SQL DATE
    sql_dtypes = {col: Date() for col in date_cols}

    # Indented inside the loop so all 4 tables are uploaded
    with engine.begin() as conn:
        df.to_sql(
            name,
            con=conn,
            schema="nhsb",
            if_exists="replace",
            index=False,
            dtype=sql_dtypes,
        )