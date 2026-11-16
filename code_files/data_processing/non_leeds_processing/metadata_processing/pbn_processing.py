# Import the required libraries
from openpyxl import load_workbook
from copy import copy
from pathlib import Path
import pandas as pd
import io

# Define the path to the PBN folder
folder = Path("C:/Users/medaamm/Box/PBN manifests & data (AA file)")

# Import functions from project directory
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# Setup GCP connection and define the env path
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()
env_path = project_root / "code_files" / "utilities" / "io" / ".env"

# create the workbook to file name mapping
wb_filename_mapping = {
    "PBI 020-1516-0946 Plasma samples list shipped 15.02.22.xlsx": "PBI_020_1516_0946",
    "PBI FFPE manifest 11 cases replacements FFPEs.xlsx": "PBI_FFPE",
    "PBI Ukraine manifest September 2021_49 cases_shipped.xlsx": "PBI_Ukraine_Manifest_Sept_2021",
    "PBN manifest March 2022.xlsx": "PBN_Manifest_March_2022",
    "PBN Plasma vials list .xlsx": "PBN_Plasma_Vials_List",
    "PBN_DHL2836648476_24.12.2021_Lufthansa020-1516-0946_14.02.2022_PO1002151.xlsx": "PBN_DHL2836648476",
    "PBN_DHL5373309830_01.02.2021_TurkishAir235-1661-5432_26.02.2021_PO1002151.xlsx": "PBN_DHL5373309830",
    "PBN_DHL8338472236_30.09.2021_Lufthansa020-1516-0320_26.11.2021_PO1002151.xlsx": "PBN_DHL8338472236",
    "PBN_EmiratesAir176-3594-3331_31.03.2022_PO1002151.xlsx": "PBN_EmiratesAir176_3594_3331",
    "PBN_LOT080-37222430_21.06.2024_PO1002151_without rejects.xlsx": "PBN_LOT080_37222430",
    "PBN_LOT080-37222500_11.09.2024_PO1002157.xlsx": "PBN_LOT080_37222500",
    "PBN_manifest_TriStar_26-Jan-2021.xlsx": "PBN_manifest_TriStar_26_Jan_2021",
    "PBN_manifest_TriStar_26.01.2021_without rejects_ with price.xlsx":  "PBN_manifest_TriStar_26_01_2021",
    "Recieved list - PBN Manifest February 2024 (20240327)_HAWB EXP24020": "PBN_manifest_Feb_2024"
}

# create the sheet names mapping
sheet_name_mapping = {
    "samples shipped 15.02.22": "Samples_Data",
    "Sep-2021 replacements": "Replacement_Samples_Data",
    "plasma samples": "Plasma_Samples_Data",
    "Feb - 2022": "Samples_Data",
    "samples list": "Samples_List",
    "Plasma details PBN": "Plasma_Details_PBN",
    "New sets Dec 2021": "Samples_Dec_2021",
    "plasma ship 15.02.22": "Plasma_Data_15_2_22",
    "Additional FFPEs shipped 24.12": "Additional_FFPEs_Shipped_24_12",
    "Replasements (sets)": "Replacement_Samples_Data",
    "26-Jan-2021": "Samples_Data",
    "plasma samples": "Plasma_Samples_Data",
    "Mar 2022": "Samples_Data",
    "samples list": "Samples_List",
    "May2024": "Samples_Data",
    "Samples list": "Samples_List",
    "QC": "QC_Data",
    "080-37222500": "Samples_Data",
    "Samples list": "Samples_List",
    "26-Jan-2021": "Samples_Data",
    "26-Jan-2021": "Samples_Data",
    "Feb2024": "Samples_Data"
}

# create the column names mapping
col_mapping = {
    # Identifiers & Record Metadata
    '#': 'Record_No',
    'NO 1': 'Record_No',
    'Patient ID': 'PatientID',
    'Sample ID': 'SampleID',
    'Replacement samples ID': 'ReplacementSampleID',
    'Selected specimen Anonymised ID 9866N2': 'SpecimenAnonymisedID',
    'Site Code': 'SiteCode',
    'Country of collection': 'CountryOfCollection',
    'ICF date': 'ICF_Date',
    'Sample type': 'SampleType',

    # Commercial / Metadata
    'Pricing, USD': 'Price_USD',
    'Price, USD': 'Price_USD',
    'Manufacturer': 'Manufacturer',
    'Storage tube/vial type': 'StorageTubeType',
    'Type  of vacutainer': 'BloodCollectionTubeType',
    'Blood collection tubes': 'BloodCollectionTubeType',

    # Demographics
    'Demographics Age': 'Age',
    'Demographics Patient Race': 'Race',
    'Demographics Ethnicity': 'Ethnicity',
    'Demographics Gender': 'Gender',
    'Demographics Height, cm': 'Height_cm',
    'Demographics Weight, kg': 'Weight_kg',

    # Tissue Processing (FFPE & NAT)
    'FFPE tumor ID block 1': 'FFPE_Tumor_Block1_ID',
    'FFPE tumor ID block 2': 'FFPE_Tumor_Block2_ID',
    'FFPE NAT block 1': 'FFPE_NAT_Block1_ID',
    'Organ FFPE NAT': 'FFPE_NAT_Organ',
    'uterus': 'Uterus',
    'Block type\n(General/Mega)': 'BlockType',
    'Block ID of replacement received': 'ReplacementBlockID',
    'MPL-Tissue Sample weight, mg': 'Tissue_Weight_mg',
    'MPL-Tissue Date of tissue collection': 'Tissue_CollectionDate',
    'MPL-Tissue Time of tumour excision': 'Tissue_ExcisionTime',
    'MPL-Tissue Minutes (") before fixation to Formalin': 'Tissue_MinutesBeforeFixation',
    'MPL-Tissue Storage start date': 'Tissue_StorageStartDate',
    'MPL-Tissue Storage start time': 'Tissue_StorageStartTime',
    'MPL-Tissue Storage conditions': 'Tissue_StorageConditions',

    # Plasma Processing & Aliquots
    'Total plasma volume (ml)': 'Plasma_TotalVolume_mL',
    'Total volume, ml': 'Plasma_TotalVolume_mL',
    'Plasma ID / Volume (ml) Vial 1 ID': 'Plasma_Vial1_ID',
    'Plasma ID / Volume (ml) Vial 1 (ml)': 'Plasma_Vial1_Volume_mL',
    'Plasma ID / Volume (ml) Vial 2 ID': 'Plasma_Vial2_ID',
    'Plasma ID / Volume (ml) Vial 2 (ml)': 'Plasma_Vial2_Volume_mL',
    'Plasma ID / Volume (ml) Vial 3 ID': 'Plasma_Vial3_ID',
    'Plasma ID / Volume (ml) Vial 3 (ml)': 'Plasma_Vial3_Volume_mL',
    'Plasma ID / Volume (ml) Vial 4 ID': 'Plasma_Vial4_ID',
    'Plasma ID / Volume (ml) Vial 4 (ml)': 'Plasma_Vial4_Volume_mL',
    'Plasma ID / Volume (ml) Vial 5 ID': 'Plasma_Vial5_ID',
    'Plasma ID / Volume (ml) Vial 5 (ml)': 'Plasma_Vial5_Volume_mL',
    'Plasma ID / Volume (ml) Vial 6 ID': 'Plasma_Vial6_ID',
    'Plasma ID / Volume (ml) Vial 6 (ml)': 'Plasma_Vial6_Volume_mL',

    # Buffy Coat Aliquots
    'Buffy coat Vial ID': 'BuffyCoat_Vial1_ID',
    'Buffy coatVial (µL)': 'BuffyCoat_Vial1_Volume_uL',
    'Buffy coat ID / Volume (µL) Vial 1 ID': 'BuffyCoat_Vial1_ID',
    'Buffy coat ID / Volume (µL) Vial 1 (µL)': 'BuffyCoat_Vial1_Volume_uL',
    'Buffy coat ID / Volume (µL) Vial 2 ID': 'BuffyCoat_Vial2_ID',
    'Buffy coat ID / Volume (µL) Vial 2 (µL)': 'BuffyCoat_Vial2_Volume_uL',
    'Buffy coat ID / Volume (µL) Vial 3 ID': 'BuffyCoat_Vial3_ID',
    'Buffy coat ID / Volume (µL) Vial 3 (µL)': 'BuffyCoat_Vial3_Volume_uL',
    'Buffy coat ID / Volume (µL) Vial 4 ID': 'BuffyCoat_Vial4_ID',
    'Buffy coat ID / Volume (µL) Vial 4 (µL)': 'BuffyCoat_Vial4_Volume_uL',
    'Total buffy coat': 'BuffyCoat_TotalVolume_uL',

    # Blood Processing / Centrifugation
    'MPL-Blood Date of blood collection': 'Blood_CollectionDate',
    'MPL-Blood Time of blood collection': 'Blood_CollectionTime',
    'MPL-Blood Start time of processing': 'Blood_ProcessingStartTime',
    'MPL-Blood Centrifugation #1 date&time': 'Blood_Centrifugation1_DateTime',
    'MPL-Blood Centrifugation #1 duration/speed (g)': 'Blood_Centrifugation1_Speed_g',
    'MPL-Blood Centrifugation #2 date&time': 'Blood_Centrifugation2_DateTime',
    'MPL-Blood Centrifugation #2 duration/speed (g)': 'Blood_Centrifugation2_Speed_g',
    'MPL-Blood Storage start date': 'Blood_StorageStartDate',
    'MPL-Blood Storage start time': 'Blood_StorageStartTime',
    'MPL-Blood Storage start date&time': 'Blood_StorageStartDateTime',
    'MPL-Blood Storage conditions': 'Blood_StorageConditions',
    'MPL-Blood Hemolysis Index': 'Blood_HemolysisIndex',

    # Clinical & Pathology
    'Clinical diagnosis': 'ClinicalDiagnosis',
    'Localization': 'TumorLocalization',
    'Pathology Tumour Location': 'TumorLocation',
    'Pathology Surgical procedure': 'SurgicalProcedure',
    'Pathology Method of verification': 'MethodOfVerification',
    'Pathology Tumor size (cm)-Height  X width X Length': 'TumorSize_cm',
    'Pathology Distant metastasis location': 'DistantMetastasisLocation',
    'Pathology Date of diagnosis': 'DateOfDiagnosis',
    'Pathology Histologic Tumour Type': 'HistologicTumorType',
    'Pathology Grade': 'TumorGrade',
    'Pathology Cancer Stage': 'CancerStage',
    'Pathology T': 'TNM_T',
    'Pathology N': 'TNM_N',
    'Pathology M': 'TNM_M',
    'Gross description': 'GrossDescription',
    'Microscopic description': 'MicroscopicDescription',

    # Medical History & Treatments
    'Subject received anticancer therapy in the preceding 6 months': 'History_AnticancerTherapyPast6Mo',
    'Subject has a personal history of a malignancy of any type': 'History_PriorMalignancy',
    'Subject has previous history of organ transplantation or blood product transfusion': 'History_TransplantOrTransfusion',
    'Treatment Type of Therapy/Dose': 'Treatment_TypeAndDose',
    'Treatment Route': 'Treatment_Route',
    'Treatment Frequency': 'Treatment_Frequency',

    # Serology / Testing
    'Testing Hep B': 'Testing_HepB',
    'Testing Hep C': 'Testing_HepC',
    'Testing HIV': 'Testing_HIV',

    # Quality Control & Receipt
    'QC': 'QC_Status',
    'QC comment': 'QC_Comment',
    'QC \ncomment': 'QC_Comment',
    'QC results X': 'QC_Result_X',
    'QC results No tissue': 'QC_Result_NoTissue',
    'Slide\nQC': 'Slide_QC_Status',
    'QC acceptance (Indicate total plasma recd  (ml)': 'QC_PlasmaReceived_mL',
    'received\nFFPE': 'Received_FFPE',
    'received\nplasma': 'Received_Plasma',
    'received\nBuffy coat': 'Received_BuffyCoat',
}

# Define an empty dict for the combined file to sheet name mapping
combined_file_sheet_mapping = {}

# Define an empty list of PBN dataframes
pbn_dfs = []

# Loop through the PBN folder
for file in folder.glob("*.xlsx"):
    # Open each file in binary mode
    with open(file, "rb") as f:
        # read the contents of the file-like object into RAM as a raw bytes object
        # file_bytes acts like an open file but lives purely in RAM rather than on disk
        # it closes the original file handle immediately
        file_bytes = io.BytesIO(f.read())

    try:
        # load the workbook
        wb = load_workbook(file_bytes)
    except Exception as e:
        # if that fails, the file is corrupted
        print(f"⚠️ Skipping corrupted file: {file.name} - {e}")
        continue

    # Loop through the sheet names
    for sheet in wb.sheetnames:
        # extract the name of the workbook from the file path
        file_key = str(file).split("AA file)\\")[1]
        # create the combined file and sheet mapping by combining the file name with the sheet name from their respective mappings
        combined_file_sheet_mapping[sheet] = (
            f"{wb_filename_mapping.get(file_key)}_{sheet_name_mapping.get(sheet)}"
        )

        # load the worksheet
        ws = wb[sheet]

        # get the merged ranges from the worksheet
        merged_ranges = list(ws.merged_cells.ranges)

        # loop through the merged ranges
        for merged_range in merged_ranges:
            # for each merged range, get the boundary rows and columns
            min_col, min_row, max_col, max_row = merged_range.bounds
            # get the value from the top left cell of the merged range (the minimum i.e. top row and the minimum column i.e. first column)
            value = ws.cell(min_row, min_col).value
            # Unmerge the merged range
            ws.unmerge_cells(str(merged_range))
            # loop through the unmerged range and set the value of each cell equal to that of the top left cell
            for row in range(min_row, max_row + 1):
                for col in range(min_col, max_col + 1):
                    ws.cell(row, col).value = value

        # read in the cell values from the worksheet and store them as a list of tuples in memory
        data = list(ws.values)
        if not data:
            continue

        # read in the cell values as a dataframe
        df = pd.DataFrame(data, dtype=str)

        # -------------------------------------------------------------
        # Find the header row: first row where the first 3 cells have values
        # -------------------------------------------------------------
        # Define the index of the header row as None
        header_row_idx = None

        # Loop through the rows of the dataframe
        for idx in range(len(df)):
            # Extract first 3 cells, clean whitespace, and drop "nan" strings
            first_three = (
                df.iloc[idx, :3]
                .fillna("")
                .astype(str)
                .str.strip()
                .replace({"nan": "", "None": ""})
            )
            # Check if all 3 cells are non-empty
            if len(first_three) >= 3 and (first_three != "").all():
                header_row_idx = idx
                break

        # if the header row is not found
        if header_row_idx is None:
            print(f"Header row not found in {file.name} [{sheet}], skipping sheet.")
            continue

        # Extract headers starting from the identified header row
        # since there are double headers, for each column get the first header and the second header as separate variables
        # header1 and header2 store all the headers from the first header row - similarly for header2
        header1 = (
            df.iloc[header_row_idx]
            .fillna("")
            .astype(str)
            .str.strip()
            .replace({"nan": "", "None": ""})
        )
        header2 = (
            df.iloc[header_row_idx + 1]
            .fillna("")
            .astype(str)
            .str.strip()
            .replace({"nan": "", "None": ""})
        )

        # loop through the header rows
        # if the headers are the same, choose 1. if the headers are different, combine them. if one of the headers is empty, choose the other
        # add each header to the unified list new_headers
        new_headers = []
        for h1, h2 in zip(header1, header2):
            if h1 == h2:
                new_headers.append(h1)
            elif not h1:
                new_headers.append(h2)
            elif not h2:
                new_headers.append(h1)
            else:
                new_headers.append(f"{h1}\n{h2}")

        # Slice data to start after the two header rows
        df = df.iloc[header_row_idx + 2 :].reset_index(drop=True)
        # rename the headers of the dataframe
        df.columns = new_headers
        df.columns = [col.replace('*\n', ' ').replace('\n', ' ').strip() for col in df.columns]
        df = df.rename(columns=col_mapping)
        # add the dataframe to the list of dfs
        pbn_dfs.append(df)
        print(f"Uploading {file.name} [{sheet}] to database as {combined_file_sheet_mapping[sheet]}...")
        print(df.columns)
        # Upload the df to SQL
        df.to_sql(
            combined_file_sheet_mapping[sheet],
            con=engine,
            if_exists="replace",
            index=False,
            schema="pbn"
        )

# Combine the PBN dataframes into one
pbn_combined = pd.concat(pbn_dfs, ignore_index=True).drop_duplicates(ignore_index=True)

# upload the combined dataset to GCP
pbn_combined.to_sql(
    "pbn_combined",
    con=engine,
    if_exists="replace",
    index=False,
    schema="pbn"

)
# %%
