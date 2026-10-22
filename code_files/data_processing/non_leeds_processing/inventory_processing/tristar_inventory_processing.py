#!/usr/bin/env python

# In[1]: Import the Required Libraries
import logging
import re
from datetime import datetime

logger = logging.getLogger(__name__)

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

pd.set_option("display.max_columns", None)
pd.set_option("display.max_rows", None)


# In[2]: Connect to the TriStar Database
# host = "35.221.48.238"
# user = "appsmith_user"
# password = "Garfield/@123!"
# encoded_password = quote_plus(password)
# dbname = "tristar-database"

# conn = psycopg2.connect(host=host, user=user, password=password, dbname=dbname)
# cur = conn.cursor()
# engine = create_engine(f'postgresql+psycopg2://{user}:{encoded_password}@{host}:5432/{dbname}')


# In[3]: Import the TriStar Inventory
inventory = pd.read_excel(
    "C:/Users/medaamm/Box/Data file (AA)/TriStar Inventory/TriStar Inventory Mar 2023 (7)- MB_edits 092529 (111923).xlsx",
    sheet_name="FFPE+useless",
    header=1,
)


# In[4]: Get the unique values in each column
column_uniqueVals = {}
for column in inventory.columns:
    column_uniqueVals[column] = inventory[column].unique()

# In[5]: Process the Origin column
inventory["Origin"].unique()
inventory["Origin"] = inventory["Origin"].astype(str)
for index, row in inventory.iterrows():
    if str(row["Origin"]).startswith("T"):
        inventory.loc[index, "Original_ID"] = row["Origin"]
        inventory.loc[index, "Origin"] = "NR"

# Replace NA values with NR
inventory["Origin"] = inventory["Origin"].fillna("NR")

# Create the replacements dictionary for the Origin column
origin_replacements = {
    "Unknown": "NR",
    #'\u3000': 'NR',
    "-": "NR",
}

# Make replacements using the created dictionary
inventory["Origin"] = inventory["Origin"].replace(origin_replacements)


# In[6]: Process the Ethnicity column

# View the unique values in the Ethnicity column
inventory["Ethnicity"].unique()

# trim the Ethnicity column values and convert to title case
inventory["Ethnicity"] = inventory["Ethnicity"].str.strip().str.title()

# fill the NA values in the Ethnicity column
inventory["Ethnicity"] = inventory["Ethnicity"].fillna("NR")

# Create the replacements dictionary for the Ethnicity column
ethnicity_replacements = {
    "Not Stated": "NR",
    "Not Specified": "NR",
    "Not Sure About This Case": "NR",
    "/": "NR",
    "Unspecified": "NR",
    "Nd": "NR",
    "Black-Carribean": "Black Caribbean",
    "Black-Caribbean": "Black Caribbean",
    "Black Caribean": "Black Caribbean",
    "-": "NR",
    "Nc": "NR",
    "": "NR",
    "N/S": "NR",
}

# Make replacements in the Ethnicity column using the created dictionary
inventory["Ethnicity"] = inventory["Ethnicity"].replace(ethnicity_replacements)


# Make specific corrections to the NHSB entries
for index, row in inventory.iterrows():
    if row["Ethnicity"] == "Nhsb":
        inventory.loc[index, "Origin"] = "NHSB"
        inventory.loc[index, "Ethnicity"] = "NR"

# View the unique values again
inventory["Ethnicity"].unique()


# In[7]: Process the Gender column

# View the unique values under the Gender column
inventory["Gender"].unique()

# Loop through the rows and process the swapped values between the Age and Gender columns
for index, row in inventory.iterrows():
    if row["Gender"] in [59, 73, 60, 44, 48]:
        inventory.loc[index, ["Gender", "Age"]] = inventory.loc[
            index, ["Age", "Gender"]
        ].values

# trim the Gender column values and convert to title case
inventory["Gender"] = inventory["Gender"].str.title().str.strip()

# Fill the NA values under the Gender column
inventory["Gender"] = inventory["Gender"].fillna("NR")

# Create the replacements dictionary for the Gender column
gender_replacements = {
    "F": "Female",
    "M": "Male",
    "H": "Male",
    "N": "NR",
    "No Data Recorded": "NR",
    "N/A": "NR",
    "Nr": "NR",
    "": "NR",
    "/": "NR",
    "No Data Rec.": "NR",
    "Ctr File Data": "CTR File Data",
    "-": "NR",
    "MALE": "Male",
    "FEMALE": "Female",
}

# Make replacements in the Gender column using the created dictionary
inventory["Gender"] = inventory["Gender"].replace(gender_replacements)

# View the unique values again
inventory["Gender"].unique()

# Processing the Age Column
inventory["Age"].unique()

# Replace the blank values
inventory["Age"] = inventory["Age"].astype(str)
inventory["Age"] = inventory["Age"].str.strip()

# Fill the NA values
inventory["Age"] = inventory["Age"].fillna("NR")

# Create the replacements dictionary for the Age column
age_replacements = {
    "NR": "NR",
    "/": "NR",
    "": "NR",
    "Not communicated": "NR",
    "No Data Rec.": "NR",
    "-": "NR",
    "NC": "NR",
    "H": "Male",
    "M": "Male",
    "F": "Female",
    "CTR File DATA": "CTR File Data",
}

# Make replacements in the Age column using the created dictionary
inventory["Age"] = inventory["Age"].replace(age_replacements)

# View the unique values again
inventory.Age.unique()


# In[8]: Process the T Stage column

# View the unique values under the T Stage column
inventory["T"].unique()

# fill the NA values in the T Stage column
inventory["T"] = inventory["T"].fillna("NR")

# Create the replacements dictionary for the T Stage column
T_replacements = {
    "/": "NR",
    "-": "NR",
    "pT4a/b": "pT4a/pT4b",
    "Pt4": "pT4",
    "4C": "T4c",
    "3A": "T3a",
    "4A": "T4a",
    "3B": "T3b",
    "2C": "T2c",
    "2 (m)": "T2(m)",
    1: "T1",
    2: "T2",
    3: "T3",
    4: "T4",
    "2a": "T2a",
    "1a": "T1a",
    "2b": "T2b",
    "1b": "T1b",
    "NR": "NR",
    "3b": "T3b",
    "4a": "T4a",
    "1c": "T1c",
    "3a": "T3a",
    "2c": "T2c",
    "T3C": "T3c",
    "T3A": "T3a",
    "4": "T4",
    "4b": "T4b",
    "3c": "T3c",
    "IIB": "T2b",
    "III": "T3",
    "IC": "T1c",
    "IIA": "T2A",
    "2(m)": "T2m",
    "3a ": "T3a",
    "Not stated": "NR",
    "T2A": "T2a",
    "1b1": "T1b1",
    "4B": "T4b",
    "2A": "T2a",
    "2b ": "T2b",
    "1b2": "T1b2",
    "2a ": "T2a",
    "3b ": "T3b",
    "T1B": "T1b",
    "1b ": "T1b",
    "3bm": "T3bm",
    "3A ": "T3a",
    "2m": "T2m",
    "3m": "T3m",
    "1B": "T1b",
    "3a (m)": "T3a(m)",
    "a": "Ta",
    "pt3": "pT3",
    "?": "NR",
    "N/R": "NR",
    "T2m": "T2(m)",
    "1bm": "T1b(m)",
    "1C": "T1c",
    "T1C": "T1c",
    "T1A": "T1a",
    "2B": "T2b",
    "1A": "T1a",
    "3A": "T3",
    "3A(m)": "T3a(m)",
    "pt3a": "pT3a",
    "4(am)": "T4a(m)",
    "4am": "T4a(m)",
    "1(m)": "T1(m)",
    "Pt3": "pT3",
    "pt4": "pT4",
    "1A (Regressed)": "T1a (Regressed)",
    "1A (regressed)": "T1a (Regressed)",
    "x": "X",
    "3c1": "T3c1",
    "3C": "T3c",
    "4x": "T4X",
    "1s": "Tis",
    "c4": "cT4",
    "No Data Rec.": "NR",
    "No Data received": "NR",
    "N.A.": "NR",
    "N.A": "NR",
    "PT2": "pT2",
    "y2": "yT2",
    "y1cm": "yT1c(m)",
    "y1c": "yT1c",
    "y1bm": "yT1b(m)",
    "y4": "yT4",
    "y3": "yT3",
    "y1a(m)": "yT1a(m)",
    "y1": "yT1",
    "y4b": "yT4b",
    "1c m": "T1c(m)",
    "2 m": "T2(m)",
    "03월 04일": "NR",
    "21": "T2",
    "IB": "T1b",
    "is": "Tis",
    "yT2 (4th)": "yT2",
    "y1a": "yT1a",
    "y1b": "yT1b",
    "y1mi": "yT1mi",
    " ": "NR",
    "1m": "T1m",
    "2a (m)": "T2a(m)",
    "2n": "T2n",
    "no TNM (7cm mass with ovary involvement)": "NR",
    "\u3000": "NR",
    "T2A": "T2a",
    "mri 3b": "T3b",
    21: "NR",
    "K": "NR",
}

# Make replacements in the T Stage column using the created dictionary
inventory["T"] = inventory["T"].replace(T_replacements)

# View the unique values again
inventory["T"].unique()


# In[9]: Process the N Stage column

# View the unique values under the N Stage column
inventory["N"].unique()

# Loop through the rows and process the swapped values between the N and M columns
for index, row in inventory.iterrows():
    if row["N"] in ["pM1"]:
        inventory.loc[index, ["N", "M"]] = inventory.loc[index, ["M", "N"]].values

# Create the replacements dictionary for the N Stage column
N_replacements = {
    0: "N0",
    1: "N1",
    2: "N2",
    3: "N3",
    "T2": "N2",
    "M0": "N0",
    "NR": "NR",
    "x": "X",
    "/": "NR",
    "3a": "N3a",
    "2b": "N2b",
    "2c": "N2c",
    "1": "N1",
    "1c": "N1c",
    "1b": "N1b",
    "2a": "N2a",
    "1a": "N1a",
    "3b": "N3b",
    "II": "N2",
    "I": "N1",
    "IIB": "N2b",
    "IB": "N1b",
    "IA": "N1a",
    "2a1": "N2a1",
    "N2c": "N2c",
    "N/R": "NR",
    "1B": "N1b",
    "1s": "N1s",
    "3B": "N3b",
    "2B": "N2b",
    "3C": "N3c",
    "2A": "N2a",
    "2C": "N2c",
    "ND": "NR",
    "2b(mi)": "N2b(mi)",
    "NR ": "NR",
    "N3A": "N3a",
    "N1A": "N1a",
    "1ami": "N1a(mi)",
    "?": "NR",
    "NR (station 11 nodes positive)": "Not Recorded (Station 11 Nodes Positive)",
    "1A": "N1a",
    "Pn1": "pN1",
    "N2C": "N2c",
    "Pn0": "pN0",
    "pN1 p": "pN1",
    "1C": "N1c",
    "CN0": "cN0",
    "3c": "N3c",
    "pN2 p": "pN2",
    "pN0 p": "pN0",
    "N2A": "N2a",
    "yTN1a": "N1",
    "NX": "X",
    "YpN0": "ypN0",
    4: "N4",
    "No Data Rec.": "NR",
    "No Data received": "NR",
    "N.A.": "NR",
    "1mi": "N1(mi)",
    "0(sn)": "N0(sn)",
    "CTR File DATA": "CTR File Data",
    "N.A": "NR",
    "Not stated": "NR",
    "0sn": "N0(sn)",
    "2d": "N2d",
    "yT2b": "yN2b",
    "-": "NR",
    "Nx": "X",
    "pT4": "NR",
    "Not Recorded (Station 11 Nodes Positive)": "NR",
}

# Make replacements in the N Stage column using the created dictionary
inventory["N"] = inventory["N"].replace(N_replacements)

# Fill the NA values in the N Stage column
inventory["N"] = inventory["N"].fillna("NR")

# View the unique values again
inventory.N.unique()


# In[10]: Process the M Stage column

# View the unique values under the M Stage column
inventory["M"].unique()

# Fill the NA values in the M Stage column
inventory["M"] = inventory["M"].fillna("NR")

# Create the replacements dictionary for the M Stage column
M_replacements = {
    0: "M0",
    "NR": "NR",
    "x": "X",
    "/": "NR",
    1: "M1",
    "-": "NR",
    "1b": "M1b",
    "0": "M0",
    "R0": "rM0",
    "1c": "M1c",
    "1a": "M1a",
    "1B": "M1b",
    "A": "M1a",
    "Mx": "X",
    "1 (Lung)": "M1 (Lung)",
    "1 (Omentum)": "M1 (Omentum)",
    "1A": "M1a",
    "1 (omentum)": "M1 (Omentum)",
    "1 (Left common iliac)": "M1 (Left Common Iliac)",
    "NR ": "NR",
    "1C": "M1c",
    "MI": "M1",
    "1 (Adrenal gland)": "M1 (Adrenal Gland)",
    "omentum": "Not Record",
    "?": "NR",
    "1 (Left adrenal)": "N1 (Left Adrenal)",
    "B": "N1b",
    "pMx pR1": "X",
    "pR0": "X",
    "1 (brain)": "M1 (Brain)",
    "1 (lung)": "M1 (Lung)",
    "1 (Bone)": "M1 (Bone)",
    "pR1": "X",
    "R1": "X",
    "R2": "X",
    "RX": "X",
    "pRx": "X",
    "1A (Inguinal and Para-Aortic Lymph Nodes)": "M1a (Inguinal and Para-Aortic Lymph Nodes)",
    "1C (right ovary, peritoneum, bladder)": "M1c (Right Ovary, Peritoneum, Bladder)",
    "1C (liver and peritoneum)": "M1c (Liver and Peritoneum)",
    "1a (pleural effusion only)": "M1a (Pleural Effusion Only)",
    "1b (Spine)": "M1b (Spine)",
    "1c (Omentum and para colic cutter)": "M1c (Omentum and Para Coliuc Cutter)",
    "1 (intralobar)": "M1 (Intralobar)",
    "1d": "M1d",
    "la": "M1a",
    "1a (liver)": "M1a (Liver)",
    "1c (peritoneu)": "M1c (Peritoneum)",
    3: "M3",
    2: "M2",
    "No Data Rec.": "NR",
    "No Data received": "NR",
    " x": "X",
    "N.A.": "NR",
    "CTR File DATA": "CTR File Data",
    "N.A": "NR",
    "Not stated": "NR",
    "Not Record": "NR",
}

# Make replacements in the M Stage column using the created dictionary
inventory["M"] = inventory["M"].replace(M_replacements)

# View the unique values again
inventory.M.unique()

# In[11]: Process the Stage column

# View the unique values under the Stage column
inventory["Stage"].unique()

# Loop through the rows and process the swapped values between the Stage and Grade columns
for index, row in inventory.iterrows():
    if row["Stage"] in ["Grade 1", "G1", "G2", "G3", "G4"]:
        inventory.loc[index, ["Stage", "Grade"]] = inventory.loc[
            index, ["Grade", "Stage"]
        ].values
for index, row in inventory.iterrows():
    if row["Stage"] in ["T4 N1 R0", "T4 N1"]:
        inventory.loc[index, "T"] = "T4"
        inventory.loc[index, "N"] = "N1"
        inventory.loc[index, "M"] = "M0"
        row["Stage"] = "NR"
    elif row["Stage"] == "T3 N0 R0":
        inventory.loc[index, "T"] = "T3"
        inventory.loc[index, "N"] = "N0"
        inventory.loc[index, "M"] = "M0"
    elif row["Stage"] == "pT3":
        inventory.loc[index, "T"] = "pT3"
        inventory.loc[index, "N"] = "N0"
        inventory.loc[index, "M"] = "M0"
    elif row["Stage"] == "PT2":
        inventory.loc[index, "T"] = "pT2"
        inventory.loc[index, "N"] = "N0"
        inventory.loc[index, "M"] = "M0"
    elif row["Stage"] == "T2 N1 R0":
        inventory.loc[index, "T"] = "pT2"
        inventory.loc[index, "N"] = "N1"
        inventory.loc[index, "M"] = "M0"


# Fill the NA values in the Stage column
inventory["Stage"] = inventory["Stage"].fillna("NR")


# Create the replacements dictionary for the Stage column
stage_replacements = {
    "NR": "NR",
    "Not stated": "NR",
    "/": "NR",
    "-": "NR",
    "IV-A": "IVA",
    "3c": "IIIC",
    "1c": "IC",
    3: "III",
    "3b": "IIIb",
    4: "IV",
    "2b": "IIB",
    "3c ": "IIIc",
    "1c3": "III",
    "1a": "IA",
    "3a": "IIIA",
    2: "II",
    "1c2": "II",
    "2a": "IIA",
    "2c": "IIC",
    "4b": "IVB",
    "No information": "NR",
    "1b": "IB",
    "4a": "IVA",
    "2b ": "IIB",
    "3b ": "IIIB",
    "2B": "IIB",
    "4a ": "IVA",
    "4b ": "IVB",
    "3B": "IIIB",
    "IVA/IVB (AJCC 7th edition)": "IVA/IVB (AJCC 7)",
    "IB  (AJCC7)": "IB (AJCC 7)",
    "IA (AJCC7)": "IA (AJCC 7)",
    "TBC": "NR",
    "IIIC (AJCC7)": "IIIC (AJCC 7)",
    "IIIA (AJCC7)": "IIIA (AJCC 7)",
    "IB (AJCC7)": "IB (AJCC 7)",
    "IIB (AJCC7)": "IIB (AJCC 7)",
    "IIIB (AJCC7)": "IIIB (AJCC 7)",
    "I (AJCC8; post neoadjuvant)": "I (AJCC 8 Post Neoadjuvant)",
    "IIB (AJCC8)": "IIB (AJCC 8)",
    "Clarke level 5": "Clarke Level 5",
    1: "I",
    "IIIB (AJCC8 post neoadjuvant)": "IIIB (AJCC 8 Post Neoadjuvant)",
    "Masoka stage IVb": "Masoka Stage IVB",
    "II (AJCC8 post neoadjuvant)": "II (AJCC 8 Post Neoadjuvant)",
    "IC (AJCC8)": "IC (AJCC 8)",
    "IB (AJCC8)": "IB (AJCC 8)",
    "IVA (AJCC post neoadjuvant)": "IVA (AJCC Post Neoadjuvant)",
    "IVB (AJCC8 post neoadjuvant)": "IVB (AJCC Post Neoadjuvant)",
    "IIIA (AJCC8 post neoadjuvant)": "IIIA (AJCC 8 Post Neoadjuvant)",
    "?": "NR",
    "Cannot be determined": "NR",
    "2A (FIGO)": "IIA (FIGO)",
    "II (presumed)": "II (Presumed)",
    "IB (provisional; pt. declined SLNB)": "IB (Provisional; pt. declined SLNB)",
    0: "0",
    "III (provisional)": "III (Provisional)",
    "IB (presumed)": "IB (Presumed)",
    "IIA (presumed)": "IIA (Presumed)",
    "II (bladder cancer)": "II (Bladder Cancer)",
    "IIa": "IIA",
    "IIIB (7th)": "IIIB (AJCC 7)",
    "IB - IIA": "IB/IIA",
    "Not Stated": "NR",
    "IIA - IIB": "IIA/IIB",
    "IIIC ": "IIIC",
    "IIB\xa0": "IIB",
    "NK": "NR",
    "Not Stated ": "NR",
    "Not Stated\xa0": "NR",
    "No Data Rec.": "NR",
    "No Data received": "NR",
    "Not Cited": "NR",
    "Not cited": "NR",
    "FIGO 1B": "FIGO IB",
    "1B": "IB",
    "1A": "IA",
    "1C": "IC",
    "2A": "IIA",
    "3A": "IIIA",
    "at least III": "III (at least)",
    "CTR File DATA": "CTR File Data",
    "N.A": "NR",
    "at least III A": "IIIA (at least)",
    "III B": "IIIB",
    "at least II": "II (at least)",
    "III A": "IIIA",
    "2 o 3a": "II/IIIA",
    "unknown": "NR",
    "IIb": "IIB",
    "IIIc": "IIIC",
}

# Make replacements in the Stage column using the created dictionary
inventory["Stage"] = inventory["Stage"].replace(stage_replacements)

# In[12]: Process the Date of Diagnosis column

# Create the replacements dictionary for the Date of Diagnosis column
date_diagnosis_replacements = {
    "No information": "NR",
    "collected: 2011": "2011",
    "NR": "NR",
    "primary": "NR",
}

# Make replacements in the Date of Diagnosis column using the created dictionary
inventory["Date_Diagnosis"] = inventory["Date_Diagnosis"].replace(
    date_diagnosis_replacements
)

# Fill the NA values in the Date of Diagnosis column
inventory["Date_Diagnosis"] = inventory["Date_Diagnosis"].fillna("NR")

# Convert the Date of Diagnosis column to datetime format
# inventory['Date_Diagnosis'] = pd.to_datetime(inventory['Date_Diagnosis'].apply(lambda x: f"{x}-01-01" if isinstance(x, (int, str)) and str(x).isdigit() else x), errors='coerce')


# Create a function to convert date to Excel serial number
def date_to_excel_serial(date):
    excel_start = datetime(1899, 12, 30)
    delta = date - excel_start
    return delta.days


# Create the Year_Diagnosis and Month_Diagnosis columns
for i, date in enumerate(inventory["Date_Diagnosis"]):
    if isinstance(date, datetime) and date.year == 1905:
        inventory.loc[i, "Year_Diagnosis"] = date_to_excel_serial(date)
        inventory.loc[i, "Month_Diagnosis"] = 1
    elif isinstance(date, datetime) and date.year != 1905:
        inventory.loc[i, "Year_Diagnosis"] = date.year
        inventory.loc[i, "Month_Diagnosis"] = date.month


# Convert the Year_Diagnosis and Month_Diagnosis columns to numeric
inventory["Year_Diagnosis"] = pd.to_numeric(
    inventory["Year_Diagnosis"], errors="coerce"
).astype("Int64")
inventory["Month_Diagnosis"] = pd.to_numeric(
    inventory["Month_Diagnosis"], errors="coerce"
).astype("Int64")


# In[13]: Process the Others column
for i, val in enumerate(inventory["Others"]):
    if isinstance(val, str) and val.startswith("Year of Collection"):
        val = val.replace("Year of Collection: ", "")
        if len(val) == 6:
            val = "20" + val[-2:]
            inventory.loc[i, "Others"] = int(val)
            inventory.loc[i, ["Others", "Date_Diagnosis"]] = inventory.loc[
                i, ["Date_Diagnosis", "Others"]
            ].values
        else:
            val = val[:4]
            inventory.loc[i, "Others"] = int(val)
            inventory.loc[i, ["Others", "Date_Diagnosis"]] = inventory.loc[
                i, ["Date_Diagnosis", "Others"]
            ].values
    if isinstance(val, str) and val.startswith("Date of Surgery"):
        val = val.replace("Date of Surgery : ", "")
        inventory.loc[i, "Others"] = int(val)
        inventory.loc[i, ["Others", "Date_Diagnosis"]] = inventory.loc[
            i, ["Date_Diagnosis", "Others"]
        ].values

    if isinstance(val, str):
        if val.startswith("ypT3"):
            inventory.loc[i, "Others"] = ""
            inventory.loc[i, "T"] = "ypT3"

        elif val.startswith("ypT2"):
            inventory.loc[i, "Others"] = ""
            inventory.loc[i, "T"] = "ypT2"

# In[14]: Prepare the Diagnosis and Subtype Columns

# Fill the NA values in both columns first
inventory["Sub_Type"] = inventory["Sub_Type"].fillna("-")
inventory["Specimen_Histological_Diagnosis"] = inventory[
    "Specimen_Histological_Diagnosis"
].fillna("-")


# Trim extra spaces and convert to title case
inventory["Specimen_Histological_Diagnosis"] = (
    inventory["Specimen_Histological_Diagnosis"].str.strip().str.title()
)
inventory["Sub_Type"] = inventory["Sub_Type"].str.strip().str.title()


# In[15]: Add the Cancer Type Column

# Lung
keywords_lung = [
    "Nsclc",
    "Lung",
    "Left Lower Lobe",
    "Left Upper Lobe",
    "Right Upper Lobe",
    "Right Lower Lobe",
    "Right Middle Lobe",
    "Right Middle And Lower Lobe",
    "Lower Left Lobe",
    "Sclc",
    "Right Upper And Middle Lobe",
    "Bronchus",
    "Cdi",
    "Wedge Of Upper Lobe & Node",
]
# Pleura
keywords_pleura = ["Parietal Pleura", "Visceral Pleura", "Pleura"]

# Prostate
keywords_prostate = ["Prostate", "Prostatic"]

# stomach
keywords_gastric = ["Gastric", "Gastrectomy", "Stomach"]

# breast
keywords_breast = ["Breast", "Tnbc"]


# intestines
keywords_intestines = [
    "Crc",
    "Colorectal",
    "Rectum",
    "Sigmoid",
    "Caecum",
    "Cecum",
    "Hepatic Flexure",
    "Duodenal",
    "Duodenum",
    "Ileum",
    "Ileum",
    "Ileum & Caecum",
    "Bowel",
    "Intestine",
    "Intestil",
    "Intestinal",
    "Crohn'S Disease",
    "Appendix",
    "Ulcerative Colitis",
    "Colon",
    "Gist",
]

# liver
keywords_liver = [
    "Hcc",
    "Hepatocellullar",
    "Hepatocellular",
    "Liver",
    "Hepatocarcinoma",
]

# pancreas
keywords_pancreatic = ["Pancreatic", "Pancreas", "Pnet Primary", "Pnet"]

# kidney
keywords_kidney = ["Kidney", "Renal", "Rcc"]

# head and neck
keywords_HN = [
    "Head and Neck",
    "Head & Neck",
    "Gbm",
    "Brain Tissue",
    "Brain",
    "H&N",
    "Cerebellum",
    "Cerebrum",
    "Tongue",
    "Ear",
    "Mouth",
    "Cheek",
    "Head&Neck",
    "Head And Neck",
    "Neck",
    "Nose",
    "Hypopharynx",
    "Larynx",
    "Hen",
    "Laryngopharynx",
    "Nasal Cavity",
    "Forehead",
    "Gum",
    "Chin",
    "Oropharynx",
    "Parotid",
    "Cerebral Cortex",
    "Basal Cell Carcinoma - Vertex Of Scalp",
    "Basal Cell Carcinoma - Scalp",
    "Scalp",
    "Basal Cell Carcinoma - Right Post Auricular",
    "Basal Cell Carcinoma - Superior Temple",
    "Thyroid",
    "Glioblastoma Cancer",
    "Lip",
    "Tonsil",
    "Palate",
    "Vocal Cord",
    "Mandible",
    "Maxilla",
    "Jaw",
    "Glossectomy",
    "Hemiglossectomy",
    "Bucca",
    "Orpopharynx",
    "Oropharyngectomy",
    "Temple",
    "Temporal",
    "Nasopharynx",
    "Npc",
    "Occiput",
    "Salivary Gland",
    "Auricular",
    "Neck, Lymph Node & Partoid Gland",
    "Neck, Level 2 Node, Superficial Right Parotid",
]

# eye
keywords_eye = ["Choroidal", "Uveal Melanoma"]


# oesophagus
keywords_oesophagal = [
    "Oesophagus",
    "Oesophagal",
    "Esophagus",
    "Esophagal",
    "Esophageal",
    "Oesophageal",
]

# female reproductive system
keywords_femalereprodsystem = [
    "Ovarian",
    "Ovary",
    "Ovaries",
    "Fallopian Tube",
    "Salpinx",
    "Papillary Serous Carcinoma",
    "uterus",
    "Uterus",
    "Myometrium",
    "Placenta",
    "Endometrium",
    "Endometrial",
    "Endometrioid",
    "Cervical",
    "Cervix",
    "cervix",
]

# adrenal gland
keywords_adrenal = ["Adrenal", "Adrel"]

# bile duct
keywords_bileduct = ["Bile Duct", "Ductal Cholangiocarcinoma"]

# gallbladder
keywords_gallbladder = ["Galbladder", "Gallbladder", "Gall Bladder"]


# heart
keywords_heart = ["Heart"]

# omentum
keywords_omentum = ["Omentum"]

# GOJ
keywords_goj = ["Gej", "Goj"]

# shoulder
keywords_shoulder = ["Shoulder", "Subscapular"]

# pelvis
keywords_pelvis = ["Pelvis", "Pelvic", "Seminal Vesicle", "Semil Vesicle"]

# buttock
keywords_buttock = ["Buttock"]

# lower body
keywords_lowerbody = [
    "Leg",
    "Knee",
    "Foot",
    "Feet",
    "Thigh",
    "Femur",
    "Fibula",
    "Lower Limb",
    "Right Flank, Ilio-Inguinal Lymph Node, Skin",
]

# chest
keywords_chest = ["Chest", "Mediastinal", "Mediastinum", "Rib", "Thymus"]

# back
keywords_back = ["Back", "Spine", "L4 Vertebrae", "T7 Vertebrae"]

# ampulla
keywords_ampulla = ["Ampulla"]

# arms
keywords_arms = [
    "Arm",
    "Elbow",
    "Hand",
    "Forearm",
    "Axilla",
    "Humerus",
    "Triceps",
    "Wrist",
    "Thumb",
]

# male reproductive system
keywords_malereprod = ["Penis", "Testis"]

# anal canal
keywords_anal = ["Anal Canal"]

# skin
keywords_skin = [
    "Cutaneous Melanoma",
    "Cutaneus Melanoma",
    "Merkel Cell Carcinoma",
    "Excision Of Skin Lesion",
    "Skin",
    "Melanoma",
    "Cscc",
    "Cutaneous",
    "Cute",
    "Right Nasojugal Fold",
]

# Lymphatic system
keywords_lymphatic = [
    "Lymphoma",
    "Bone Marrow",
    "Nhl",
    "Ln",
    "Lymph Node",
    "Node",
    "Nodal",
    "Spleen",
    "Dlbcl",
    "Ptcl, Nos",
    "Mzl",
    "Speen",
    "Short Level 2A Long Level 4",
    "Sll",
    "Hen Nodal Met.",
    "Neck: Left Level 5A",
]

# peritoneum
keywords_peritoneum = ["Peritoneum"]

# digestive tract
keywords_digestivetract = ["Gist"]

# urinary system
keywords_urinary = ["Ureter", "Upper Urotelium Tract", "Bladder"]

# asbcess
keywords_abscess = ["Abscess"]

# mesothelium
keywords_mesothelium = ["Mesothelioma"]

# vessels
keywords_vessels = ["Angiosarcoma", "Vascular Invasion"]

# cortex
keywords_cortex = ["Cortex"]

# Smooth muscle
keywords_smoothmuscle = ["Leiomyo Sarcoma", "Leiomyosarcoma"]

# muscle
keywords_muscle = ["Muscle", "Skeletal"]

# retroperitoneum
keywords_retroperitoneum = ["Retroperitoneum", "Retroperitoneal"]

# connective tissue
keywords_connectivetissue = ["Undifferentiated Sarcoma", "Fibrosarcoma"]

# multiple sites
keywords_multiple = ["Adrenal, Thymus & Bladder"]

# plasma
keywords_plasma = ["Plasma"]


for i, (diagnosis, subtype) in enumerate(
    zip(inventory["Specimen_Histological_Diagnosis"], inventory["Sub_Type"])
):
    if (
        any(keyword in diagnosis for keyword in keywords_lung)
        and "Lung Nodal Met." not in diagnosis
        and subtype != "Ln"
    ):
        inventory.loc[i, "Tissue Source"] = "Lung"
    if (
        any(keyword in diagnosis for keyword in keywords_prostate)
        and "Plasma" not in diagnosis
    ):
        inventory.loc[i, "Tissue Source"] = "Prostate"
    if (
        any(keyword in diagnosis for keyword in keywords_gastric)
        and "Gist" not in diagnosis
        and "Gist" not in subtype
    ):
        inventory.loc[i, "Tissue Source"] = "Stomach"
    if any(keyword in diagnosis for keyword in keywords_breast):
        inventory.loc[i, "Tissue Source"] = "Breast"
    if any(keyword in diagnosis for keyword in keywords_liver) and all(
        keyword not in diagnosis for keyword in ["Crc", "Colon"]
    ):
        inventory.loc[i, "Tissue Source"] = "Liver"
    if any(keyword in diagnosis for keyword in keywords_pancreatic):
        inventory.loc[i, "Tissue Source"] = "Pancreas"
    if any(keyword in diagnosis for keyword in keywords_kidney):
        inventory.loc[i, "Tissue Source"] = "Kidney"
    if (
        any(keyword in diagnosis for keyword in keywords_HN)
        and all(
            keyword not in diagnosis
            for keyword in ["Nodal", "Node", "Nodes", "Liposarcoma", "Sarcoma"]
        )
        and all(
            keyword not in subtype
            for keyword in [
                "Lower limb and hip",
                "Dorsum lefrt hand",
                "Dorsum lefrt hand",
            ]
        )
    ):
        inventory.loc[i, "Tissue Source"] = "Head & Neck"
    if (
        any(keyword in diagnosis for keyword in keywords_intestines)
        and all(
            keyword not in diagnosis
            for keyword in [
                "Plasma",
                "Galbladder",
                "Galbladder Pri.",
                "Gastrectomy",
                "Gastric Adenocarcinomas",
                "Gastric Pri.",
                "Gastric Pri. And Nodal Met.",
                "Lung Pri.",
                "Lymph Node",
                "Nhl",
                "Nodal Met.",
                "Ovary",
                "Rectum Adenocarcinoma Lung Metastase",
                "Representive Examples Of Serosa Nodules",
                "Sigmoid Adenocarcinoma Lung Metastase",
                "Tumor?",
                "Vascular Invasion",
                "CRC Nodal Met",
                "Lymph Node",
                "Nodal Met.",
                "Crc Nodal Met",
            ]
        )
        and "CRC_Matched Met" not in subtype
    ):
        inventory.loc[i, "Tissue Source"] = "Intestines"
    if any(keyword in diagnosis for keyword in keywords_oesophagal):
        inventory.loc[i, "Tissue Source"] = "Oesophagus"
    if (
        any(keyword in diagnosis for keyword in keywords_femalereprodsystem)
        and all(keyword not in diagnosis for keyword in ["Plasma", "Node"])
        and "Node" not in subtype
    ):
        inventory.loc[i, "Tissue Source"] = "Female Reproductive System"
    if any(keyword in diagnosis for keyword in keywords_adrenal):
        inventory.loc[i, "Tissue Source"] = "Adrenal Gland"
    if any(keyword in diagnosis for keyword in keywords_bileduct):
        inventory.loc[i, "Tissue Source"] = "Bile Duct"
    if any(keyword in diagnosis for keyword in keywords_gallbladder):
        inventory.loc[i, "Tissue Source"] = "Gallbladder"
    if any(keyword in diagnosis for keyword in keywords_heart):
        inventory.loc[i, "Tissue Source"] = "Heart"
    if any(keyword in diagnosis for keyword in keywords_omentum):
        inventory.loc[i, "Tissue Source"] = "Omentum"
    if any(keyword in diagnosis for keyword in keywords_goj):
        inventory.loc[i, "Tissue Source"] = "GOJ"
    if any(keyword in diagnosis for keyword in keywords_shoulder):
        inventory.loc[i, "Tissue Source"] = "Shoulder"
    if any(keyword in diagnosis for keyword in keywords_pelvis):
        inventory.loc[i, "Tissue Source"] = "Pelvis"
    if any(keyword in diagnosis for keyword in keywords_lowerbody):
        inventory.loc[i, "Tissue Source"] = "Lower Body"
    if any(keyword in diagnosis for keyword in keywords_chest):
        inventory.loc[i, "Tissue Source"] = "Chest"
    if any(keyword in diagnosis for keyword in keywords_back):
        inventory.loc[i, "Tissue Source"] = "Back"
    if any(keyword in diagnosis for keyword in keywords_ampulla):
        inventory.loc[i, "Tissue Source"] = "Ampulla"
    if any(keyword in diagnosis for keyword in keywords_arms):
        inventory.loc[i, "Tissue Source"] = "Arms"
    if any(keyword in diagnosis for keyword in keywords_eye):
        inventory.loc[i, "Tissue Source"] = "Eye"
    if any(keyword in diagnosis for keyword in keywords_anal):
        inventory.loc[i, "Tissue Source"] = "Anal Canal"
    if any(keyword in diagnosis for keyword in keywords_malereprod):
        inventory.loc[i, "Tissue Source"] = "Male Reproductive System"
    if any(keyword in diagnosis for keyword in keywords_skin) and all(
        keyword not in diagnosis for keyword in ["Node", "Nodes", "Level 5A"]
    ):
        inventory.loc[i, "Tissue Source"] = "Skin"
    if (
        any(keyword in diagnosis for keyword in keywords_lymphatic)
        and all(
            keyword not in diagnosis
            for keyword in [
                "Gastric",
                "Stomach",
                "Crc",
                "Lung + Nodes",
                "Lung Cancer",
                "Nsclc",
                "Breast Pri. And Nodal Met.",
                "Kidney Pri. And Nodal Met.",
                "Lung Pri. And Nodal Met.",
                "Lung Nodal Met.",
                "Lung Pri.",
                "Pancreas Pri. And Nodal Met.",
                "Pancreatic Pri. And Nodal Met.",
                "Hen Nodal Met.",
                "Hen Pri. And Nodal Met.",
                "Pancreas Pri. And Nodal Met.",
                "Pancreatic Pri. And Nodal Met.",
                "Pancreatic Nodal Met.",
                "Pharynx Cancer With Ln Neck Metastasis",
                "Esophageal Adc",
            ]
        )
        and subtype == "Ln"
    ):
        inventory.loc[i, "Tissue Source"] = "Lymphatic System"
    if any(keyword in diagnosis for keyword in keywords_peritoneum):
        inventory.loc[i, "Tissue Source"] = "Peritoneum"
    if any(keyword in diagnosis for keyword in keywords_digestivetract):
        inventory.loc[i, "Tissue Source"] = "Digestive Tract"
    if (
        any(keyword in diagnosis for keyword in keywords_urinary)
        and "Gall Bladder" not in diagnosis
    ):
        inventory.loc[i, "Tissue Source"] = "Urinary System"
    if any(keyword in diagnosis for keyword in keywords_abscess):
        inventory.loc[i, "Tissue Source"] = "Abscess"
    if any(keyword in diagnosis for keyword in keywords_mesothelium):
        inventory.loc[i, "Tissue Source"] = "Mesothelium"
    if any(keyword in diagnosis for keyword in keywords_buttock):
        inventory.loc[i, "Tissue Source"] = "Buttock"
    if any(keyword in diagnosis for keyword in keywords_cortex):
        inventory.loc[i, "Tissue Source"] = "Cortex"
    if diagnosis == "Fl":
        inventory.loc[i, "Tissue Source"] = "Lymphatic System"
    if any(keyword in diagnosis for keyword in keywords_vessels):
        inventory.loc[i, "Tissue Source"] = "Lymphatic/Blood Vessels"
    if any(keyword in diagnosis for keyword in keywords_smoothmuscle):
        inventory.loc[i, "Tissue Source"] = "Smooth Muscle Tissue"
    if any(keyword in diagnosis for keyword in keywords_muscle):
        inventory.loc[i, "Tissue Source"] = "Muscle"
    if any(keyword in diagnosis for keyword in keywords_pleura):
        inventory.loc[i, "Tissue Source"] = "Pleura"
    if any(keyword in diagnosis for keyword in keywords_retroperitoneum):
        inventory.loc[i, "Tissue Source"] = "Retroperitoneum"
    if any(keyword in diagnosis for keyword in keywords_connectivetissue):
        inventory.loc[i, "Tissue Source"] = "Connective Tissue"
    if any(keyword in diagnosis for keyword in keywords_multiple):
        inventory.loc[i, "Tissue Source"] = "Multiple Sites"
    if any(keyword in diagnosis for keyword in keywords_plasma):
        inventory.loc[i, "Tissue Source"] = "Plasma"


for i, (diagnosis, subtype) in enumerate(
    zip(inventory["Specimen_Histological_Diagnosis"], inventory["Sub_Type"])
):
    if any(keyword in subtype for keyword in keywords_skin) and all(
        keyword not in diagnosis for keyword in ["Node", "Nodes", "Level 5A"]
    ):
        inventory.loc[i, "Tissue Source"] = "Skin"
    if (
        any(keyword in subtype for keyword in keywords_femalereprodsystem)
        and all(keyword not in diagnosis for keyword in ["Plasma", "Node"])
        and "Node" not in subtype
    ):
        inventory.loc[i, "Tissue Source"] = "Female Reproductive System"
    if (
        any(keyword in subtype for keyword in keywords_prostate)
        and "Plasma" not in diagnosis
    ):
        inventory.loc[i, "Tissue Source"] = "Prostate"
    if (
        (any(keyword in subtype for keyword in keywords_intestines))
        and all(
            keyword not in diagnosis
            for keyword in [
                "Plasma",
                "Galbladder",
                "Galbladder Pri.",
                "Gastrectomy",
                "Gastric Adenocarcinomas",
                "Gastric Pri.",
                "Gastric Pri. And Nodal Met.",
                "Lung Pri.",
                "Lymph Node",
                "Nhl",
                "Nodal Met.",
                "Ovary",
                "Rectum Adenocarcinoma Lung Metastase",
                "Representive Examples Of Serosa Nodules",
                "Sigmoid Adenocarcinoma Lung Metastase",
                "Tumor?",
                "Vascular Invasion",
                "CRC Nodal Met",
                "Lymph Node",
                "Nodal Met.",
                "Crc Nodal Met",
            ]
        )
        and all(keyword not in subtype for keyword in ["CRC_Matched Met"])
    ):
        inventory.loc[i, "Tissue Source"] = "Intestines"
    if (
        any(keyword in subtype for keyword in keywords_HN)
        and all(
            keyword not in diagnosis
            for keyword in ["Nodal", "Node", "Nodes", "Liposarcoma", "Sarcoma"]
        )
        and all(
            keyword not in subtype
            for keyword in [
                "Lower limb and hip",
                "Dorsum lefrt hand",
                "Dorsum lefrt hand",
            ]
        )
    ):
        inventory.loc[i, "Tissue Source"] = "Head & Neck"
    if any(keyword in subtype for keyword in keywords_liver) and all(
        keyword not in diagnosis for keyword in ["Crc", "Colon"]
    ):
        inventory.loc[i, "Tissue Source"] = "Liver"
    if any(keyword in subtype for keyword in keywords_breast):
        inventory.loc[i, "Tissue Source"] = "Breast"
    if (
        any(keyword in subtype for keyword in keywords_lymphatic)
        and all(
            keyword not in diagnosis
            for keyword in [
                "Gastric",
                "Stomach",
                "Crc",
                "Lung + Nodes",
                "Lung Cancer",
                "Nsclc",
                "Breast Pri. And Nodal Met.",
                "Kidney Pri. And Nodal Met.",
                "Lung Pri. And Nodal Met.",
                "Lung Nodal Met.",
                "Lung Pri.",
                "Pancreas Pri. And Nodal Met.",
                "Pancreatic Pri. And Nodal Met.",
                "Hen Nodal Met.",
                "Hen Pri. And Nodal Met.",
                "Pancreas Pri. And Nodal Met.",
                "Pancreatic Pri. And Nodal Met.",
                "Pancreatic Nodal Met.",
                "Pharynx Cancer With Ln Neck Metastasis",
                "Esophageal Adc",
            ]
        )
        and subtype == "Ln"
    ):
        inventory.loc[i, "Tissue Source"] = "Lymphatic System"
    if any(keyword in subtype for keyword in keywords_smoothmuscle):
        inventory.loc[i, "Tissue Source"] = "Smooth Muscle Tissue"
    if any(keyword in subtype for keyword in keywords_arms):
        inventory.loc[i, "Tissue Source"] = "Arms"
    if any(keyword in subtype for keyword in keywords_lowerbody):
        inventory.loc[i, "Tissue Source"] = "Lower Body"


for i, (diagnosis, subtype) in enumerate(
    zip(inventory["Specimen_Histological_Diagnosis"], inventory["Sub_Type"])
):
    if diagnosis == "-" and "Basaloid Squamous Cell Carcinoma" in subtype:
        inventory.loc[i, "Tissue Source"] = "NR"
    if diagnosis == "Sarcoma" and subtype == "Myxoid Liposarcoma":
        inventory.loc[i, "Tissue Source"] = "NR"


for i, (diagnosis, subtype) in enumerate(
    zip(inventory["Specimen_Histological_Diagnosis"], inventory["Sub_Type"])
):
    if (diagnosis == "-" or diagnosis == "" or diagnosis == "Nr") and (
        subtype == "-" or subtype == "" or subtype == "Nr"
    ) or (diagnosis == "Nr" or diagnosis == "-" or diagnosis == "") and (
        subtype == "Normal" or subtype == "No Cancer"
    ) or diagnosis == "Nr" and subtype == "Nr" or diagnosis == "Nr" and subtype == "No Cancer" or diagnosis == "-" and subtype == "Basaloid Squamous Cell Carcinoma" or diagnosis == "Liposarcoma" and subtype in ["Dedifferentiated", "Myxoid"]:
        inventory.loc[i, "Tissue Source"] = "NR"


# Correct the below
inventory.loc[
    (inventory["Specimen_Histological_Diagnosis"] == "Lung, Right Lower Lobe")
    & (inventory["Sub_Type"] == "Squamous Cell Carcinoma (Tonsil)"),
    "Tissue Source",
] = "Lung"

# View the unique diagnoses for samples assigned to NR tissue source
inventory[inventory["Tissue Source"] == "NR"][
    "Specimen_Histological_Diagnosis"
].unique()

# the remaining records without an assigned tissue source are from the lymphatic system
inventory.loc[inventory["Tissue Source"].isna(), "Tissue Source"] = "Lymphatic System"


# View the records without an assigned tissue source
inventory[inventory["Tissue Source"].isna()][
    ["Specimen_Histological_Diagnosis", "Sub_Type"]
]


# In[16]: Define a function to assign values to the new columns
def assign_values(
    anatomical_location,
    anatomical_sublocation,
    specimen_type,
    diagnosis,
    subtype,
    i,
    pairing_status="NR",
    pairing_site="NA",
    molecular_information="NR",
):
    inventory.loc[i, "Anatomical Location"] = (
        anatomical_location  # specify a location or NR
    )
    inventory.loc[i, "Anatomical Sub-Location"] = (
        anatomical_sublocation  # specify a sub-location or NR
    )
    inventory.loc[i, "Specimen Type"] = (
        specimen_type  # specify the specimen type (assume primary by default)
    )
    inventory.loc[i, "Histological Diagnosis (Original)"] = (
        diagnosis  # specify the diagnosis or NR
    )
    inventory.loc[i, "Histological Subtype"] = (
        subtype  # specify the subtype, including Matched Met/Primary or NR
    )
    inventory.loc[i, "Paired Sample"] = pairing_status  # Y or NR
    inventory.loc[i, "Pairing Site / Primary Tumour Site"] = (
        pairing_site  # specify site or NA or NR
    )
    inventory.loc[i, "Molecular Information"] = (
        molecular_information  # specify information or NR
    )


# Define a function for formatting and exporting the sheets to Box
def format_and_export(file_name, tissue_source):
    # Define the red and green fill colors
    greenFill = PatternFill(start_color="B1F2AA", end_color="B1F2AA", fill_type="solid")
    redFill = PatternFill(start_color="F58D7F", end_color="F58D7F", fill_type="solid")

    # Export the sheet
    data = inventory[inventory["Tissue Source"] == tissue_source][
        [
            "Tissue Source",
            "Gender",
            "Ethnicity",
            "T",
            "N",
            "M",
            "Grade",
            "Stage",
            "Specimen_Histological_Diagnosis",
            "Sub_Type",
            "Anatomical Location",
            "Anatomical Sub-Location",
            "Specimen Type",
            "Histological Diagnosis (Original)",
            "Histological Subtype",
            "Paired Sample",
            "Pairing Site / Primary Tumour Site",
            "Molecular Information",
        ]
    ]
    data.to_excel(
        f"C:/Users/medaamm/Box/Data file (AA)/Other Files/{file_name}_inventory.xlsx",
        index=False,
    )

    # Load the workbook and worksheet
    path = f"C:/Users/medaamm/Box/Data file (AA)/Other Files/{file_name}_inventory.xlsx"
    wb = load_workbook(path)
    ws = wb.active

    # format the column width
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except Exception:
                logger.exception("Error computing cell length for worksheet column %s", column)
        adjusted_width = max_length + 2
        ws.column_dimensions[column].width = adjusted_width

    columns_red = ["Specimen_Histological_Diagnosis", "Sub_Type"]
    columns_green = [
        "Anatomical Location",
        "Anatomical Sub-Location",
        "Specimen Type",
        "Histological Diagnosis (Original)",
        "Histological Subtype",
        "Paired Sample",
        "Pairing Site / Primary Tumour Site",
        "Molecular Information",
    ]
    for col_name in columns_red:
        col_index = data.columns.get_loc(col_name) + 1
        col_letter = get_column_letter(col_index)
        for cell in ws[col_letter]:
            cell.fill = redFill
    for col_name in columns_green:
        col_index = data.columns.get_loc(col_name) + 1
        col_letter = get_column_letter(col_index)
        for cell in ws[col_letter]:
            cell.fill = greenFill

    wb.save(path)

    # Export the unique combinations of diagnosis and subtype
    unique_combinations = inventory[inventory["Tissue Source"] == tissue_source][
        [
            "Specimen_Histological_Diagnosis",
            "Sub_Type",
            "Anatomical Location",
            "Anatomical Sub-Location",
            "Specimen Type",
            "Histological Diagnosis (Original)",
            "Histological Subtype",
            "Paired Sample",
            "Pairing Site / Primary Tumour Site",
            "Molecular Information",
        ]
    ]
    unique_combinations = unique_combinations.drop_duplicates(keep="first")
    unique_combinations.to_excel(
        f"C:/Users/medaamm/Box/Data file (AA)/Other Files/{file_name}_unique_combos.xlsx",
        index=False,
    )

    # Load the unique combinations workbook and worksheet
    path = f"C:/Users/medaamm/Box/Data file (AA)/Other Files/{file_name}_unique_combos.xlsx"
    wb = load_workbook(path)
    ws = wb.active

    # format the column width
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except Exception as exc:
                logger.exception("Error computing cell length for worksheet column %s: %s", column, exc)
        adjusted_width = max_length + 2
        ws.column_dimensions[column].width = adjusted_width

    columns_red = ["Specimen_Histological_Diagnosis", "Sub_Type"]
    columns_green = [
        "Anatomical Location",
        "Anatomical Sub-Location",
        "Specimen Type",
        "Histological Diagnosis (Original)",
        "Histological Subtype",
        "Paired Sample",
        "Pairing Site / Primary Tumour Site",
        "Molecular Information",
    ]
    for col_name in columns_red:
        col_index = unique_combinations.columns.get_loc(col_name) + 1
        col_letter = get_column_letter(col_index)
        for cell in ws[col_letter]:
            cell.fill = redFill
    for col_name in columns_green:
        col_index = unique_combinations.columns.get_loc(col_name) + 1
        col_letter = get_column_letter(col_index)
        for cell in ws[col_letter]:
            cell.fill = greenFill

    wb.save(path)
    return data[
        [
            "Tissue Source",
            "Specimen_Histological_Diagnosis",
            "Sub_Type",
            "Anatomical Location",
            "Anatomical Sub-Location",
            "Specimen Type",
            "Histological Diagnosis (Original)",
            "Histological Subtype",
            "Paired Sample",
            "Pairing Site / Primary Tumour Site",
            "Molecular Information",
        ]
    ].drop_duplicates(keep="first"), inventory[
        inventory["Tissue Source"] == tissue_source
    ]


# Define a function to create the combinations dictionnary
def create_combinations_dict(tissue_source):
    combinations = {}
    unique_diagnosis = list(
        inventory[inventory["Tissue Source"] == tissue_source][
            "Specimen_Histological_Diagnosis"
        ].unique()
    )
    for diagnosis in unique_diagnosis:
        combinations[diagnosis] = inventory[
            (inventory["Tissue Source"] == tissue_source)
            & (inventory["Specimen_Histological_Diagnosis"] == diagnosis)
        ]["Sub_Type"].unique()
    return combinations


# Define a function to show the values of specific keys inside the combinations dictionnary
def show_values(lst):
    for key in lst:
        values = combinations.get(key, None)
        if isinstance(values, (list, tuple, set)):
            print(f"{key}: {', '.join(map(str, values))}")
        else:
            print(f"{key}: {values}")


# In[17]: Pancreatic (Verified)

# View the unique combinations for Pancreas
create_combinations_dict("Pancreas")

# Loop through the records for Pancreas and assign values based on diagnosis and subtype
for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Pancreas":
        if diagnosis in [
            "Pancreatic",
            "Pancreatic Cancer",
            "Pancreatic Adc (Primary)",
            "Pancreatic Adc Primary Only",
            "Pancreas",
            "Pancreatic Pri.",
            "Pancreatic Pri.",
            "Pancreas Pri.",
        ]:
            if subtype in [
                "Adc",
                "Adenocarcinoma",
                "-",
                "Adenocacinoma",
                "No Data Received",
            ]:
                assign_values(
                    anatomical_location="Pancreas",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype="NR",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

            elif subtype == "Pancreas Tail Adc":
                assign_values(
                    anatomical_location="Pancreas",
                    anatomical_sublocation="Pancreas (Tail)",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype="NR",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

            elif subtype in ["Ductal Adc", "Ductal Adenocarcinoma"]:
                assign_values(
                    anatomical_location="Pancreas",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype="Ductal Adenocarcinoma",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )
            elif subtype in ["Ductal Exocrine Adc", "Ductal  Exocrine Adc"]:
                assign_values(
                    anatomical_location="Pancreas",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype="Ductal Exocrine Adenocarcinoma",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

            elif subtype == "Normal (Non Malignant)" or subtype == "Normal":
                assign_values(
                    anatomical_location="Pancreas",
                    anatomical_sublocation="NR",
                    specimen_type="N",
                    diagnosis="No Cancer",
                    subtype="No Cancer",
                    pairing_status="No Cancer",
                    pairing_site="No Cancer",
                    molecular_information="NR",
                    i=i,
                )

            elif subtype == "No Data Received":
                assign_values(
                    anatomical_location="Pancreas",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="NR",
                    subtype="NR",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

        if diagnosis == "Pancreatic Adc Matched Primary":
            assign_values(
                anatomical_location="Pancreas",
                anatomical_sublocation="NR",
                specimen_type="P",
                diagnosis="Adenocarcinoma",
                subtype="Matched Primary",
                pairing_status="P",
                pairing_site="Pancreas",
                molecular_information="NR",
                i=i,
            )


for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Pancreas":
        if diagnosis == "Pancreas-Normal" or diagnosis == "Pancreas Nor.":
            assign_values(
                anatomical_location="Pancreas",
                anatomical_sublocation="NR",
                specimen_type="N",
                diagnosis="No Cancer",
                subtype="No Cancer",
                pairing_status="No Cancer",
                pairing_site="No Cancer",
                molecular_information="NR",
                i=i,
            )

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Pancreas":
        if diagnosis in ["Pancreatic Adc Matched Mets"]:
            assign_values(
                anatomical_location="Pancreas",
                anatomical_sublocation="NR",
                specimen_type="M",
                diagnosis="Adenocarcinoma",
                subtype="Matched Met",
                pairing_status="P",
                pairing_site="Pancreas",
                molecular_information="NR",
                i=i,
            )

        if diagnosis == "Pancreatic Cholangiocarcinoma":
            inventory.loc[i, "Histological Diagnosis (Original)"] = (
                "Distal Cholangiocarcinoma"
            )
            inventory.loc[i, "Molecular Information"] = "NR"
            if subtype == "Matched Mets":
                inventory.loc[i, "Specimen Type"] = "M"
                inventory.loc[i, "Paired Sample"] = "P"
                inventory.loc[i, "Pairing Site / Primary Tumour Site"] = "Pancreas"
                inventory.loc[i, "Histological Subtype"] = "Matched Met"
                inventory.loc[i, "Anatomical Location"] = "Pancreas"
                inventory.loc[i, "Anatomical Sub-Location"] = "Pancreas (Head)"
            else:
                inventory.loc[i, "Specimen Type"] = "P"
                inventory.loc[i, "Paired Sample"] = "NR"
                inventory.loc[i, "Pairing Site / Primary Tumour Site"] = "NA"
                inventory.loc[i, "Histological Subtype"] = "NR"
                inventory.loc[i, "Anatomical Location"] = "Pancreas"
                inventory.loc[i, "Anatomical Sub-Location"] = "NR"

        if diagnosis == "Pnet" or diagnosis == "Pnet Primary":
            inventory.loc[i, "Histological Diagnosis (Original)"] = (
                "Neuroendocrine Tumour"
            )
            inventory.loc[i, "Molecular Information"] = "NR"

            if subtype == "Matched Mets":
                inventory.loc[i, "Specimen Type"] = "M"
                inventory.loc[i, "Paired Sample"] = "P"
                inventory.loc[i, "Pairing Site / Primary Tumour Site"] = "Pancreas"
                inventory.loc[i, "Histological Subtype"] = "Matched Met"
                inventory.loc[i, "Anatomical Location"] = "Pancreas"
                inventory.loc[i, "Anatomical Sub-Location"] = "NR"
            else:
                inventory.loc[i, "Specimen Type"] = "P"
                inventory.loc[i, "Paired Sample"] = "NR"
                inventory.loc[i, "Pairing Site / Primary Tumour Site"] = "NA"
                inventory.loc[i, "Histological Subtype"] = "NR"
                inventory.loc[i, "Anatomical Location"] = "Pancreas"
                inventory.loc[i, "Anatomical Sub-Location"] = "NR"


for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Pancreas":
        if diagnosis in [
            "C25.2 Malignant Neoplasm: Tail Of Pancreas",
            "C25.0 Malignant Neoplasm: Head Of Pancreas",
            "C25 Malignant Neoplasm Of Pancreas",
            "C25.1 Malignant Neoplasm: Body Of Pancreas",
        ]:
            if subtype in [
                "Ductal Adenocarcinoma",
                "Ffpe - 10% Nbf",
                "Ductal Adenocarcinoma With Focal Squamous Differentiation",
            ]:
                inventory.loc[i, "Histological Diagnosis (Original)"] = "Adenocarcinoma"
                inventory.loc[i, "Histological Subtype"] = (
                    "Adenocarcinoma With Focal Squamous Differentiation"
                    if "With Focal Squamous Differentiation" in subtype
                    else "Ductal Adenocarcinoma"
                    if "Ductal" in subtype
                    else "NR"
                )
                inventory.loc[i, "Specimen Type"] = "P"
                inventory.loc[i, "Anatomical Location"] = "Pancreas"
                inventory.loc[i, "Anatomical Sub-Location"] = (
                    "Pancreas (Head)"
                    if "Head" in diagnosis
                    else "Pancreas (Tail)"
                    if "Tail" in diagnosis
                    else "Pancreas (Body)"
                    if "Body" in diagnosis
                    else "NR"
                )
                inventory.loc[i, "Paired Sample"] = "NR"
                inventory.loc[i, "Pairing Site / Primary Tumour Site"] = "NA"
                inventory.loc[i, "Molecular Information"] = "NR"

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Pancreas":
        if diagnosis in [
            "Pancreas Pri. And Nodal Met.",
            "Pancreatic Pri. And Nodal Met.",
            "Pancreatic Pri. And Nodal Met.",
            "Pancreatic Pri. And Nodal Met.",
        ]:
            assign_values(
                anatomical_location="Pancreas",
                anatomical_sublocation="NR",
                specimen_type="P",
                diagnosis=(
                    "Adenocarcinoma"
                    if subtype in ["Adc", "Exocrine Adc", "Ductal Adc"]
                    else "Acinar Cell Carcinoma"
                    if subtype in ["Acir Cell Carcinoma"]
                    else "NR"
                ),
                subtype=(
                    "Ductal Adenocarcinoma"
                    if subtype in ["Ductal Adc"]
                    else "Exocrine Adenocarcinoma"
                    if subtype in ["Exocrine Adc"]
                    else "NR"
                ),
                pairing_status="P",
                pairing_site="Lymph Nodes",
                i=i,
            )
        elif diagnosis in ["Pancreatic Nodal Met."]:
            assign_values(
                anatomical_location="Lymph Nodes",
                anatomical_sublocation="NR",
                specimen_type="M",
                diagnosis="Adenocarcinoma",
                subtype="Ductal Adenocarcinoma" if subtype in ["Ductal Adc"] else "NR",
                pairing_status="P",
                pairing_site="Pancreas",
                i=i,
            )

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Pancreas":
        if diagnosis in ["Pancreas"] and subtype in ["-"]:
            assign_values(
                anatomical_location="Pancreas",
                anatomical_sublocation="NR",
                specimen_type="NR",
                diagnosis="NR",
                subtype="NR",
                i=i,
            )


# Export to Box & Load it to format the column width
pancreatic, pancreatic_inventory = format_and_export(
    file_name="pancreatic", tissue_source="Pancreas"
)


# In[18]: Gastric (Verified)

# View the unique combinations for Stomach
create_combinations_dict("Stomach")

# Loop through the records for Stomach and assign values based on diagnosis and subtype
for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Stomach":
        if diagnosis in [
            "Gastric Carcinoma",
            "Gastrectomy",
            "Gastric Pri.",
            "Gastric Pri",
            "Stomach Cancer",
            "Stomach",
            "Gastric",
            "Gastric Ca",
            "Gastric Cohort",
            "Gastric Adenocarcinomas",
            "Gastric Adc",
        ]:
            if subtype in [
                "Gastric Adc",
                "Adc",
                "Agc",
                "-",
                "Adenocarcinoma And Carcinoma",
                "Stomach Adc",
                "Adenocarcinoma (Background Helicobacter Gastritis)",
                "Annular Cell Carcinoma",
                "Total Gastrectomy And Anastomosis Of Oesophagus To Jejunum Nec",
            ]:
                assign_values(
                    anatomical_location="Stomach",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype="NR",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )
            elif subtype in ["Gastric Adc Mucosecretory", "Adc Mucosecretory"]:
                assign_values(
                    anatomical_location="Stomach",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype="Mucinous Adenocarcinoma",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

            elif subtype in [
                "Gastric Adc  Intestinal Type",
                "Gastric Adc Enteroid Type",
                "Adc Intestinal Type",
                "Intestinal Type Adenocarcinoma",
            ]:
                assign_values(
                    anatomical_location="Stomach",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype="Intestinal Adenocarcinoma",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

            elif subtype in [
                "Gastric Adc Signet Ring Cell",
                "Adc Signet Ring Cell",
                "Signet Ring Adc",
            ]:
                assign_values(
                    anatomical_location="Stomach",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype="Signet Ring Cell Adenocarcinoma",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

            elif subtype in ["Adc Papillary", "Papillary Gastric Adc"]:
                assign_values(
                    anatomical_location="Stomach",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype="Papillary Adenocarcinoma",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

            elif subtype in [
                "Diffuse Gastric Adc",
                "Diffuse Type Adenocarcinoma",
                "Diffuse Type  Gastric Adenocarcinoma",
                "Diffusely Infiltrating Adenocarcinoma",
            ]:
                assign_values(
                    anatomical_location="Stomach",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype="Diffuse Adenocarcinoma",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

            elif subtype in ["Mucinous Adenocarcinoma"]:
                assign_values(
                    anatomical_location="Stomach",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype="Mucinous Adenocarcinoma",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

            elif subtype in ["Gastric Adc Tubular Type"]:
                assign_values(
                    anatomical_location="Stomach",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype="Tubular Adenocarcinoma",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )


for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Stomach":
        if diagnosis in [
            "Gastric Carcinoma",
            "Gastrectomy",
            "Gastric Pri.",
            "Gastric Pri",
            "Stomach Cancer",
            "Stomach",
            "Gastric",
            "Gastric Ca",
            "Gastric Cohort",
            "Gastric Adenocarcinomas",
            "Gastric Adc",
        ]:
            if subtype in ["Adjacent Normal", "Normal Adjacent Tissues"]:
                assign_values(
                    anatomical_location="Stomach",
                    anatomical_sublocation="NR",
                    specimen_type="N (Adjacent)",
                    diagnosis="No Cancer",
                    subtype="No Cancer",
                    pairing_status="No Cancer",
                    pairing_site="No Cancer",
                    molecular_information="NR",
                    i=i,
                )

            elif subtype == "Adenocarcinoma":
                assign_values(
                    anatomical_location="Stomach",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype="NR",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

            elif subtype in ["Matched Met"]:
                assign_values(
                    anatomical_location="Check",
                    anatomical_sublocation="Check",
                    specimen_type="M",
                    diagnosis="Adenocarcinoma",
                    subtype="Matched Met",
                    pairing_status="P",
                    pairing_site="Stomach",
                    molecular_information="NR",
                    i=i,
                )

            elif subtype in ["Mixed Glandular Neuroendocrine Carcinoma"]:
                assign_values(
                    anatomical_location="Stomach",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Mixed Neuroendocrine-Non-Neuroendocrine Neoplasm",
                    subtype="Adenocarcinoma-Neuroendocrine Carcinoma",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

            elif subtype == "Gist":
                assign_values(
                    anatomical_location="Stomach",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Gastrointestinal Stromal Tumor",
                    subtype="NR",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

            elif subtype == "Intramucosal Adenocarcinoma":
                assign_values(
                    anatomical_location="Stomach",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype="Intramucosal Adenocarcinoma",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

            elif subtype == "Undifferentiated Carcinoma":
                assign_values(
                    anatomical_location="Stomach",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype="Undifferentiated Adenocarcinoma",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

            elif subtype == "Polypoid Adenocarcinoma":
                assign_values(
                    anatomical_location="Stomach",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype="Polypoid Adenocarcinoma",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

            elif subtype == "Gastric Adc Lymphoepithelial Typy":
                assign_values(
                    anatomical_location="Stomach",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype="Adenocarcinoma With Lymphoid Stroma",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

            elif subtype == "Cardiac Tumour" or subtype == "Cardia Adenocarcinoma":
                assign_values(
                    anatomical_location="Stomach",
                    anatomical_sublocation="Stomach (Cardia)",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype="NR",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

            elif subtype == "Adc Diffuse / Mixed":
                assign_values(
                    anatomical_location="Stomach",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Mixed Adenocarcinoma",
                    subtype="Intestinal And Diffuse Adenocarcinoma",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

            elif subtype == "Adc Mixed":
                assign_values(
                    anatomical_location="Stomach",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Mixed Adenocarcinoma",
                    subtype="NR",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

            elif subtype == "Residual Papillary Cardia Adenocarcinoma":
                assign_values(
                    anatomical_location="Stomach",
                    anatomical_sublocation="Stomach (Cardia)",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype="Residual Papillary Adenocarcinoma",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Stomach":
        if diagnosis in [
            "Gastric Carcinoma",
            "Gastrectomy",
            "Gastric Pri.",
            "Gastric Pri",
            "Stomach Cancer",
            "Stomach",
            "Gastric",
            "Gastric Ca",
            "Gastric Cohort",
            "Gastric Adenocarcinomas",
            "Gastric Adc",
        ]:
            if subtype == "Mixed Intestinal And Diffuse Adenocarcinoma":
                assign_values(
                    anatomical_location="Stomach",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Mixed Adenocarcinoma",
                    subtype="Intestinal And Diffuse Adenocarcinoma",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

            elif subtype == "Adenocarcinoma With A Diffuse Pattern Of Growth":
                assign_values(
                    anatomical_location="Stomach",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype="Adenocarcinoma With A Diffuse Pattern Of Growh",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

            elif subtype == "Gastric Carcinoma With Lymphoid Stroma":
                assign_values(
                    anatomical_location="Stomach",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype="Adenocarcinoma With Lymphoid Stroma",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

            elif subtype == "Metastatic Adenocarcinoma":
                assign_values(
                    anatomical_location="Check",
                    anatomical_sublocation="Check",
                    specimen_type="M",
                    diagnosis="Adenocarcinoma",
                    subtype="NR",
                    pairing_status="P",
                    pairing_site="Stomach",
                    molecular_information="NR",
                    i=i,
                )

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Stomach":
        if diagnosis in ["Gastric Adc Signet Ring Cell"]:
            assign_values(
                anatomical_location="Stomach",
                anatomical_sublocation="NR",
                specimen_type="P",
                diagnosis="Adenocarcinoma",
                subtype="Signet Ring Cell Adenocarcinoma",
                pairing_status="NR",
                pairing_site="NA",
                molecular_information="NR",
                i=i,
            )

        elif diagnosis in ["Stomach (Non Malignant)", "Gastric Nor."]:
            if subtype == "Normal Adjacent":
                assign_values(
                    anatomical_location="Stomach",
                    anatomical_sublocation="NR",
                    specimen_type="N",
                    diagnosis="No Cancer",
                    subtype="No Cancer",
                    pairing_status="No Cancer",
                    pairing_site="No Cancer",
                    molecular_information="NR",
                    i=i,
                )

            elif subtype == "Stomach Muskular":
                assign_values(
                    anatomical_location="Stomach",
                    anatomical_sublocation="Stomach (Muskular)",
                    specimen_type="N",
                    diagnosis="No Cancer",
                    subtype="No Cancer",
                    pairing_status="No Cancer",
                    pairing_site="No Cancer",
                    molecular_information="NR",
                    i=i,
                )
        elif diagnosis in ["Gastric Adc Primary & Matched Mets"]:
            assign_values(
                anatomical_location="Stomach",
                anatomical_sublocation="NR",
                specimen_type="P",
                diagnosis="Adenocarcinoma",
                subtype=(
                    "Low-Differentiated Adenocarcinoma - Matched Met"
                    if "Low-Differentiated" in subtype
                    else "NR"
                ),
                pairing_status="P",
                pairing_site="Check",
                molecular_information="NR",
                i=i,
            )

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Stomach":
        if (
            diagnosis
            in [
                "Stomach, Fundus",
                "Stomach",
                "Stomach, Muskular",
                "Gastric Nor.",
                "Stomach Nor.",
                "Stomach Fundus Nor.",
                "Stomach Muscular Nor.",
            ]
            and subtype == "Normal"
        ):
            assign_values(
                anatomical_location="Stomach",
                anatomical_sublocation=(
                    "Stomach (Fundus)"
                    if "Fundus" in subtype or "Fundus" in diagnosis
                    else "Stomach (Muscular)"
                    if any(keyword in subtype for keyword in ["Muskular", "Muscular"])
                    or any(keyword in diagnosis for keyword in ["Muskular", "Muscular"])
                    else "NR"
                ),
                specimen_type="N",
                diagnosis="No Cancer",
                subtype="No Cancer",
                pairing_status="No Cancer",
                pairing_site="No Cancer",
                molecular_information="NR",
                i=i,
            )
        elif diagnosis in [
            "Gastric Pri. And Nodal Met.",
            "Gastric Nodal Met.",
            "Stomach",
        ] and subtype in [
            "Adc",
            "Adc Mucinous",
            "Adc Signet Ring Cell",
            "Adc Intestil Type",
            "Adc Intestinal Type",
            "Diffuse Gastric Adc",
            "Gastric Adc  Intestinal Type",
            "Mzl",
        ]:
            assign_values(
                anatomical_location="Stomach",
                anatomical_sublocation="NR",
                specimen_type=("M" if diagnosis in ["Gastric Nodal Met."] else "P"),
                diagnosis=(
                    "Adenocarcinoma"
                    if subtype
                    in [
                        "Adc",
                        "Adc Mucinous",
                        "Adc Signet Ring Cell",
                        "Adc Intestil Type",
                        "Adc Intestinal Type",
                        "Diffuse Gastric Adc",
                        "Gastric Adc  Intestinal Type",
                    ]
                    else "Marginal Zone Lymphoma"
                    if subtype in ["Mzl"]
                    else "NR"
                ),
                subtype=(
                    "Mucinous Adenocarcinoma"
                    if subtype in ["Adc Mucinous"]
                    else "Signet Ring Cell Adenocarcinoma"
                    if subtype in ["Adc Signet Ring Cell"]
                    else "Intestinal Adenocarcinoma"
                    if subtype
                    in [
                        "Adc Intestinal Type",
                        "Gastric Adc  Intestinal Type",
                        "Adc Intestil Type",
                    ]
                    else "Diffuse Adenocarcinoma"
                    if subtype in ["Diffuse Gastric Adc"]
                    else "NR"
                ),
                pairing_status=(
                    "P" if diagnosis in ["Gastric Pri. And Nodal Met."] else "NR"
                ),
                pairing_site=(
                    "Lymph Nodes"
                    if diagnosis
                    in ["Gastric Pri. And Nodal Met.", "Gastric Nodal Met."]
                    else "NA"
                ),
                i=i,
            )


# Export to Box & Load it to format the column width
gastric, gastric_inventory = format_and_export(
    file_name="gastric", tissue_source="Stomach"
)


# In[19]: Breast (Verified)

# Update the unique combinations for Breast
for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Breast":
        if diagnosis in [
            "Breast",
            "Left Breast",
            "Breast Cancer",
            "Right Breast",
            "Breast Pri.",
            "Breast Primary Tumor",
            "Breast Primary",
            "Breast Pri. (No Neoadjuvant)",
            "Breast Pri. (Neoadjuvant)",
        ]:
            # Identical conditions that apply to all samples below
            inventory.loc[i, "Anatomical Location"] = "Breast"
            inventory.loc[i, "Anatomical Sub-Location"] = (
                "Breast (Right)"
                if "Right" in diagnosis
                else "Breast (Left)"
                if "Left" in diagnosis
                else "NR"
            )
            inventory.loc[i, "Specimen Type"] = (
                "NR" if diagnosis == "Breast" and subtype == "" else "P"
            )
            inventory.loc[i, "Paired Sample"] = "NR"
            inventory.loc[i, "Pairing Site / Primary Tumour Site"] = "NA"

            if subtype in [
                "Cdi",
                "Infiltrating Duct Carcinoma",
                "Hr+/ Ductal",
                "Idc",
                "Invasive Duct Carcinoma",
                "Cdi (Er+)/Her2+, Brca-",
            ]:
                inventory.loc[i, "Histological Diagnosis (Original)"] = (
                    "Invasive Carcinoma"
                )
                inventory.loc[i, "Histological Subtype"] = "Invasive Ductal Carcinoma"
                inventory.loc[i, "Molecular Information"] = (
                    "HR+"
                    if subtype == "Hr+/ Ductal"
                    else "ER+ HER2+"
                    if subtype == "Cdi (Er+)/Her2+, Brca-"
                    else "ER+ PR+ HER2-"
                    if subtype == "Infiltrative Carcinoma (Er+, Pr+, Her2Neu- )"
                    else "ER+"
                    if subtype == "Nst (Er+)"
                    else "NR"
                )

            elif subtype in [
                "Invasive Carcinoma Nst",
                "Infiltrative Carcinoma",
                "Invasive Carcinma (Nst).",
                "Infiltrative Carcinoma (Er+, Pr+, Her2Neu- )",
                "Nst (Er+)",
            ]:
                inventory.loc[i, "Histological Diagnosis (Original)"] = (
                    "Invasive Carcinoma"
                )
                inventory.loc[i, "Histological Subtype"] = "Invasive Carcinoma (NST)"
                inventory.loc[i, "Molecular Information"] = (
                    "HR+"
                    if subtype == "Hr+/ Ductal"
                    else "ER+ HER2+ BRCA-"
                    if subtype == "Cdi (Er+)/Her2+, Brca-"
                    else "ER+ PR+ HER2-"
                    if subtype == "Infiltrative Carcinoma (Er+, Pr+, Her2Neu- )"
                    else "ER+"
                    if subtype == "Nst (Er+)"
                    else "NR"
                )

            elif subtype in ["Cdi (Mucinous)"]:
                inventory.loc[i, "Histological Diagnosis (Original)"] = (
                    "Invasive Carcinoma"
                )
                inventory.loc[i, "Histological Subtype"] = (
                    "Mucinous Invasive Ductal Carcinoma"
                )
                inventory.loc[i, "Molecular Information"] = "NR"

            elif subtype in ["Ilc", "Er+, Lobular", "Hr+/ Lobular", "Cli"]:
                inventory.loc[i, "Histological Diagnosis (Original)"] = (
                    "Invasive Carcinoma"
                )
                inventory.loc[i, "Histological Subtype"] = "Invasive Lobular Carcinoma"
                inventory.loc[i, "Molecular Information"] = (
                    "ER+"
                    if subtype == "Er+, Lobular"
                    else "HR+"
                    if subtype == "Hr+/ Lobular"
                    else "NR"
                )

            elif subtype in ["Hr+/ Papillary"]:
                inventory.loc[i, "Histological Diagnosis (Original)"] = (
                    "Invasive Carcinoma"
                )
                inventory.loc[i, "Histological Subtype"] = (
                    "Invasive Papillary Carcinoma"
                )
                inventory.loc[i, "Molecular Information"] = (
                    "HR+" if subtype == "Hr+/ Papillary" else "NR"
                )

            elif subtype in [
                "Hr+/ Mixed Ductal And Lobular",
                "Mixed Ductal And Lobular (Er+)",
                "Multifocal Mixed Nst And Lobular \xa0(Er+)/Her2-/Brca (Mut)",
            ]:
                inventory.loc[i, "Histological Diagnosis (Original)"] = (
                    "Mixed Invasive Carcinoma"
                )
                inventory.loc[i, "Histological Subtype"] = (
                    "Multifocal Invasive Ductal And Lobular Carcinoma"
                    if "Multifocal" in subtype
                    else "Invasive Ductal And Lobular Carcinoma"
                )
                inventory.loc[i, "Molecular Information"] = (
                    "HR+"
                    if subtype == "Hr+/ Mixed Ductal And Lobular"
                    else "ER+"
                    if subtype == "Mixed Ductal And Lobular (Er+)"
                    else "ER+ HER2- BRCA (MUT)"
                    if subtype
                    == "Multifocal Mixed Nst And Lobular \xa0(Er+)/Her2-/Brca (Mut)"
                    else "NR"
                )

            elif subtype in [
                "Breast Ca (Er+)",
                "Er+",
                "-",
                "Hr+",
                "",
                "Her2+/ Hr-",
                "Her2+/ Hr+",
            ]:
                inventory.loc[i, "Histological Diagnosis (Original)"] = "Carcinoma"
                inventory.loc[i, "Histological Subtype"] = "NR"
                inventory.loc[i, "Molecular Information"] = (
                    "ER+"
                    if "Er+" in subtype
                    else "HER2+ HR+"
                    if "Her2+/ Hr+" in subtype
                    else "HER2+ HR-"
                    if "Her2+/ Hr-" in subtype
                    else "HR+"
                    if "Hr+" in subtype
                    else "NR"
                )

            elif subtype in ["Lobular-Classical (Er+)"]:
                inventory.loc[i, "Histological Diagnosis (Original)"] = (
                    "Invasive Carcinoma"
                )
                inventory.loc[i, "Histological Subtype"] = (
                    "Classic Invasive Lobular Carcinoma"
                )
                inventory.loc[i, "Molecular Information"] = "ER+"

            elif subtype in ["Basal (Er+)/Her2-,Brca-"]:
                inventory.loc[i, "Histological Diagnosis (Original)"] = (
                    "Breast Carcinoma"
                )
                inventory.loc[i, "Histological Subtype"] = (
                    "Breast Carcinoma With Basal-Like Features"
                )
                inventory.loc[i, "Molecular Information"] = "ER+ HER2- BRCA-"

            elif subtype in ["Pleomorphic Carcinoma"]:
                inventory.loc[i, "Histological Diagnosis (Original)"] = (
                    "Invasive Carcinoma"
                )
                inventory.loc[i, "Histological Subtype"] = (
                    "Pleomorphic Invasive Carcinoma"
                )
                inventory.loc[i, "Molecular Information"] = "NR"

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Breast":
        if diagnosis in [
            "Breast",
            "Left Breast",
            "Breast Cancer",
            "Right Breast",
            "Breast Pri.",
            "Breast Primary Tumor",
            "Breast Primary",
            "Breast Pri. (No Neoadjuvant)",
            "Breast Pri. (Neoadjuvant)",
        ]:
            if subtype in ["Adjacent Normal", "Normal Adjacent Tissues"]:
                inventory.loc[i, "Histological Diagnosis (Original)"] = "No Cancer"
                inventory.loc[i, "Histological Subtype"] = "No Cancer"
                inventory.loc[i, "Anatomical Location"] = "Breast"
                inventory.loc[i, "Anatomical Sub-Location"] = (
                    "Breast (Right)"
                    if "Right" in diagnosis
                    else "Breast (Left)"
                    if "Left" in diagnosis
                    else "NR"
                )
                inventory.loc[i, "Specimen Type"] = "N (Adjacent)"
                inventory.loc[i, "Paired Sample"] = "No Cancer"
                inventory.loc[i, "Pairing Site / Primary Tumour Site"] = "No Cancer"
                inventory.loc[i, "Molecular Information"] = "NR"


for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Breast":
        if diagnosis in ["Breast Pri."]:
            if subtype in ["Unknown"]:
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Carcinoma",
                    subtype="NR",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

            elif subtype in ["Cdi (Unknown)", "Cdi - (Nst)", "Cdi (Er+)"]:
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Invasive Carcinoma",
                    subtype="Invasive Ductal Carcinoma",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="ER+" if subtype == "Cdi (Er+)" else "NR",
                    i=i,
                )

            elif subtype in ["Cdi Tubular (Er+)"]:
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Invasive Carcinoma",
                    subtype="Tubular Invasive Carcinoma",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="ER+"
                    if subtype == "Cdi Tubular (Er+)"
                    else "NR",
                    i=i,
                )

            elif subtype in ["Tnbc"]:
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Carcinoma",
                    subtype="NR",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="TNBC",
                    i=i,
                )

            elif subtype in [
                "Cdi Mucinous",
                "Mucinous Carcinoma",
                "Idc Muc.",
                "Cdi  (Tnbc)",
            ]:
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Invasive Carcinoma",
                    subtype=(
                        "Mucinous Invasive Ductal Carcinoma"
                        if any(word in subtype for word in ["Mucinous", "Muc."])
                        else "Invasive Ductal Carcinoma"
                    ),
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information=("TNBC" if "Tnbc" in subtype else "NR"),
                    i=i,
                )

            elif subtype in [
                "Mixed Nst And Lobular (Er+)",
                "Mixed Ductal Carcinoma No Specific Type (Nst) And Lobular Carcinoma",
                "Mixed Nst And Lobular (Tnbc)",
            ]:
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Mixed Invasive Carcinoma",
                    subtype="Invasive Ductal And Lobular Carcinoma",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information=(
                        "ER+"
                        if subtype == "Mixed Nst And Lobular (Er+)"
                        else "TNBC"
                        if subtype == "Mixed Nst And Lobular (Tnbc)"
                        else "NR"
                    ),
                    i=i,
                )

            elif subtype == "Mixed Nst, Lobular And Tubular (Tnbc)":
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Mixed Invasive Carcinoma",
                    subtype="Invasive Ductal, Lobular And Tubular Carcinoma",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="TNBC",
                    i=i,
                )

            elif subtype in [
                "Invasive Carcinoma Of No Special Type (Nst) - Triple-Negative Breast Cancer"
            ]:
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Invasive Carcinoma",
                    subtype="Invasive Carcinoma (NST)",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="TNBC",
                    i=i,
                )
            elif subtype in [
                "Idc / (Nst) (Tnbc)",
                "(Idc) - Pair Tnbc",
                "Ductal/Nst (Tnbc)",
                "Ductal Nst (Basal Type) (Tnbc)",
            ]:
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Invasive Carcinoma",
                    subtype="Invasive Ductal Carcinoma",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="TNBC",
                    i=i,
                )
            elif subtype == "Cdi Papillary":
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation="NR",
                    diagnosis="Invasive Carcinoma",
                    subtype="Invasive Papillary Carcinoma",
                    specimen_type="P",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

            elif subtype in ["(Idc) - Pair Her2 Neg Neoadjuvant"]:
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation="NR",
                    diagnosis="Invasive Carcinoma",
                    subtype="Invasive Ductal Carcinoma",
                    specimen_type="P",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="HER2-",
                    i=i,
                )

            elif subtype in [
                "(Idc) - Diff. Rec.-Status Neoadjuvant Tx",
                "(Ilc) - Diff. Rec.-Status Neoadjuvant Tx",
                "(Idc) - Diff. Rec.-Status Neoadjuvant Tx (Her2 Neg)",
                "(Ilc) - Diff. Rec.-Status Neoadjuvant Tx (Her2 Neg)",
            ]:
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation="NR",
                    diagnosis="Invasive Carcinoma",
                    subtype=(
                        "Invasive Lobular Carcinoma"
                        if "(Ilc)" in subtype
                        else "Invasive Ductal Carcinoma"
                    ),
                    specimen_type="P",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information=(
                        "HER2-" if "Her2 Neg" in subtype else "Diff. Receptor Status"
                    ),
                    i=i,
                )

            elif subtype in ["Idc + Scc", "Scc"]:
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis=(
                        "Mixed Carcinoma"
                        if subtype == "Idc + Scc"
                        else "Squamous Cell Carcinoma"
                    ),
                    subtype=(
                        "Invasive Ductal and Squamous Cell Carcinoma"
                        if subtype == "Idc + Scc"
                        else "NR"
                    ),
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

            elif subtype in [
                "Ca Metaplastico",
                "Metaplastic Carcinoma",
                "(Mc) - Pair Tnbc",
            ]:
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Metaplastic Carcinoma",
                    subtype="NR",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="TNBC",
                    i=i,
                )

            elif subtype in ["Medullary-Like/Nst  (Tnbc)"]:
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Invasive Carcinoma",
                    subtype="Invasive Carcinoma (NST) With Medullary-Like Features",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="TNBC",
                    i=i,
                )

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Breast":
        if diagnosis in ["Breast", "Tnbc", "Breast Pri. And Met."]:
            if subtype == "Tnbc" or subtype == "-":
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation="NR",
                    specimen_type=(
                        "NR" if diagnosis == "Breast" and subtype == "" else "P"
                    ),
                    diagnosis=(
                        "NR" if diagnosis == "Breast" and subtype == "" else "Carcinoma"
                    ),
                    subtype="NR",
                    pairing_status=("P" if "Pri. And Met" in subtype else "NR"),
                    pairing_site=("Check" if "Pri. And Met" in subtype else "NA"),
                    molecular_information="TNBC"
                    if (diagnosis == "Tnbc" or subtype == "Tnbc")
                    else "NR",
                    i=i,
                )
            elif diagnosis == "Breast" and subtype in ["-", ""]:
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation="NR",
                    specimen_type="NR",
                    diagnosis="NR",
                    subtype="NR",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

        elif diagnosis in [
            "Invasive Ductal/No Special Type (Nst)",
            "Nst",
            "Nst/Apocrine",
            "Medullary-Like/Nst",
        ]:
            if subtype == "Tnbc":
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Invasive Carcinoma",
                    subtype=(
                        "Invasive Carcinoma (NST) With Medullary-Like Features"
                        if "Medullary-Like" in diagnosis
                        else "Invasive Carcinioma (NST) With Apocrine Differentiation"
                        if "Apocrine" in diagnosis
                        else "Invasive Ductal Carcinoma"
                        if "Ductal" in diagnosis
                        else "NR"
                    ),
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="TNBC",
                    i=i,
                )

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Breast":
        if diagnosis in [
            "Breast Primary Tumor",
            "Breast Cancer - Triple-Negative Breast Cancer",
            "Right Breast",
            "Breast Cancer",
            "Breast Primary",
        ]:
            if subtype in ["Ductal Carcinoma In Situ - Triple-Negative Breast Cancer"]:
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Ductal Carcinoma In Situ",
                    subtype="NR",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="TNBC",
                    i=i,
                )

            elif subtype in ["Metaplastic Carcinoma - Triple-Negative Breast Cancer"]:
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Metaplastic Carcinoma",
                    subtype="NR",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="TNBC",
                    i=i,
                )

            elif subtype in [
                "Invasive Ductal Carcinoma - Triple-Negative Breast Cancer",
                "Invasive Carcinoma Of No Special Type (Nst) - Triple-Negative Breast Cancer",
                "Invasive Carcinoma Nst (Tnbc)",
                "Idc (Nst) (Tnbc)",
            ]:
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation=(
                        "Breast (Right)" if "Right" in diagnosis else "NR"
                    ),
                    specimen_type="P",
                    diagnosis="Invasive Carcinoma",
                    subtype=(
                        "Invasive Ductal Carcinoma"
                        if "Ductal" in subtype
                        else "Invasive Carcinoma (NST)"
                    ),
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="TNBC",
                    i=i,
                )

            elif diagnosis == "Breast Cancer" and subtype == "Tnbc":
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Carcinoma",
                    subtype="NR",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="TNBC",
                    i=i,
                )

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Breast":
        if diagnosis == "Breast Cancer":
            if subtype == "Cli - Solid And Classical.Er+/Her2-/Brca (Mut)":
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Invasive Carcinoma",
                    subtype="Invasive Lobular Carcinoma With Solid and Classical Pattern",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="ER+ HER2-",
                    i=i,
                )

            elif subtype == "Cdi (Her-2 (+))":
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Invasive Carcinoma",
                    subtype="Invasive Ductal Carcinoma",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="HER2+",
                    i=i,
                )

        if diagnosis == "Breast Pri. And Nodal Met.":
            if subtype == "Nst / Apocrine (Tnbc)":
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Invasive Carcinoma",
                    subtype="Invasive Carcinioma (NST) With Apocrine Differentiation",
                    pairing_status="P",
                    pairing_site="Lymph Nodes",
                    molecular_information="TNBC",
                    i=i,
                )

            elif subtype == "Lobular Pleomorphic And Apocrine (Tnbc)":
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Invasive Carcinoma",
                    subtype="Pleomorphic And Apocrine Carcinoma",
                    pairing_status="P",
                    pairing_site="Lymph Nodes",
                    molecular_information="TNBC",
                    i=i,
                )

            elif subtype == "Basal (Tnbc)":
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Invasive Carcinoma",
                    subtype="Invasive Carcinoma (NST)",
                    pairing_status="P",
                    pairing_site="Lymph Nodes",
                    molecular_information="Basal Type TNBC",
                    i=i,
                )
            elif subtype == "Nst (Er+)":
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Invasive Carcinoma",
                    subtype="Invasive Carcinoma (NST)",
                    pairing_status="P",
                    pairing_site="Lymph Nopdes",
                    molecular_information="ER+",
                    i=i,
                )
            elif subtype == "Mixed Ductal And Lobular (Er+)":
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Mixed Invasive Carcinoma",
                    subtype="Invasive Ductal And Lobular Carcinoma",
                    pairing_status="P",
                    pairing_site="Lymph Nopdes",
                    molecular_information="ER+",
                    i=i,
                )

        if diagnosis == "Basal":
            if subtype in ["Tnbc"]:
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Carcinoma",
                    subtype="NR",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="Basal Type TNBC",
                    i=i,
                )


for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Breast":
        if diagnosis == "Breast Nor.":
            assign_values(
                anatomical_location="Breast",
                anatomical_sublocation=(
                    "Breast (Glands)" if "Glands" in subtype else "NR"
                ),
                specimen_type="N",
                diagnosis="No Cancer",
                subtype="No Cancer",
                pairing_status="No Cancer",
                pairing_site="No Cancer",
                molecular_information="NR",
                i=i,
            )
        elif diagnosis in ["Breast", "Breast, Glands"] and subtype == "Normal":
            assign_values(
                anatomical_location="Breast",
                anatomical_sublocation=(
                    "Breast (Glands)" if "Glands" in diagnosis else "NR"
                ),
                specimen_type="N",
                diagnosis="No Cancer",
                subtype="No Cancer",
                pairing_status="No Cancer",
                pairing_site="No Cancer",
                molecular_information="NR",
                i=i,
            )

        elif diagnosis == "Breast Cancer - Triple-Negative Breast Cancer":
            if subtype == "C50 Malignant Neoplasm Of Breast":
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Carcinoma",
                    subtype="NR",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="TNBC",
                    i=i,
                )

        elif diagnosis == "Bone Met. (Breast)":
            if subtype == "Ilc":
                assign_values(
                    anatomical_location="Bone",
                    anatomical_sublocation="NR",
                    specimen_type="M",
                    diagnosis="Invasive Carcinoma",
                    subtype="Invasive Lobular Carcinoma",
                    pairing_status="P",
                    pairing_site="Breast",
                    molecular_information="NR",
                    i=i,
                )

        elif diagnosis == "Cdi (Her-2 (+))":
            if subtype == "Breast Ca (Her2+)":
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Invasive Carcinoma",
                    subtype="Invasive Carcinoma (NST)",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="HER2+",
                    i=i,
                )


for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Breast":
        if diagnosis == "Breast Pri. And Nodal Met.":
            if subtype == "Lobular-Classical (Er+)":
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Invasive Carcinoma",
                    subtype="Classic Invasive Lobular Carcinoma",
                    pairing_status="P",
                    pairing_site="Lymph Nodes",
                    molecular_information="ER+",
                    i=i,
                )
            elif subtype == "Mucinous (Er+)":
                assign_values(
                    anatomical_location="Breast",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Invasive Carcinoma",
                    subtype="Mucinous Invasive Ductal Carcinoma",
                    pairing_status="P",
                    pairing_site="Lymph Nodes",
                    molecular_information="ER+",
                    i=i,
                )
        elif diagnosis == "Lymph Node Met. (Breast)" and subtype == "Idc":
            assign_values(
                anatomical_location="Lymph Nodes",
                anatomical_sublocation="NR",
                specimen_type="M",
                diagnosis="Invasive Carcinoma",
                subtype="Invasive Ductal Carcinoma",
                pairing_status="NR",
                pairing_site="Breast",
                i=i,
            )


for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Breast":
        if diagnosis == "Breast Metastatic":
            assign_values(
                anatomical_location="Breast",
                anatomical_sublocation="NR",
                specimen_type="M",
                diagnosis=(
                    "Ductal Carcinoma In Situ"
                    if subtype
                    in ["Ductal Carcinoma In Situ - Triple-Negative Breast Cancer"]
                    else "Metaplastic Carcinoma"
                    if subtype
                    in ["Metaplastic Carcinoma - Triple-Negative Breast Cancer"]
                    else "Invasive Carcinoma"
                    if subtype
                    in [
                        "Invasive Ductal Carcinoma - Triple-Negative Breast Cancer",
                        "Invasive Carcinoma Of No Special Type (Nst) - Triple-Negative Breast Cancer",
                    ]
                    else "NR"
                ),
                subtype=(
                    "Invasive Ductal Carcinoma"
                    if subtype
                    in ["Invasive Ductal Carcinoma - Triple-Negative Breast Cancer"]
                    else "Invasive Carcinoma (NST)"
                    if subtype
                    in [
                        "Invasive Carcinoma Of No Special Type (Nst) - Triple-Negative Breast Cancer"
                    ]
                    else "NR"
                ),
                pairing_status="NR",
                pairing_site="Breast",
                molecular_information=(
                    "TNBC" if "Triple-Negative Breast Cancer" in subtype else "NR"
                ),
                i=i,
            )

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Breast":
        if diagnosis == "Breast" and subtype == "-":
            assign_values(
                anatomical_location="Breast",
                anatomical_sublocation="NR",
                specimen_type="NR",
                diagnosis="NR",
                subtype="NR",
                i=i,
            )


# Export to Box & Load it to format the column width
breast, breast_inventory = format_and_export(file_name="breast", tissue_source="Breast")


# In[20]: Lung (Verified)

# Create combinations dictionary for Lung
combinations = create_combinations_dict("Lung")
show_values(list(combinations.keys()))


for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Lung":
        base_diagnosis = ""

        if subtype in [
            "Acinar Adc",
            "Acinar Adenocarcinoma",
            "Acinar Adenocarcinoma With Colloid Component (20%)",
            "Adc",
            "Adc (Acinar Adenocarcinoma With Lepidic Component)",
            "Adc (Acinar Adenocarcinoma)",
            "Adc (Enteric Adenocarcinoma)",
            "Adc (Invasive Mucinous Adenocarcinoma)",
            "Adc (Lung Adenocarcinoma Lung Metastasis)",
            "Adc (Minimal Scc Component)",
            "Adc (Minimally Invasive  Adenocarcinoma )",
            "Adc (Mucinous Invasive Adenocarcinoma)",
            "Adc (Nos Adenocarcinoma), Adc (Papillary )",
            "Adc (Papillary Adenocarcinoma)",
            "Adc (Small Squamous Cell Component)",
            "Adc (Solid Adenocarcinoma)",
            "Adc (Solid, Non-Mucinous)",
            "Adc / Acinar",
            "Adc / Brochioloalveolar",
            "Adc / Bronchioalveolar Cell Type",
            "Adc / Broncho-Alveolar/Variable Differentiation",
            "Adc / Clear Cell Pattern",
            "Adc / Focal Squamous",
            "Adc / Mucin Secreting",
            "Adc / Mucinous",
            "Adc / Papillary",
            "Adc / Prominent Solid Component",
            "Adc Acinar",
            "Adc Acinar With Areas Of Lepidic",
            "Adc Micropapillary",
            "Adc Mucinous",
            "Adc Multifocal",
            "Adc Papillary",
            "Adc Post-Treatment",
            "Adc Pre-Treatment",
            "Adenocacinoma",
            "Adenocarcinoma",
            "Adenocarcinoma [Adc]",
            "Adenocarcinoma, Lung",
            "Adénocarcinome Invasif Mucineux",
            "Enteric Adenocarcinoma",
            "Fetal Adenocarcinoma",
            "Invasive Acinar Adenocarcinoma",
            "Invasive Mucinous Adenocarcinoma",
            "Lepedic Adenocarcinoma",
            "Lepidic Adenocarcinoma",
            "Lung Adenocarcinoma",
            "Lung Adenocarcinoma Lymph Node Mestastasis",
            "Lung Adenocarcinoma Metastasis",
            "Lung Adenocarcinoma Parietal Metastasis",
            "Lung Adenocarcinoma Pleura Metastasis",
            "Lung Adenocarcinoma Pleural Metastasis",
            "Lung Metastasis Of Lung Adenocarcinoma",
            "Lung Mucinous Lepidic Adenocarcinoma",
            "Lymph Node Metastasis Of Lung Adenocarcinoma",
            "Metastatic, Invasive, Lung Adenocarcinoma",
            "Micropapillary Adenocarcinoma",
            "Minimal Invasive Adenocarcinoma",
            "Minimally Invasive Adenocarcinoma",
            "Mucinous Adc Post-Treatment",
            "Mucinous Adc Pre Treatment",
            "Mucinous Adenocarcinoma",
            "Mucinous Lepidic Adenocarcinoma",
            "Papillar Adenocarcinoma",
            "Papillary Adenocarcinoma",
            "Pleural Metastasis Of Lung Adenocarcinoma",
            "Poorly Differentiated Adenocarcinoma",
            "Solid Adenocarcinoma",
            "Tubuloacinar Adenocarcinoma.",
            "Undifferentiated Adenocarcinoma",
            "Bronchio-Alveolar Cell Carcinoma",
            "Bronchoalveolar Cell Carcinoma",
            "Adc (Nos Adenocarcinoma)",
            "Adc (Papillary )",
            "Lung Adenocarcinoma Lymph Node Metastasis",
            "Papillary Adenocarinoma",
            "Lepidic Mucinous Adenocarcinoma",
            "Lieberkuhn Adenocarcinoma Lung Metastase",
            "Adc (Papillary Adenocarcinoma With Micropapillary And Solid Components)",
            "Adc (Papillary And Micropapillary Adenocarcinoma)",
            "Adc / Acinar And Papillary",
            "Adc / Acinar And Solid",
            "Adc / Small Cell Carcinoma",
            "Adc / Undifderentiated Spindle Cell",
            "Adc / Undifferentiated Foci",
            "Adc / Mixed Acinar, Papillary And Solid Pattern",
            "Adc / Mixed Pattern",
            "Adc / Mixed Patterns",
            "Adc / Mixed Pattern",
            "Adc / Mixed Patterns",
            "Adc Acinar Papillary",
            "Adc Acinar Papillary And Micropapillary",
            "Adc Lepidic And Papillary",
            "Adc Mucinous Papillary",
            "Adc Papillary And Micropapillary",
            "Lepidic And Papillary Adenocarcinoma",
            "Adc (Papillary And Micropapillary Adenocarcinoma",
            "Adenocarcinoma Of Predominantly Lepidic And Lesser Acinar Pattern",
            "Papillary And Micropapillary Adenocarcinoma",
            "Tubulopapillary And Micropapillary Adenocarcinoma (Lung Primary)",
            "Intestinal Adenocarcinoma",
            "Nsclc (Adc)",
        ] or diagnosis in ["Lieberkuhn Adenocarcinoma Lung Metastase"] or diagnosis in [
            "Rectum Adenocarcinoma Lung Metastase",
            "Sigmoid Adenocarcinoma Lung Metastase",
            "Rectum Adenocarcinoma Lung Metastase",
            "Sigmoid Adenocarcinoma Lung Metastase",
        ]:
            base_diagnosis = "Adenocarcinoma"
        elif subtype in [
            "Acinar (60%) And Solid (40%) Adenocarcinoma",
            "Acinar (70%) Solid (30%) Adenocarcinoma",
            "Mixed Carcinoma (70% Adenocarcinoma, 30% Large Cell Carcinoma)",
            "Solid (50%) Acinar (40%) And Micropapillary (10%) Adenocarcinoma",
            "Solid (90%) Adenocarcinoma With Acinar Component (10%)",
        ]:
            base_diagnosis = "Mixed Adenocarcinoma"
        elif subtype == "Other Mixed":
            base_diagnosis = "Mixed (NOS)"
        elif subtype in [
            "Adenosquamous",
            "Adenosquamous Carcinoma",
            "Adeno-Squamous Carcinoma",
            "Adc Adenosquamous",
        ]:
            base_diagnosis = "Adenosquamous Carcinoma"
        elif subtype in [
            "Basaloid Squamous Cell Carcinoma",
            "Keratinizing Squamous Cell Carcinoma",
            "Lung Squamous Cell Carcinoma Lung Metastasis",
            "Non-Keratinizing Squamous Cell Carcinoma",
            "Scc",
            "Scc (Basaloid Squamous Cell Carcinoma)",
            "Scc (Keratinizing Squamous Cell Carcinoma)",
            "Scc (Non-Keratinizing Squamous Cell Carcinoma)",
            "Scc (Squamous Cell Carcinoma)",
            "Sq Cell Carcinoma",
            "Squamous",
            "Squamous [Scc]",
            "Squamous Carcinoma",
            "Squamous Cell Carcinoma",
            "Squamous Cell Carcinoma (Lung)",
            "Squamous Cell Carcinoma (Scc)",
            "Squamous Cell Carcinoma (Tonsil)",
            "Squamous Cell Carcinoma Lung Metastasis",
            "Squamous Cell Carcinoma, Lung",
            "Squamous Cell Carcinome",
            "Squamous Large Cell Carcinoma",
            "Squamouse Cell Carcinoma",
            "Keratinizing Squamous Cell Carinoma",
            "Non-Keratinizing Squamous Cell Carinoma",
            "Scc / Cear Cell",
            "Nsclc (Scc)",
            "Nsclc-Scc",
        ]:
            base_diagnosis = "Squamous Cell Carcinoma"
        elif subtype in ["Small Cell Carcinoma"] and diagnosis not in ["Nsclc"]:
            base_diagnosis = "Small Cell Lung Cancer"
        elif subtype in ["Carcinoma", "Epidermoid Carcinoma"]:
            base_diagnosis = "Carcinoma"
        elif diagnosis in ["Lung Cancer", "Lung Pri."] and subtype in ["-", "?"]:
            base_diagnosis = "Cancer"
        elif subtype in [
            "Pleomorphic Carcinoma",
            "Pleomorphic Carcinoma (Spindle Cell And Adenocarcinoma)",
            "Sarcomatoid Carcinoma",
            "Spindle Cell Carcinoma",
        ]:
            base_diagnosis = "Sarcomatoid Carcinoma"
        elif subtype in [
            "Combined Carcinoma, With Squamous Cell Carcinoma And High Grade Neuroendocrine Carcinoma Components.",
            "Combined Large Cell Neuroendocrine And Small Cell Carcinoma",
            "Combined Large Cell Neuroendocrine Carcinoma",
            "Combined Large Cell Neuroendocrine Carcinoma And Small Cell Carcinoma",
            "Combined Large Cell Neuroendocrine Carcinoma And Small Cell Carcinoma",
            "Combined Small Cell And Large Neuroendocrine Carcinoma",
            "Combined Small Cell And Poorly Differentiated Squamous Cell Carcinoma",
            "Combined Small Cell Carcinoma And Large Cell Neuroendocrine Carcinoma",
            "Combined Small Cell/ Large Cell Neuroendocrine Carcinoma",
            "Combined Squamous And Large Cell Neuroendocrine Carcinoma",
            "Combined Squamous And Large Cell Neuroendocrine Carcinoma",
            "Combined Squamous Cell Carcinoma And Probable Small Cell Neuroendocrine Carcinoma",
            "Combined Tumour Containing Adenocarcinoma And Large Cell Neuroendocrine Carcinoma",
            "Predominantly Poorly Differentiated Neuroendocrine Carcinoma With Areas Of Small Cell Carcinoma And Focal Squamous Immunophenotype",
            "Mixed (Sc / Nsclc)",  # combined small cell lung carcinoma
            "Scc-Small Cell Carcinoma",
            "Large Cell Neuroendocrine Carcinoma With Possible Focal Squamous Differentiation And Possible Focal Areas Of A Small Cell Carcinoma",
        ]:
            base_diagnosis = "Combined Carcinoma"
        elif subtype in [
            "Adjacent Normal",
            "Matched Normal (Non-Malignant) Lung Blocks",
            "Non-Malignant",
            "Normal",
            "Normal (Non-Malignant)",
            "Normal Adjacent",
            "Normal Adjacent Tissues",
            "Normal Lung",
            "Matched Normal",
        ]:
            base_diagnosis = "No Cancer"
        elif subtype in ["Idiopathic Pulmonary Fibrosis (Ipf)", "Ipf"]:
            base_diagnosis = "Idiopathic Pulmonary Fibrosis (Ipf)"
        elif subtype in ["Large Cell Carcinoma", "Large Cell, Undifferentiated"]:
            base_diagnosis = "Large Cell Carcinoma"
        elif subtype in ["Large Cell Neuroendocrine Carcinoma"]:
            base_diagnosis = "Large Cell Neuroendocrine Carcinoma"
        elif subtype in [
            "Lung Small Cell Neuroendocrine Carcinoma Lung Metastasis",
            "Lung Small Cell Neuroendocrine Carcinoma Pleura Metastasis",
        ]:
            base_diagnosis = "Small Cell Neuroendocrine Carcinoma"
        elif subtype in [
            "Met Blocks/Metastatic Renal Cell Carcinoma (Multiple)",
            "Metastatic Renal Cell Carcinoma (Conventional)",
        ]:
            base_diagnosis = "Renal Cell Carcinoma"
        elif subtype in [
            "G12C",
            "Non-Small Cell Carcinoma",
            "Nsclc",
            "Nsclc (Adc)",
            "Nsclc (Met. +)",
            "Nsclc (Scc)",
            "Nsclc-Scc",
        ]:
            base_diagnosis = "Non-Small Cell Lung Cancer"
        elif subtype in ["Sclc"]:
            base_diagnosis = "Small Cell Lung Cancer"
        elif subtype in ["Small And Large Cell Neuroendocrine Carcinoma"]:
            base_diagnosis = "Neuroendocrine Carcinoma"
        elif subtype in [
            "Adjacent Normal",
            "Matched Normal (Non-Malignant) Lung Blocks",
            "Non-Malignant",
            "Normal",
            "Normal (Non-Malignant)",
            "Normal Adjacent",
            "Normal Adjacent Tissues",
            "Normal Lung",
            "Matched Normal",
        ] or diagnosis in ["Healthy Lung Tissue"]:
            base_diagnosis = "No Cancer"
        elif diagnosis == "Lung Metastasis":
            base_diagnosis = "Cancer"
        elif diagnosis == "Lung Pri." and subtype == "Other (Not Specified)":
            base_diagnosis == "Cancer"
        else:
            base_diagnosis = "NR"

        # get the final diagnosis
        final_diagnosis = base_diagnosis
        if (
            base_diagnosis not in ["NR", "Non-Small Cell Lung Cancer"]
            and (
                diagnosis in ["Nsclc", "Non Small Cell Lung Cancer"]
                or any(
                    keyword in subtype
                    for keyword in ["Nsclc", "Non Small Cell Lung Cancer"]
                )
            )
            and subtype not in ["Matched Normal", "Normal"]
        ):
            final_diagnosis = base_diagnosis + " (Non-Small Cell Lung Cancer)"
        elif base_diagnosis not in [
            "NR",
            "Non-Small Cell Lung Cancer",
        ] and diagnosis in ["Nsclc Post-Treatment"]:
            final_diagnosis = (
                base_diagnosis + " (Post-Treatment - Non-Small Cell Lung Cancer)"
            )
        elif base_diagnosis == "NR" and diagnosis in [
            "Nsclc",
            "Non Small Cell Lung Cancer",
            "Nsclc Post-Treatment",
        ]:
            final_diagnosis = "Non-Small Cell Lung Cancer"
        elif base_diagnosis == "NR" and diagnosis in ["Nsclc Post-Treatment"]:
            final_diagnosis = "Non-Small Cell Lung Cancer (Post-Treatment)"
        elif base_diagnosis == "Non-Small Cell Lung Cancer" and diagnosis in [
            "Nsclc",
            "Non Small Cell Lung Cancer",
            "Nsclc Post-Treatment",
        ]:
            final_diagnosis = base_diagnosis
        elif (
            base_diagnosis not in ["NR", "Small Cell Lung Cancer"]
            and (
                diagnosis in ["Sclc"]
                or any(
                    keyword in subtype for keyword in ["Sclc", "Small Cell Lung Cancer"]
                )
            )
        ) and "Normal" not in subtype:
            final_diagnosis = base_diagnosis + " (Small Cell Lung Cancer)"
        elif base_diagnosis == "NR" and diagnosis in ["Sclc"]:
            final_diagnosis = "Small Cell Lung Cancer"
        elif base_diagnosis == "Small Cell Lung Cancer":
            final_diagnosis = base_diagnosis

        assign_values(
            anatomical_location=(
                "Bronchus"
                if diagnosis in ["Bronchus"]
                else "Lung"
                if (
                    any(
                        keyword in diagnosis
                        for keyword in ["Lung", "Lobe", "Nsclc", "Sclc"]
                    )
                )
                else "Lung/Lymph Nodes"
                if diagnosis in ["Wedge Of Upper Lobe & Node", "Lung + Nodes"]
                else "NR"
            ),
            anatomical_sublocation=(
                "Lung (Left Lower Lobe)"
                if diagnosis
                in [
                    "Left Lower Lobe",
                    "Left Lower Lobe, Lung",
                    "Left Lung, Lower Lobe",
                    "Lower Left Lobe",
                    "Lung, Left Lower Lobe",
                ]
                else "Lung (Left)"
                if diagnosis in ["Left Lung"]
                else "Lung (Right)"
                if diagnosis in ["Right Lung"]
                else "Lung (Left Upper Lobe)"
                if diagnosis in ["Upper Lobe Of Left Lung"]
                else "Lung (Right Upper Lobe)"
                if diagnosis
                in [
                    "Right Upper Lobe",
                    "Lung, Right Upper Lobe",
                    "Right Lung, Upper Lobe",
                ]
                else "Lung (Left Upper)"
                if diagnosis
                in [
                    "Left Lung, Upper Lobe",
                    "Left Upper Lobe",
                    "Left Upper Lobe, Lung",
                    "Left Upper Lung",
                ]
                else "Lung (Lower Lobe)"
                if diagnosis in ["Lung, Lower Lobe"]
                else "Lung (Lower Lobe - Anterior)"
                if diagnosis in ["Lung, Lower Lobe Anterior"]
                else "Lung (Parenchyma)"
                if diagnosis in ["Lung, Parenchyma"]
                else "Lung (Right Lower Lobe)"
                if diagnosis
                in [
                    "Right Lung, Lower Lobe",
                    "Lung, Right Lower Lobe",
                    "Right Lower Lobe",
                    "Right Lower Lobe, Lung",
                ]
                else "Lung (Right Middle Lobe)"
                if diagnosis in ["Right Lung, Middle Lobe"]
                else "Lung (Right Middle And Lower Lobe)"
                if diagnosis in ["Right Middle And Lower Lobe"]
                else "Lung (Right Middle Lobe)"
                if diagnosis in ["Right Middle Lobe"]
                else "Lung (Right Middle And Upper Lobe)"
                if diagnosis in ["Right Upper And Middle Lobe"]
                else "Lung (Upper Lobe)"
                if diagnosis in ["Upper Lobe Of Right Lung"]
                else "Lung (Upper Lobe) And Node"
                if diagnosis in ["Wedge Of Upper Lobe & Node"]
                else "NR"
            ),
            specimen_type=(
                "N"
                if (
                    any(
                        keyword in diagnosis
                        for keyword in ["Nor.", "Normal", "Healthy Lung Tissue"]
                    )
                    or any(
                        keyword in subtype
                        for keyword in [
                            "Nor.",
                            "Normal",
                            "Ipf",
                            "Idiopathic Pulmonary Fibrosis (Ipf)",
                            "Non-Malignant",
                        ]
                    )
                    and "Adjacent" not in subtype
                )
                else "N (Adjacent)"
                if subtype
                in ["Adjacent Normal", "Normal Adjacent", "Normal Adjacent Tissues"]
                else "M"
                if (
                    (
                        any(
                            keyword in diagnosis
                            for keyword in [
                                "Lieberkuhn Adenocarcinoma Lung Metastase",
                                "Lung Met.",
                                "Lung Metastasis",
                                "Lung Metastase",
                            ]
                        )
                        or subtype
                        in [
                            "Metastatic, Invasive, Lung Adenocarcinoma",
                            "Matched Met (Lh18-19367)",
                            "Metastasis",
                            "Lung Adenocarcinoma Pleura Metastasis",
                            "Lung Small Cell Neuroendocrine Carcinoma Lung Metastasis",
                            "Lung Small Cell Neuroendocrine Carcinoma Pleura Metastasis",
                            "Pleural Metastasis Of Lung Adenocarcinoma",
                            "Lung Metastasis Of Lung Adenocarcinoma",
                            "Lymph Node Metastasis Of Lung Adenocarcinoma",
                            "Lung Adenocarcinoma Metastasis",
                            "Lung Adenocarcinoma Pleura Metastasis",
                            "Lung Squamous Cell Carcinoma Lung Metastasis",
                            "Squamous Cell Carcinoma Lung Metastasis",
                            "Nsclc (Met. +)",
                            "Lung Adenocarcinoma Pleural Metastasis",
                            "Lung Adenocarcinoma Lymph Node Metastasis",
                            "Lung Adenocarcinoma Parietal Metastasis",
                            "Adc (Lung Adenocarcinoma Lung Metastasis)",
                            "Metastatic Renal Cell Carcinoma (Conventional)",
                            "Met Blocks/Metastatic Renal Cell Carcinoma (Multiple)",
                            "Squamous Cell Carcinoma (Tonsil)",
                            "Lung Metastase",
                        ]
                    )
                    and (diagnosis not in ["Lung Pri", "Lung Pri."])
                )
                else "NR"
                if diagnosis in ["Lung-Pd1", "Lung + Nodes"] and subtype in ["-"]
                else "P"
            ),
            diagnosis=final_diagnosis,
            subtype=(
                "Acinar (60%) And Solid (40%) Adenocarcinoma"
                if subtype in ["Acinar (60%) And Solid (40%) Adenocarcinoma"]
                else "Acinar (70%) Solid (30%) Adenocarcinoma"
                if subtype in ["Acinar (70%) Solid (30%) Adenocarcinoma"]
                else "Adenocarcinoma (70%) And Large Cell Carcinoma (30%)"
                if subtype
                in ["Mixed Carcinoma (70% Adenocarcinoma, 30% Large Cell Carcinoma)"]
                else "Acinar (40%), Solid (50%), And Micropapillary (10%) Adenocarcinoma"
                if subtype
                in ["Solid (50%) Acinar (40%) And Micropapillary (10%) Adenocarcinoma"]
                else "Acinar (10%) And Solid (90%) Adenocarcinoma"
                if subtype in ["Solid (90%) Adenocarcinoma With Acinar Component (10%)"]
                else "Acinar Adenocarcinoma"
                if subtype
                in [
                    "Acinar Adc",
                    "Acinar Adenocarcinoma",
                    "Adc (Acinar Adenocarcinoma)",
                    "Adc / Acinar",
                    "Adc Acinar",
                ]
                else "Acinar Adenocarcinoma With Colloid Component (20%)"
                if subtype in ["Acinar Adenocarcinoma With Colloid Component (20%)"]
                else "Adenocarcinoma (NOS)"
                if subtype
                in [
                    "Adc",
                    "Adc (Lung Adenocarcinoma Lung Metastasis)",
                    "Adc (Nos Adenocarcinoma)",
                ]
                else "Acinar Adenocarcinoma With Lepidic Component"
                if subtype
                in [
                    "Adc (Acinar Adenocarcinoma With Lepidic Component)",
                    "Adc Acinar With Areas Of Lepidic",
                ]
                else "Enteric Adenocarcinoma"
                if subtype
                in [
                    "Adc (Enteric Adenocarcinoma)",
                    "Enteric Adenocarcinoma",
                ]
                else "Mucinous Invasive Adenocarcinoma"
                if subtype
                in [
                    "Adc (Invasive Mucinous Adenocarcinoma)",
                    "Adc (Mucinous Invasive Adenocarcinoma)",
                    "Adénocarcinome Invasif Mucineux",
                    "Invasive Mucinous Adenocarcinoma",
                ]
                else "Adenocarcinoma With Minimal Squamous Cell Carcinoma Component"
                if subtype in ["Adc (Minimal Scc Component)"]
                else "Minimally Invasive Adenocarcinoma"
                if subtype
                in [
                    "Adc (Minimally Invasive  Adenocarcinoma )",
                    "Minimal Invasive Adenocarcinoma",
                    "Minimally Invasive Adenocarcinoma",
                ]
                else "Papillary Adenocarcinoma"
                if subtype
                in [
                    "Adc (Papillary )",
                    "Adc (Papillary Adenocarcinoma)",
                    "Adc / Papillary",
                    "Adc Papillary",
                    "Papillar Adenocarcinoma",
                    "Papillary Adenocarcinoma",
                    "Papillary Adenocarinoma",
                ]
                else "Papillary Adenocarcinoma With Micropapillary And Solid Components"
                if subtype
                in [
                    "Adc (Papillary Adenocarcinoma With Micropapillary And Solid Components)"
                ]
                else "Papillary And Micropapillary Adenocarcinoma"
                if subtype
                in [
                    "Adc (Papillary And Micropapillary Adenocarcinoma)",
                    "Adc Papillary And Micropapillary",
                    "Papillary And Micropapillary Adenocarcinoma",
                ]
                else "Adenocarcinoma With Minimal Squamous Cell Carcinoma Component"
                if subtype in ["Adc (Small Squamous Cell Component)"]
                else "Solid Adenocarcinoma"
                if subtype in ["Adc (Solid Adenocarcinoma)", "Solid Adenocarcinoma"]
                else "Solid Non-Mucinous Adenocarcinoma"
                if subtype in ["Adc (Solid, Non-Mucinous)"]
                else "Acinar And Papillary Adenocarcinoma"
                if subtype in ["Adc / Acinar And Papillary", "Adc Acinar Papillary"]
                else "Acinar And Solid Adenocarcinoma"
                if subtype in ["Adc / Acinar And Solid"]
                else "Bronchioloalveolar Carcinoma (Invasive Adenocarcinoma With A Predominant Lepidic Component)"
                if subtype
                in [
                    "Adc / Brochioloalveolar",
                    "Adc / Bronchioalveolar Cell Type",
                    "Bronchio-Alveolar Cell Carcinoma",
                    "Bronchoalveolar Cell Carcinoma",
                ]
                else "Bronchioloalveolar Carcinoma (Invasive Adenocarcinoma With A Predominant Lepidic Component) With Variable Differentiation"
                if subtype in ["Adc / Broncho-Alveolar/Variable Differentiation"]
                else "Adenocarcinoma With A Clear Cell Pattern"
                if subtype in ["Adc / Clear Cell Pattern"]
                else "Adenocarcinoma With Focal Squamous Differentiation"
                if subtype in ["Adc / Focal Squamous"]
                else "Acinar, Papillary, And Solid Adenocarcinoma"
                if subtype in ["Adc / Mixed Acinar, Papillary And Solid Pattern"]
                else "Mucinous Adenocarcinoma"
                if subtype
                in [
                    "Adc / Mucin Secreting",
                    "Adc / Mucinous",
                    "Adc Mucinous",
                    "Mucinous Adenocarcinoma",
                ]
                else "Adenocarcinoma With A Predominant Solid Component"
                if subtype in ["Adc / Prominent Solid Component"]
                else "Combined Small Cell Lung Cancer"
                if subtype in ["Mixed (Sc / Nsclc)", "Scc-Small Cell Carcinoma"]
                else "Small And Large Cell Neuroendocrine Carcinoma"
                if subtype in ["Small And Large Cell Neuroendocrine Carcinoma"]
                else "Adenocarcinoma With Small Cell Lung Cancer"
                if subtype in ["Adc / Small Cell Carcinoma"]
                else "Pleomorphic Carcinoma"
                if subtype
                in ["Adc / Undifderentiated Spindle Cell", "Pleomorphic Carcinoma"]
                else "Pleomorphic Carcinoma (Spindle Cell And Adenocarcinoma)"
                if subtype
                in ["Pleomorphic Carcinoma (Spindle Cell And Adenocarcinoma)"]
                else "Adenocarcinoma With Areas Of Undifferentiation"
                if subtype in ["Adc / Undifferentiated Foci"]
                else "Acinar, Papillary, And Micropapillary Adenocarcinoma"
                if subtype in ["Adc Acinar Papillary And Micropapillary"]
                else "Lepidic And Papillary Adenocarcinoma"
                if subtype
                in ["Adc Lepidic And Papillary", "Lepidic And Papillary Adenocarcinoma"]
                else "Micropapillary Adenocarcinoma"
                if subtype in ["Adc Micropapillary", "Micropapillary Adenocarcinoma"]
                else "Mucinous And Papillary Adenocarcinoma"
                if subtype in ["Adc Mucinous Papillary"]
                else "Mucinous Lepidic Adenocarcinoma"
                if subtype
                in [
                    "Lepidic Mucinous Adenocarcinoma",
                    "Lung Mucinous Lepidic Adenocarcinoma",
                    "Mucinous Lepidic Adenocarcinoma",
                ]
                else "Multifocal Adenocarcinoma"
                if subtype in ["Adc Multifocal"]
                else "Adenocarcinoma (Post-Treatment)"
                if subtype in ["Adc Post-Treatment"]
                else "Adenocarcinoma (Pre-Treatment)"
                if subtype in ["Adc Pre-Treatment"]
                else "Mucinous Adenocarcinoma (Post-Treatment)"
                if subtype in ["Mucinous Adc Post-Treatment"]
                else "Mucinous Adenocarcinoma (Pre-Treatment)"
                if subtype in ["Mucinous Adc Pre Treatment"]
                else "Adenocarcinoma With A Predominant Lepidic Component And Minor Acinar Pattern"
                if subtype
                in ["Adenocarcinoma Of Predominantly Lepidic And Lesser Acinar Pattern"]
                else "No Cancer"
                if subtype
                in [
                    "Adjacent Normal",
                    "Matched Normal (Non-Malignant) Lung Blocks",
                    "Non-Malignant",
                    "Normal",
                    "Normal (Non-Malignant)",
                    "Normal Adjacent",
                    "Normal Adjacent Tissues",
                    "Normal Lung",
                    "Matched Normal",
                ]
                or diagnosis in ["Healthy Lung Tissue"]
                else "Basaloid Squamous Cell Carcinoma"
                if subtype
                in [
                    "Basaloid Squamous Cell Carcinoma",
                    "Scc (Basaloid Squamous Cell Carcinoma)",
                ]
                else "Squamous Cell Carcinoma And High-Grade Neuroendocrine Carcinoma"
                if subtype
                in [
                    "Combined Carcinoma, With Squamous Cell Carcinoma And High Grade Neuroendocrine Carcinoma Components."
                ]
                else "Large Cell Neuroendocrine Carcinoma Carcinoma And Small Cell Lung Cancer"
                if subtype
                in [
                    "Combined Large Cell Neuroendocrine And Small Cell Carcinoma",
                    "Combined Large Cell Neuroendocrine Carcinoma And Small Cell Carcinoma",
                    "Combined Small Cell And Large Neuroendocrine Carcinoma",
                    "Combined Small Cell Carcinoma And Large Cell Neuroendocrine Carcinoma",
                    "Combined Small Cell/ Large Cell Neuroendocrine Carcinoma",
                ]
                else "Small Cell Lung Cancer And Poorly Differentiated Squamous Cell Carcinoma"
                if subtype
                in [
                    "Combined Small Cell And Poorly Differentiated Squamous Cell Carcinoma"
                ]
                else "Squamous Cell Carcinoma And Large Cell Neuroendocrine Carcinoma"
                if subtype
                in ["Combined Squamous And Large Cell Neuroendocrine Carcinoma"]
                else "Squamous Cell Carcinoma And Small Cell Neuroendocrine Carcinoma"
                if subtype
                in [
                    "Combined Squamous Cell Carcinoma And Probable Small Cell Neuroendocrine Carcinoma"
                ]
                else "Adenocarcinoma And Large Cell Neuroendocrine Carcinoma"
                if subtype
                in [
                    "Combined Tumour Containing Adenocarcinoma And Large Cell Neuroendocrine Carcinoma"
                ]
                else "Fetal Adenocarcinoma"
                if subtype in ["Fetal Adenocarcinoma"]
                else "Invasive Acinar Adenocarcinoma"
                if subtype in ["Invasive Acinar Adenocarcinoma"]
                else "Keratinizing Squamous Cell Carcinoma"
                if subtype
                in [
                    "Keratinizing Squamous Cell Carcinoma",
                    "Scc (Keratinizing Squamous Cell Carcinoma)",
                    "Keratinizing Squamous Cell Carinoma",
                ]
                else "Non-Keratinizing Squamous Cell Carcinoma"
                if subtype
                in [
                    "Non-Keratinizing Squamous Cell Carcinoma",
                    "Non-Keratinizing Squamous Cell Carinoma",
                    "Scc (Non-Keratinizing Squamous Cell Carcinoma)",
                ]
                else "Large Cell Neuroendocrine Carcinoma With Focal Squamous Differentiation And Focal Areas Of Small Cell Lung Cancer"
                if subtype
                in [
                    "Large Cell Neuroendocrine Carcinoma With Possible Focal Squamous Differentiation And Possible Focal Areas Of A Small Cell Carcinoma"
                ]
                else "Undifferentiated Large Cell Carcinoma"
                if subtype in ["Large Cell, Undifferentiated"]
                else "Lepidic Adenocarcinoma"
                if subtype in ["Lepedic Adenocarcinoma", "Lepidic Adenocarcinoma"]
                else "Invasive Adenocarcinoma"
                if subtype in ["Metastatic, Invasive, Lung Adenocarcinoma"]
                else "Poorly Differentiated Adenocarcinoma"
                if subtype in ["Poorly Differentiated Adenocarcinoma"]
                else "Predominantly Poorly Differentiated Neuroendocrine Carcinoma With Areas Of Small Cell Lung Cancer And Focal Squamous Immunophenotype"
                if subtype
                in [
                    "Predominantly Poorly Differentiated Neuroendocrine Carcinoma With Areas Of Small Cell Carcinoma And Focal Squamous Immunophenotype"
                ]
                else "Clear Cell Squamous Cell Carcinoma"
                if subtype in ["Scc / Cear Cell"]
                else "Spindle Cell Carcinoma"
                if subtype in ["Spindle Cell Carcinoma"]
                else "Tubuloacinar Adenocarcinoma"
                if subtype in ["Tubuloacinar Adenocarcinoma."]
                else "Undifferentiated Adenocarcinoma"
                if subtype in ["Undifferentiated Adenocarcinoma"]
                else "Epidermoid Carcinoma"
                if subtype in ["Epidermoid Carcinoma"]
                else "No Cancer"
                if subtype in ["Idiopathic Pulmonary Fibrosis (Ipf)", "Ipf"]
                else "Tubulopapillary And Micropapillary Adenocarcinoma"
                if subtype
                in ["Tubulopapillary And Micropapillary Adenocarcinoma (Lung Primary)"]
                else "Intestinal Adenocarcinoma"
                if subtype in ["Intestinal Adenocarcinoma"]
                else "Adenocarcinoma With Mixed Patterns"
                if subtype in ["Adc / Mixed Pattern", "Adc / Mixed Patterns"]
                else "NR"
            ),
            pairing_status=(
                "P"
                if subtype
                in [
                    "Matched Met (Lh18-19367)",
                    "Matched Normal (Non-Malignant) Lung Blocks",
                    "Matched Normal",
                    "Nsclc (Met.+)",
                ]
                or diagnosis in ["Wedge Of Upper Lobe & Node", "Lung + Nodes"]
                else "No Cancer"
                if (
                    subtype
                    in [
                        "Adjacent Normal",
                        "Non-Malignant",
                        "Normal",
                        "Normal (Non-Malignant)",
                        "Normal Adjacent",
                        "Normal Adjacent Tissues",
                        "Normal Lung",
                        "Ipf",
                        "Idiopathic Pulmonary Fibrosis (Ipf)",
                    ]
                    or diagnosis in ["Healthy Lung Tissue"]
                )
                else "NR"
            ),
            pairing_site=(
                "Intestines"
                if diagnosis in ["Lieberkuhn Adenocarcinoma Lung Metastase"]
                else "Lymph Nodes"
                if (
                    diagnosis in ["Lung Pri. And Nodal Met.", "Lung Pri. And Met."]
                    or subtype
                    in [
                        "Lung Adenocarcinoma Lymph Node Metastasis",
                        "Lymph Node Metastasis Of Lung Adenocarcinoma",
                        "Lung Adenocarcinoma Lymph Node Mestastasis",
                    ]
                )
                else "Lung"
                if subtype
                in [
                    "Adc (Lung Adenocarcinoma Lung Metastasis)",
                    "Lung Adenocarcinoma Metastasis",
                    "Lung Metastasis Of Lung Adenocarcinoma",
                    "Lung Small Cell Neuroendocrine Carcinoma Lung Metastasis",
                    "Lung Squamous Cell Carcinoma Lung Metastasis",
                    "Squamous Cell Carcinoma Lung Metastasis",
                    "Matched Normal (Non-Malignant) Lung Blocks",
                    "Matched Normal",
                ]
                else "Parietal Pleura"
                if subtype
                in [
                    "Lung Adenocarcinoma Parietal Metastasis",
                    "Lung Adenocarcinoma Pleura Metastasis",
                    "Lung Adenocarcinoma Pleural Metastasis",
                    "Lung Small Cell Neuroendocrine Carcinoma Pleura Metastasis",
                    "Pleural Metastasis Of Lung Adenocarcinoma",
                ]
                else "Tonsil"
                if subtype in ["Matched Met (Lh18-19367)"]
                else "Kidney"
                if subtype
                in [
                    "Met Blocks/Metastatic Renal Cell Carcinoma (Multiple)",
                    "Metastatic Renal Cell Carcinoma (Conventional)",
                ]
                else "No Cancer"
                if subtype
                in [
                    "Adjacent Normal",
                    "Matched Normal (Non-Malignant) Lung Blocks",
                    "Non-Malignant",
                    "Normal",
                    "Normal (Non-Malignant)",
                    "Normal Adjacent",
                    "Normal Adjacent Tissues",
                    "Normal Lung",
                    "Matched Normal",
                    "Idiopathic Pulmonary Fibrosis (Ipf)",
                    "Ipf",
                ]
                or diagnosis in ["Healthy Lung Tissue"]
                else "Rectum"
                if diagnosis in ["Rectum Adenocarcinoma Lung Metastase"]
                else "Colon (Sigmoid)"
                if diagnosis in ["Sigmoid Adenocarcinoma Lung Metastase"]
                else "NA"
            ),
            molecular_information=(
                "KRAS G12C"
                if subtype == "G12C"
                else "PD1"
                if diagnosis == "Lung-Pd1"
                else "NR"
            ),
            i=i,
        )


# Export to Box & Load it to format the column width
lung, lung_inventory = format_and_export(tissue_source="Lung", file_name="lung")


# In[21]: Prostate (Verified)

# create the combinations dictionary for Prostate
create_combinations_dict("Prostate")

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Prostate":
        if diagnosis in [
            "Prostate Pri.",
            "Prostate",
            "Prostate Adenocarcinoma",
            "Prostate Cancer",
            "Prostatic Adenocarcinoma",
            "Prostatectomy",
            "Prostate Pri. And Met.",
        ]:
            if subtype in [
                "Adenocarcinoma",
                "Adc",
                "Gleason 8 & Higher Adc",
                "Each Gleason Adc",
                "Adc, Focally Squamoid",
                "Adenocarcinoma With Focal Neuroendocrine Differentiation",
                "Prostatic Adenocarcinoma With Neuroendocrine Differentiation",
                "Prostatic Adenocarcinoma With Focal Neuroendocrine Differentiation",
                "Prostatic Adenocarcinoma With At Least Focal Neuroendocrine Differentation",
                "Prostatic Adenocarcinoma",
                "Acinar Adenocarcinoma",
                "Adenocarcinoma Prostate",
                "Adc Pre - Post Treatment",
                "Prostate Adenocarcinoma",
            ]:
                assign_values(
                    anatomical_location="Prostate",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype=(
                        "Adenocarcinoma With Focal Squamoid Differentiation"
                        if any(word in subtype for word in ["Focally Squamoid"])
                        else "Adenocarcinoma With Focal Neuroendocrine Differentiation"
                        if any(word in subtype for word in ["Focal Neuroendocrine"])
                        else "Adenocarcinoma With Neuroendocrine Differentiation"
                        if any(
                            word in subtype
                            for word in ["With Neuroendocrine Differentiation"]
                        )
                        else "Acinar Adenocarcinoma"
                        if subtype == "Acinar Adenocarcinoma"
                        else "NR"
                    ),
                    pairing_status=(
                        "P" if diagnosis == "Prostate Pri. And Met." else "NR"
                    ),
                    pairing_site=(
                        "Prostate" if diagnosis == "Prostate Pri. And Met." else "NA"
                    ),
                    molecular_information="NR",
                    i=i,
                )
            elif subtype in [
                "Adenocarcinoma/Pin",
                "Microacinar Adenocarcinoma/Pin",
                "Prostate Carcinoma/Pin Iii",
                "Pin Iii",
            ]:
                assign_values(
                    anatomical_location="Prostate",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis=(
                        "Adenocarcinoma"
                        if "Adenocarcinoma" in subtype
                        else "Prostatic Intraepithelial Neoplasia III"
                        if subtype == "Pin Iii"
                        else "Carcinoma"
                    ),
                    subtype=(
                        "Adenocarcinoma (Prostatic Intraepithelial Neoplasia)"
                        if subtype == "Adenocarcinoma/Pin"
                        else "Carcinoma (Prostatic Intraepithelial Neoplasia III)"
                        if subtype == "Prostate Carcinoma/Pin Iii"
                        else "Microacinar Adenocarcinoma (Prostatic Intraepithelial Neoplasia)"
                        if subtype == "Microacinar Adenocarcinoma/Pin"
                        else "NR"
                    ),
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )
            elif (
                subtype
                == "High Grade Neoplasm, Consistent With Neuroendocrine Carcinoma, Features In Keeping With Prostatic Origin"
            ):
                assign_values(
                    anatomical_location="Prostate",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Neuroendocrine Carcinoma",
                    subtype="High Grade Neoplasm",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )
            elif subtype == "Neuroendocrine Carcinoma":
                assign_values(
                    anatomical_location="Prostate",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Neuroendocrine Carcinoma",
                    subtype="NR",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

        if diagnosis in [
            "Prostate",
            "Prostate (Non Malignant)",
            "Prostate Cancer",
            "Prostatenor.",
            "Prostate Nor.",
        ]:
            if subtype in [
                "Matched Normal",
                "Adjacent Normal",
                "Normal",
                "Normal Adjacent",
            ]:
                assign_values(
                    anatomical_location="Prostate",
                    anatomical_sublocation="NR",
                    specimen_type=("N (Adjacent)" if "Adjacent" in subtype else "N"),
                    diagnosis="No Cancer",
                    subtype="No Cancer",
                    pairing_status=(
                        "P"
                        if "Matched" in subtype
                        else "No Cancer"
                        if any(
                            keyword in subtype
                            for keyword in ["Adjacnet", "Normal", "Nor"]
                        )
                        else "NR"
                    ),
                    pairing_site=(
                        "Prostate"
                        if "Matched" in subtype
                        else "No Cancer"
                        if any(
                            keyword in subtype
                            for keyword in ["Adjacnet", "Normal", "Nor"]
                        )
                        else "NA"
                    ),
                    molecular_information="NR",
                    i=i,
                )
        if diagnosis == "Prostate Pri. And Met.":
            if subtype == "Metastatic":
                assign_values(
                    anatomical_location="Prostate",
                    anatomical_sublocation="NR",
                    specimen_type="M",
                    diagnosis="Adenocarcinoma",
                    subtype="NR",
                    pairing_status="P",
                    pairing_site="Prostate",
                    molecular_information="NR",
                    i=i,
                )
        if diagnosis in ["Prostate Pri.E Nor.", "Prostate Pri. And Nor."]:
            if subtype in ["Adc"]:
                assign_values(
                    anatomical_location="Prostate",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype="NR",
                    pairing_status="P",
                    pairing_site="Prostate",
                    molecular_information="NR",
                    i=i,
                )
        if diagnosis in ["Prostate (Hrpc)"]:
            if subtype == "Adenocarcinoma":
                assign_values(
                    anatomical_location="Prostate",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype="Adenocarcinoma (HRPC)",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )
        if diagnosis in [
            "Prostatic Adenocarcinoma",
            "Prostate Adenocarcinoma",
            "Recurrent Prostate Adenocarcinoma",
        ]:
            if subtype == "-":
                assign_values(
                    anatomical_location="Prostate",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype=(
                        "Adenocarcinoma (Recurrent)"
                        if "Recurrent" in diagnosis
                        else "NR"
                    ),
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )
        if diagnosis in ["Bladder Pri.", "Bladder", "Prostate/Bladder"]:
            if (
                subtype
                == "Tcc With Squamous Cell Differentiation Invading Left Prostate"
            ):
                assign_values(
                    anatomical_location="Bladder",
                    anatomical_sublocation="Invading Left Prostate",
                    specimen_type="P",
                    diagnosis="Transitional Cell Carcinoma",
                    subtype="Transitional Cell Carcinoma With Squamous Cell Differentiation",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )
            elif (
                subtype
                == "Prostatic Carcinoma With Areas Of Neuroendocrine Differentiation (No More Than 30%)"
            ):
                assign_values(
                    anatomical_location="Bladder",
                    anatomical_sublocation="NR",
                    specimen_type="NR",
                    diagnosis="Prostate Carcinoma",
                    subtype="Prostate Carcinoma With Areas Of Neuroendocrine Differentiation (No More Than 30%)",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )
        if diagnosis == "Prostate" and subtype == "":
            assign_values(
                anatomical_location="Prostate",
                anatomical_sublocation="NR",
                specimen_type="NR",
                diagnosis="NR",
                subtype="NR",
                pairing_status="NR",
                pairing_site="NA",
                molecular_information="NR",
                i=i,
            )


for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Prostate":
        if diagnosis == "Prostate Pri." and subtype == "/":
            assign_values(
                anatomical_location="Prostate",
                anatomical_sublocation="NR",
                specimen_type="P",
                diagnosis="Carcinoma",
                subtype="NR",
                pairing_status="NR",
                pairing_site="NA",
                molecular_information="NR",
                i=i,
            )
        elif diagnosis == "Prostate/Bladder":
            if (
                subtype
                == "Prostatic Carcinoma With Widespread Neuroendocrine Differentiation (Gleason Grade Provided For Small Acinar Type Area)"
            ):
                assign_values(
                    anatomical_location="Prostate/Bladder",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Carcinoma",
                    subtype="Carcinoma With Widespread Neuroendocrine Differentiation",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )
        elif diagnosis == "Prostate Cancer":
            if subtype == "Small Cell Neuroendocrine Carcinoma":
                assign_values(
                    anatomical_location="Prostate",
                    anatomical_sublocation="NR",
                    specimen_type="P",
                    diagnosis="Neuroendocrine Carcinoma",
                    subtype="Small Cell Neuroendocrine Carcinoma",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )
            elif subtype in ["-", "C61 Malignant Neoplasm Of Prostate", "Metastatic"]:
                assign_values(
                    anatomical_location="Prostate",
                    anatomical_sublocation="NR",
                    specimen_type=("M" if "Metastatic" in subtype else "P"),
                    diagnosis="Carcinoma",
                    subtype="NR",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Prostate":
        if diagnosis == "Prostate" and subtype == "-":
            assign_values(
                anatomical_location="Prostate",
                anatomical_sublocation="NR",
                specimen_type="NR",
                diagnosis="NR",
                subtype="NR",
                pairing_status="NR",
                pairing_site="NA",
                molecular_information="NR",
                i=i,
            )


# Export to Box & Load it to format the column width
prostate, prostate_inventory = format_and_export(
    file_name="prostate", tissue_source="Prostate"
)


# In[22]: Head & Neck (Verified)

# Create the combinations dictionary for Head & Neck
sorted(create_combinations_dict("Head & Neck").keys())

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Head & Neck":
        assign_values(
            anatomical_location=(
                "Brain"
                if diagnosis
                in [
                    "Gbm",
                    "Brain",
                    "Brain Tissue",
                    "Brain (Left Frontal Space )",
                    "Brain Nor.",
                    "Brain Pri.",
                    "Cerebellum Nor.",
                    "Cerebrum",
                    "Cerebrum Nor.",
                    "Cerebrum. Nor",
                    "Gbm Pre-Treatment",
                    "Glioblastoma Cancer",
                ]
                else "Oral Cavity"
                if diagnosis
                in [
                    "Anterior Mandible",
                    "Anterior Tongue",
                    "Base Of Tongue",
                    "Anterior Tongue.",
                    "Floor Of Mouth",
                    "Floor Of Mouth And Tongue",
                    "Gum",
                    "H&N Pri. Hemiglossectomy",
                    "H&N Pri. Lateral Partial Glossectomy",
                    "H&N Pri. Left Partial Glossectomy",
                    "H&N Pri. Partial Glossectomy",
                    "H&N Pri. Tongue Base",
                    "Left Lateral Tongue",
                    "Left Partial Glossectomy",
                    "Left Soft Palate",
                    "Hen Pri. (Tongue)",
                    "Left Tongue",
                    "Left Tongue And Mandible",
                    "Lip",
                    "Mouth",
                    "R Extended Hemiglossectomy",
                    "R Lateral Tongue",
                    "R Tongue",
                    "R Ventral Tongue",
                    "Right Floor Of Mouth/Tongue",
                    "Right Hemiglossectomy",
                    "Left Chin",
                    "Right Jaw Tumour Deposit",
                    "Right Lateral Tongue",
                    "Right Lingual Mandible",
                    "Right Partial Glossectomy",
                    "Right Posterolateral Tongue",
                    "Right Soft Palate",
                    "Right Tongue",
                    "Right Tongue Base",
                    "Right Upper Lip",
                    "Right Ventral Tongue",
                    "Soft Palate",
                    "Tongue",
                ]
                else "Oral Cavity - Neck"
                if diagnosis
                in [
                    "H&N Pri. Hemiglossectomy  And Neck Dissection",
                    "H&N Pri. Hemiglossectomy And Neck Dissection",
                    "H&N Pri. Hemiglossectomy And Neck Dissections",
                    "H&N Pri. Hemiglossectomy And Right Neck Dissection",
                    "H&N Pri. Hemiglossectomy With Adjacent Floor Of Mouth And Bilateral Selective Neck Dissections",
                    "H&N Pri. Hemiglossectomy, Right Neck Dissection",
                    "H&N Pri. Left Sub-Total Glossectomy And Left Selective Neck Dissection",
                    "H&N Pri. Partial Glossectomy And Left Selective Neck Dissection",
                    "H&N Pri. Partial Glossectomy And Neck Dissection.",
                    "Tongue/Mandible/Neck",
                ]
                else "Oral Cavity - Pharynx"
                if diagnosis
                in [
                    "H&N Pri. Hemiglossectomy And Left Piriform Fossa",
                    "Tongue Base; Involves L Pyriform Fossa But Tumour Appears To Originate In Tongue Base",
                ]
                else "Oral Cavity - Pharynx"
                if diagnosis in ["D- Left Soft Palate, B-Left Piriform Fossa"]
                else "Nasal Sinus"
                if diagnosis in ["Anterior Maxilla", "Maxilla", "Left Maxilla"]
                or subtype in ["Left Maxilla"]
                else "Ear"
                if diagnosis
                in [
                    "Basal Cell Carcinoma - Left Ear",
                    "Basal Cell Carcinoma - Right Post Auricular",
                    "Ear",
                    "Left Ear",
                    "Posterior Right Ear",
                    "Right Auricular",
                    "Right Ear",
                    "Right Post Auricular",
                ]
                else "Nose"
                if diagnosis
                in [
                    "Basal Cell Carcinoma - Nose",
                    "Left Nasal Cavity",
                    "Nose",
                    "Right Nasal Cavity",
                ]
                or subtype in ["Nasal Cavity"]
                else "Nose - Oral Cavity"
                if diagnosis
                in [
                    "Right Nasal Cavity (Primary Tumours In R Nasal Cavity And L Hard/Soft Palate)"
                ]
                else "Face"
                if diagnosis
                in [
                    "Basal Cell Carcinoma - Right Upper Cheek",
                    "Cheek",
                    "Left Cheek",
                    "Righ Upper Cheek",
                    "Right Bucca",
                    "Right Buccal Sulcus",
                    "Right Cheek",
                ]
                or subtype in ["Right Cheek", "Left Cheek"]
                else "Scalp"
                if diagnosis
                in [
                    "Basal Cell Carcinoma - Scalp",
                    "Basal Cell Carcinoma - Vertex Of Scalp",
                    "Left Parietal Scalp",
                    "Left Vertex Scalp",
                    "Occipital Scalp",
                    "Right Posterior Parietal Scalp",
                    "Scalp",
                    "Vertex Of Scalp",
                ]
                or subtype in ["Right Scalp", "Left Scalp"]
                else "Temple"
                if diagnosis
                in [
                    "Basal Cell Carcinoma - Superior Temple",
                    "Left Temple",
                    "Left Temporal",
                    "Superior Temple",
                ]
                else "Pharynx"
                if diagnosis
                in [
                    "Basaloid Squamous Cell Carcinoma",
                    "Hypopharynx",
                    "Keratinising Squamous Cell Carcinoma",
                    "Laryngopharynx",
                    "Left Oropharynx",
                    "Nasopharynx",
                    "Npc",
                    "Oropharynx",
                    "Pharynx Cancer With Ln Neck Metastasis",
                    "Right Lateral Oropharyngectomy",
                    "Right Oropharynx",
                    "Right Pharyngeal Wall (Main Tumour In Supraglottis Although Tumour And High Grade Dysplasia Also Found In R Pharyngeal Wall And Tongue Base)",
                ]
                or subtype in ["Basaloid Squamous Cell Carcinoma (Oropharyngeal)"]
                else "Forehead"
                if diagnosis in ["Forehead", "Right Forehead"]
                or subtype in ["Right Forehead"]
                else "Oral Cavity - Salivary Gland"
                if diagnosis in ["Gum/Parotid Gland"]
                else "Larynx"
                if diagnosis
                in [
                    "H&N Pri. Total Laryngectomy",
                    "Larynx",
                    "Vocal Cord (Tumour Originated In L Pyriform Fossa)",
                ]
                else "Larynx - Neck"
                if diagnosis
                in [
                    "H&N Pri. Total Laryngectomy And Bilateral Selective Neck Dissections"
                ]
                else "Salivary Gland"
                if diagnosis
                in [
                    "Left Parotid",
                    "Parotid",
                    "Parotid Galnd",
                    "Parotid Gland",
                    "Salivary Gland",
                ]
                else "Salivary Gland And Lymph Nodes"
                if diagnosis
                in [
                    "Left Lower Parotid And Left Level Ii",
                    "Left Parotid And Neck Nodes",
                ]
                or subtype in ["Left Parotid And Neck Nodes"]
                else "Neck"
                if diagnosis
                in [
                    "Left Neck",
                    "Left Neck Mass",
                    "Left Radical Neck Dissection",
                    "Left Side Of Neck",
                    "Neck: Left Level 5A",
                    "Posterior Neck",
                    "Right Neck",
                    "Right Posterior Neck",
                    "Ln, Neck",
                ]
                else "Neck (NOS)"
                if (
                    diagnosis in ["Neck"]
                    and subtype
                    in [
                        "Squamous Cell Carcinoma (Head And Neck)",
                        "Scc",
                        "Centrocytic Lymphoma, Recurrent",
                        "Basaloid Squamous Cell Carcinoma (Oropharyngeal)",
                    ]
                )
                else "Tonsil"
                if diagnosis
                in [
                    "Left Tonsil",
                    "Left Tonsillar Mass",
                    "Normal Tonsil",
                    "Right Palatine Tonsil",
                    "Right Tonsil",
                    "Tonsil",
                    "Tonsil Tumour",
                    "Tonsil-Normal",
                    "Tonsils",
                ]
                else "Occiput"
                if diagnosis in ["Occiput"]
                else "Face - Salivary Gland"
                if diagnosis in ["Orbit, Maxilla, Surrounding Tissues"]
                else "Thyroid"
                if diagnosis
                in ["Thyroid", "Thyroid Ca.", "Thyroid Cancer", "Thyroid Nor."]
                else "Tonsil - Pharynx"
                if diagnosis in ["Tonsil, Lateral Pharynx"]
                # Other cases
                else "Pharynx"
                if (
                    diagnosis in ["Head & Neck Scc", "Squamous Cell Carcinoma"]
                    and subtype
                    in [
                        "Oropharynx Cancer",
                        "Hipopharynx Carcinoma",
                        "Hypopharynx Carcinoma",
                        "Pharynx Carcinoma",
                        "Nasopharynx",
                        "Hypopharynx",
                        "Orpopharynx_Matched Met",
                    ]
                )
                else "Larynx"
                if (
                    diagnosis in ["Head & Neck Scc", "Head And Neck", "Head & Neck"]
                    and subtype in ["Larynx Carcinoma", "Larynx"]
                )
                else "Oral Cavity"
                if (
                    diagnosis
                    in ["Head & Neck Scc", "Head And Neck", "Hen Pri. (Tongue)"]
                    and subtype
                    in ["Oral Cavity Cancer", "Tongue", "Oral Cavity  Cancer"]
                )
                else "Nose"
                if (diagnosis in ["Head & Neck Scc"] and subtype in ["Nose Cancer"])
                else "Ear"
                if (diagnosis in ["Head & Neck Scc"] and subtype in ["Pinna Cancer"])
                else "Brain"
                if (diagnosis in ["Hen Pri."] and subtype in ["Gbm"])
                else "Pharynx"
                if (
                    diagnosis in ["Head & Neck", "Head And Neck"]
                    and subtype in ["Hypopharynx", "Nasopharynx", "Hyphphrynx"]
                )
                else "Pharynx"
                if (
                    diagnosis in ["Left Radical Neck Dissection"]
                    and subtype in ["Squamous Cell Carcinoma (Pharyngeal Primary)"]
                )
                else "Head & Neck (NOS)"
            ),
            anatomical_sublocation=(
                "Brain (Left Frontal Space)"
                if diagnosis in ["Brain (Left Frontal Space )"]
                else "Brain (Cerebellum)"
                if diagnosis in ["Cerebellum Nor."]
                else "Brain (Cerebrum)"
                if diagnosis in ["Cerebrum", "Cerebrum Nor.", "Cerebrum. Nor"]
                or subtype in ["Cerebrum"]
                else "Mandible (Anterior)"
                if diagnosis in ["Anterior Mandible"]
                else "Mandible (Right Lingual)"
                if diagnosis in ["Right Lingual Mandible"]
                else "Tongue (Anterior)"
                if diagnosis in ["Anterior Tongue", "Anterior Tongue."]
                else "Tongue (Left Lateral)"
                if diagnosis in ["Left Lateral Tongue"]
                else "Tongue (Right Lateral)"
                if diagnosis in ["R Lateral Tongue", "Right Lateral Tongue"]
                else "Tongue (Right)"
                if diagnosis in ["R Tongue", "Right Tongue"]
                else "Tongue (Right Posterolateral)"
                if diagnosis in ["Right Posterolateral Tongue"]
                else "Tongue (Right Ventral)"
                if diagnosis in ["R Ventral Tongue", "Right Ventral Tongue"]
                else "Tongue (Base)"
                if diagnosis in ["Base Of Tongue", "H&N Pri. Tongue Base"]
                else "Tongue (Right Base)"
                if diagnosis in ["Right Tongue Base"]
                else "Soft Palate (Left) And Piriform Fossa (Left)"
                if diagnosis in ["D- Left Soft Palate, B-Left Piriform Fossa"]
                else "Mouth (Floor)"
                if diagnosis in ["Floor Of Mouth"]
                else "Mouth (Floor) And Tongue"
                if diagnosis in ["Floor Of Mouth And Tongue"]
                else "Mouth (Right Floor) And Tongue"
                if diagnosis in ["Right Floor Of Mouth/Tongue"]
                else "Gum"
                if diagnosis in ["Gum"]
                else "Lip (Right Upper)"
                if diagnosis in ["Right Upper Lip"]
                else "Tongue (Lateral Half)"
                if diagnosis in ["H&N Pri. Hemiglossectomy"]
                else "Tongue (Lateral Part)"
                if diagnosis in ["H&N Pri. Lateral Partial Glossectomy"]
                else "Tongue (Part)"
                if diagnosis in ["H&N Pri. Partial Glossectomy"]
                else "Tongue (Part Of Right Side)"
                if diagnosis in ["Right Partial Glossectomy"]
                else "Tongue (Part Of Left Side)"
                if diagnosis
                in ["H&N Pri. Left Partial Glossectomy", "Left Partial Glossectomy"]
                else "Tongue (Lateral Half) And Neck (Lymph Nodes)"
                if diagnosis
                in [
                    "H&N Pri. Hemiglossectomy  And Neck Dissection",
                    "H&N Pri. Hemiglossectomy And Neck Dissection",
                    "H&N Pri. Hemiglossectomy And Neck Dissections",
                ]
                else "Tongue (Lateral Half) And Neck (Right - Lymph Nodes)"
                if diagnosis
                in [
                    "H&N Pri. Hemiglossectomy And Right Neck Dissection",
                    "H&N Pri. Hemiglossectomy, Right Neck Dissection",
                ]
                else "Tongue (Left)"
                if diagnosis in ["Left Tongue"]
                else "Tongue (Right Lateral Half)"
                if diagnosis in ["Right Hemiglossectomy"]
                else "Tongue (Right Lateral Half) And Surrounding Tissue"
                if diagnosis in ["R Extended Hemiglossectomy"]
                else "Lip"
                if diagnosis in ["Lip"]
                else "Mouth"
                if diagnosis in ["Mouth"]
                else "Tongue (Left) And Mandible"
                if diagnosis in ["Left Tongue And Mandible"]
                else "Tongue (Lateral Half), Mouth (Floor), And Neck (Both Sides - Lymph Nodes)"
                if diagnosis
                in [
                    "H&N Pri. Hemiglossectomy With Adjacent Floor Of Mouth And Bilateral Selective Neck Dissections"
                ]
                else "Tongue (Left) And Neck (Left - Lymph Nodes)"
                if diagnosis
                in [
                    "H&N Pri. Left Sub-Total Glossectomy And Left Selective Neck Dissection"
                ]
                else "Tongue (Part) And Neck (Left - Lymph Nodes)"
                if diagnosis
                in ["H&N Pri. Partial Glossectomy And Left Selective Neck Dissection"]
                else "Tongue (Part) And Neck (Lymph Nodes)"
                if diagnosis in ["H&N Pri. Partial Glossectomy And Neck Dissection."]
                else "Tongue (Lateral Half) And Piriform Fossa (Left)"
                if diagnosis in ["H&N Pri. Hemiglossectomy And Left Piriform Fossa"]
                else "Maxilla (Anterior)"
                if diagnosis in ["Anterior Maxilla"]
                else "Maxilla"
                if diagnosis in ["Maxilla"]
                else "Maxilla (Left)"
                if subtype in ["Left Maxilla"]
                else "Ear (Left)"
                if diagnosis in ["Basal Cell Carcinoma - Left Ear"]
                else "Ear (Right Post Auricular)"
                if diagnosis
                in [
                    "Basal Cell Carcinoma - Right Post Auricular",
                    "Right Post Auricular",
                ]
                else "Cheek"
                if diagnosis in ["Cheek"]
                else "Cheek (Right Upper)"
                if diagnosis
                in ["Basal Cell Carcinoma - Right Upper Cheek", "Righ Upper Cheek"]
                else "Cheek (Left)"
                if diagnosis in ["Left Cheek"] or subtype in ["Left Cheek"]
                else "Cheek (Right)"
                if diagnosis in ["Right Bucca", "Right Cheek"]
                or subtype in ["Right Cheek"]
                else "Cheek (Right Sulcus)"
                if diagnosis in ["Right Buccal Sulcus"]
                else "Scalp (Vertex)"
                if diagnosis
                in ["Basal Cell Carcinoma - Vertex Of Scalp", "Vertex Of Scalp"]
                else "Scalp (Left Parietal)"
                if diagnosis in ["Left Parietal Scalp"]
                else "Scalp (Left Vertex)"
                if diagnosis in ["Left Vertex Scalp"]
                else "Scalp (Occipital)"
                if diagnosis in ["Occipital Scalp"]
                else "Scalp (Left)"
                if subtype in ["Left Scalp"]
                else "Scalp (Right)"
                if subtype in ["Right Scalp"]
                else "Scalp (Right Posterior Parietal)"
                if diagnosis in ["Right Posterior Parietal Scalp"]
                else "Oropharynx"
                if diagnosis in ["Pharynx"]
                and subtype in ["Basaloid Squamous Cell Carcinoma"]
                else "Oropharynx (Left)"
                if diagnosis in ["Left Oropharynx"]
                else "Oropharynx (Right)"
                if diagnosis in ["Right Lateral Oropharyngectomy", "Right Oropharynx"]
                else "Gum/Parotid Gland"
                if diagnosis in ["Gum/Parotid Gland"]
                else "Larynx And Neck (Both Sides - Lymph Nodes)"
                if diagnosis
                in [
                    "H&N Pri. Total Laryngectomy And Bilateral Selective Neck Dissections"
                ]
                else "Oropharynx"
                if diagnosis in ["Oropharynx Cancer", "Oropharynx"]
                else "Hypopharynx"
                if diagnosis
                in ["Hipopharynx Carcinoma", "Hypopharynx", "Hypopharynx Carcinoma"]
                else "Nasopharynx"
                if diagnosis in ["Nasopharynx", "Npc"]
                else "Tongue"
                if (diagnosis in ["Head & Neck Scc"] and subtype in ["Tongue"])
                else "Tongue"
                if "Tongue" in diagnosis
                else "Ear (Pinna)"
                if (diagnosis in ["Head & Neck Scc"] and subtype in ["Pinna Cancer"])
                else "Soft Palate (Left)"
                if diagnosis in ["Left Soft Palate"]
                else "Soft Palate (Right)"
                if diagnosis in ["Right Soft Palate"]
                else "Soft Palate"
                if diagnosis in ["Soft Palate"]
                else "Chin (Left)"
                if diagnosis in ["Left Chin"]
                else "Jaw (Right Tumour Deposit)"
                if diagnosis in ["Right Jaw Tumour Deposit"]
                else "Tongue, Mandible, And Neck"
                if diagnosis in ["Tongue/Mandible/Neck"]
                else "Tongue (Base) And Piriform Fossa (Left)"
                if diagnosis
                in [
                    "Tongue Base; Involves L Pyriform Fossa But Tumour Appears To Originate In Tongue Base"
                ]
                else "Ear"
                if diagnosis in ["Ear"]
                else "Ear (Left)"
                if diagnosis in ["Left Ear"]
                else "Ear (Posterior Right)"
                if diagnosis in ["Posterior Right Ear"]
                else "Ear (Right Auricular)"
                if diagnosis in ["Right Auricular"]
                else "Ear (Right)"
                if diagnosis in ["Right Ear"]
                else "Nose (Left Nasal Cavity)"
                if diagnosis in ["Left Nasal Cavity"]
                else "Nose (Right Nasal Cavity)"
                if diagnosis in ["Right Nasal Cavity"]
                else "Nose (Right Nasal Cavity) And Hard/Soft Palate (Left)"
                if diagnosis
                in [
                    "Right Nasal Cavity (Primary Tumours In R Nasal Cavity And L Hard/Soft Palate)"
                ]
                else "Temple (Superior)"
                if diagnosis
                in ["Basal Cell Carcinoma - Superior Temple", "Superior Temple"]
                else "Temple (Left)"
                if diagnosis in ["Left Temple", "Left Temporal"]
                else "Laryngopharynx"
                if diagnosis in ["Laryngopharynx"]
                else "Pharynx (Right Wall)"
                if diagnosis
                in [
                    "Right Pharyngeal Wall (Main Tumour In Supraglottis Although Tumour And High Grade Dysplasia Also Found In R Pharyngeal Wall And Tongue Base)"
                ]
                else "Oropharynx"
                if (
                    diagnosis
                    in [
                        "Basaloid Squamous Cell Carcinoma",
                        "Keratinising Squamous Cell Carcinoma",
                    ]
                    and subtype in ["Orpopharynx_Matched Met"]
                )
                or subtype in ["Basaloid Squamous Cell Carcinoma (Oropharyngeal)"]
                else "Vocal Cord"
                if diagnosis in ["Vocal Cord (Tumour Originated In L Pyriform Fossa)"]
                else "Parotid Gland (Left Lower) And Level II Lymph Nodes (Left)"
                if diagnosis in ["Left Lower Parotid And Left Level Ii"]
                else "Parotid Gland (Left)"
                if diagnosis in ["Left Parotid", "Left Parotid And Neck Nodes"]
                else "Parotid Gland"
                if diagnosis in ["Parotid", "Parotid Gland", "Parotid Galnd"]
                else "Parotid Gland (Left) And Neck (Lymph Nodes)"
                if subtype in ["Left Parotid And Neck Nodes"]
                else "Neck (Left)"
                if diagnosis in ["Left Neck", "Left Side Of Neck"]
                else "Neck (Right)"
                if diagnosis in ["Right Neck"]
                else "Neck (Right Posterior)"
                if diagnosis in ["Right Posterior Neck"]
                else "Neck (Left Mass)"
                if diagnosis in ["Left Neck Mass"]
                else "Neck (Left - Lymph Nodes) And Surroundings"
                if diagnosis in ["Left Radical Neck Dissection"]
                else "Neck (Posterior)"
                if diagnosis in ["Posterior Neck"]
                else "Neck (Left - Level 5A)"
                if diagnosis in ["Neck: Left Level 5A"]
                else "Neck (Lymph Nodes)"
                if diagnosis in ["Ln, Neck"]
                else "Tonsil (Left)"
                if diagnosis in ["Left Tonsil"]
                else "Tonsil (Left Mass)"
                if diagnosis in ["Left Tonsillar Mass"]
                else "Tonsil (Right Palatine)"
                if diagnosis in ["Right Palatine Tonsil"]
                else "Tonsil (Right)"
                if diagnosis in ["Right Tonsil"]
                else "Orbit, Maxilla, And Surrounding Tissues"
                if diagnosis in ["Orbit, Maxilla, Surrounding Tissues"]
                else "Tonsil And Pharynx (Lateral)"
                if diagnosis in ["Tonsil, Lateral Pharynx"]
                else "Tonsils (Both)"
                if diagnosis in ["Tonsils"]
                else "Oropharynx"
                if diagnosis in ["Head & Neck Scc", "Squamous Cell Carcinoma"]
                and subtype in ["Oropharynx Cancer", "Orpopharynx_Matched Met"]
                else "Hypopharynx"
                if diagnosis
                in [
                    "Head & Neck Scc",
                    "Squamous Cell Carcinoma",
                    "Head & Neck",
                    "Head And Neck",
                ]
                and subtype
                in [
                    "Hipopharynx Carcinoma",
                    "Hypopharynx Carcinoma",
                    "Hypopharynx",
                    "Hyphphrynx",
                ]
                else "Nasopharynx"
                if diagnosis
                in [
                    "Head & Neck Scc",
                    "Squamous Cell Carcinoma",
                    "Head & Neck",
                    "Head And Neck",
                ]
                and subtype in ["Nasopharynx"]
                else "Tongue"
                if diagnosis
                in ["Head & Neck Scc", "Head And Neck", "Hen Pri. (Tongue)"]
                and subtype in ["Tongue"]
                else "Supraglottic Squamous Cell Carcinoma"
                if subtype in ["Supraglottic Squamous Cell Carcinoma"]
                else "Forehead (Right)"
                if subtype in ["Right Forehead"]
                else "NR"
            ),
            specimen_type=(
                "M"
                if (
                    any(keyword in diagnosis for keyword in ["Nodal Met", "Metastasis"])
                    or any(
                        keyword in subtype
                        for keyword in [
                            "Met Blocks",
                            "Matched Met",
                            "Metastatic",
                            "Metched Met",
                        ]
                    )
                )
                and diagnosis not in "Hen Pri. And Nodal Met."
                else "N"
                if (
                    any(
                        keyword in diagnosis
                        for keyword in [
                            "Nor.",
                            "Normal",
                        ]
                    )
                    or any(keyword in subtype for keyword in ["Normal", "Nor."])
                )
                else "NR"
                if diagnosis
                in [
                    "Brain Tissue",
                    "Head&Neck",
                    "Head And Neck",
                    "Parotid Gland",
                    "Head & Neck",
                ]
                and subtype
                in [
                    "Unknown",
                    "Not Stated",
                    "-",
                    "Larynx",
                    "Hypopharynx",
                    "Hyphphrynx",
                    "Tongue",
                    "Nasopharynx",
                    "Parotid Gland",
                    " ",
                ]
                else "P"
            ),
            diagnosis=(
                "Glioblastoma"
                if (
                    diagnosis in ["Gbm", "Gbm Pre-Treatment", "Glioblastoma Cancer"]
                    or subtype
                    in ["Glioblastoma (Primary)", "Glioblastoma", "Gbm", "Gliobastoma"]
                )
                else "Gliosarcoma"
                if (subtype in ["Gliosarcoma", "Gliosarcoma (Wild Type Idh-1 )"])
                else "Squamous Cell Carcinoma"
                if (
                    subtype
                    in [
                        "Scc",
                        "Squamous Cell Carcinoma",
                        "Basaloid Squamous Cell Carcinoma",
                        "Recurrent Scc",
                        "Residual Scc",
                        "Squamous Carcinoma",
                        "Squamous Cell Carcinoma (Head And Neck)",
                        "Squamous Cell Carcinoma (Laryngeal)",
                        "Squamous Cell Carcinoma (Hypopharyngeal)",
                        "Supraglottic Squamous Cell Carcinoma",
                        "Squamous Cell Carcinoma With Basaloid Features",
                        "Squamous Cell Carcinoma Basaloid Features And Clear Cell Change",
                        "Squamous Cell Carcinoma Maxilla",
                        "Squamous Cell Carcinoma, H&N",
                        "Basaloid Squamous Cell Carcinoma (Oropharyngeal)",
                        "Keratinising Squamous Cell Carcinoma",
                        "Spindle Cell Squamous Cell Carcinoma",
                        "Squamous Cell Carcinoma Tonsil",
                        "Squamous Cell Carcinoma Oropharynx",
                        "Orpopharynx_Matched Met",
                        "Squamous Cell Carcinoma (Pharyngeal Primary)",
                        "Metastatic Squamous Cell Carcinoma Maxilla (Lh18-18903)",
                        "Metastatic Scc",
                        "Squamous Cell Carcinoma (Unknown Primary, Likely Oropharyngeal)",
                        "Metastatic/Recurrent Basaloid Scc",
                        "Metastatic/Recurrent Scc With Basaloid Features",
                        "Squamous Cell Carcinoma Maxillia",
                    ]
                    or diagnosis
                    in [
                        "Cscc",
                        "Head & Neck Scc",
                        "Npc",
                        "Scc",
                        "Right Facial Lymph Node",
                    ]
                )
                and "Normal" not in subtype
                else "Basal Cell Carcinoma"
                if subtype
                in [
                    "Bcc",
                    "Basal Cell Carcinoma Of Superficial, Nodular And Focal Infiltrative Type",
                ]
                else "Astrocytoma"
                if subtype in ["Astrocytoma"]
                else "Oligodendroglioma"
                if subtype in ["Oligodendroglioma"]
                else "Astroglial"
                if subtype
                in ["Astroglial"]  # subtype should be astrocytoma (needs verification)
                else "Large Cell Neuroendocrine Carcinoma"
                if subtype in ["Large Cell Neuroendocrine Carcinoma"]
                else "Follicular Lymphoma"
                if subtype in ["Fl"]
                else "Melanoma"
                if diagnosis in ["Malignant Melanoma", "Mucosal Melanoma"]
                or subtype in ["Metastatic Melanoma, Nos"]
                else "Inverted Papilloma"
                if subtype in ["Inverted Papilloma"]
                else "Papillary Carcinoma"
                if subtype
                in [
                    "Ptc",
                    "Papillary Carcinoma",
                    "Papillary Thyroid Carcinoma",
                    "Multifocal Papillary Thyroid Carcinoma",
                    "Cystic Papillary Carcinoma",
                    "Multifocal Papillary Thyroid",
                    "Partially Cystic Papillary Thyroid Carcinoma",
                    "Papillary Thyroid Carcinoma, Pleomorphic Salivary Adenoma.",
                    "Papillary Ca.",
                    "Metastatic Papillary Thyroid Carcinoma",
                    "Papillary Thyroid Carcinoma, Metastatic Papillary Thyroid Carcinoma",
                ]
                else "Medullary Carcinoma"
                if subtype
                in [
                    "Medullary Carcinoma",
                    "Medullary Thyroid Carcinoma",
                    "Medullary Thyroid Carcinoma And Focal C-Cell Hyperplasia",
                ]
                else "Follicular Carcinoma"
                if subtype in ["Ftc"]
                else "Anaplastic Carcinoma"
                if subtype in ["Anaplastic Ca."]
                else "Hurthle Cell Carcinoma"
                if subtype in ["Hurthle Cell Carcinoma"]
                else "Adenoma And Carcinoma"
                if subtype
                in ["Hurthle Cell Follicular Adenoma And Papillary Thyroid Carcinoma"]
                else "No Cancer"
                if (
                    any(
                        keyword in diagnosis
                        for keyword in [
                            "Nor.",
                            "Normal",
                        ]
                    )
                    or any(keyword in subtype for keyword in ["Normal", "Nor."])
                )
                else "Renal Cell Carcinoma"
                if subtype
                in ["Metastatic Clear Cell Carcinoma, Consistent With Renal Origin"]
                else "Adenocarcinoma"
                if subtype in ["Metastatic Lung Adenocarcinoma"]
                else "Peripheral T-Cell Lymphoma"
                if subtype in ["Angioimmunoblastic T-Cell L"]
                else "Peripheral B-Cell Lymphoma"
                if subtype
                in [
                    "Dlbcl",
                    "Centrocytic Lymphoma, Recurrent",
                    "Centrocytic Lymphoma, B Cell",
                ]
                else "Carcinoma"
                if (
                    diagnosis
                    in ["Pharynx Cancer With Ln Neck Metastasis", "Thyroid Cancer"]
                    and subtype in ["-"]
                )
                else "Marginal Zone Lymphoma"
                if subtype in ["Mzl"]
                else "Cancer"
                if diagnosis in ["Hen", "Hen Pri. (Tongue)"]
                and subtype in ["Primary", "Tongue"]
                else "NR"
            ),
            subtype=(
                "Astrocytoma (NOS)"
                if subtype in ["Astrocytoma"]
                else "Astrocytoma (Needs Verification)"
                if subtype in ["Astroglial"]
                else "Gliosarcoma (Wild Type IDH-1)"
                if subtype in ["Gliosarcoma (Wild Type Idh-1 )"]
                else "Squamous Cell Carcinoma With Basaloid Features"
                if subtype in ["Squamous Cell Carcinoma With Basaloid Features"]
                else "Basaloid Squamous Cell Carcinoma"
                if subtype
                in [
                    "Basaloid Squamous Cell Carcinoma",
                    "Basaloid Squamous Cell Carcinoma (Oropharyngeal)",
                ]
                or diagnosis in ["Basaloid Squamous Cell Carcinoma"]
                else "Squamous Cell Carcinoma With Basaloid Features And Clear Cell Change"
                if subtype
                in ["Squamous Cell Carcinoma Basaloid Features And Clear Cell Change"]
                else "No Cancer"
                if (
                    any(
                        keyword in diagnosis
                        for keyword in [
                            "Nor.",
                            "Normal",
                        ]
                    )
                    or any(keyword in subtype for keyword in ["Normal", "Nor."])
                )
                else "Squamous Cell Carcinoma (Recurrent)"
                if subtype in ["Recurrent Scc"]
                else "Multifocal Papillary Thyroid Carcinoma"
                if subtype
                in [
                    "Multifocal Papillary Thyroid Carcinoma",
                    "Multifocal Papillary Thyroid",
                ]
                else "Cystic Papillary Carcinoma"
                if subtype
                in [
                    "Cystic Papillary Carcinoma",
                    "Partially Cystic Papillary Thyroid Carcinoma",
                ]
                else "Medullary Carcinoma And Focal C-Cell Hyperplasia"
                if subtype
                in ["Medullary Thyroid Carcinoma And Focal C-Cell Hyperplasia"]
                else "Papillary Carcinoma And Pleomorphic Salivary Adenoma"
                if subtype
                in ["Papillary Thyroid Carcinoma, Pleomorphic Salivary Adenoma."]
                else "Superficial, Nodular, And Focal Infiltrative Basal Cell Carcinoma"
                if subtype
                in [
                    "Basal Cell Carcinoma Of Superficial, Nodular And Focal Infiltrative Type"
                ]
                else "Basaloid Squamous Cell Carcinoma (Recurrent)"
                if subtype in ["Metastatic/Recurrent Basaloid Scc"]
                else "Squamous Cell Carcinoma With Basaloid Features (Recurrent)"
                if subtype in ["Metastatic/Recurrent Scc With Basaloid Features"]
                else "Squamous Cell Carcinoma (Residual)"
                if subtype in ["Residual Scc"]
                else "Keratinising Squamous Cell Carcinoma"
                if subtype in ["Keratinising Squamous Cell Carcinoma"]
                or diagnosis in ["Keratinising Squamous Cell Carcinoma"]
                else "Spindle Cell Squamous Cell Carcinoma"
                if subtype in ["Spindle Cell Squamous Cell Carcinoma"]
                else "Squamous Cell Carcinoma (Cutaneous)"
                if diagnosis in ["Cscc"]
                else "Supraglottic Squamous Cell Carcinoma"
                if subtype in ["Supraglottic Squamous Cell Carcinoma"]
                else "Angioimmunoblastic T-Cell Lymphoma"
                if subtype in ["Angioimmunoblastic T-Cell L"]
                else "Diffuse Large B-Cell Lymphoma"
                if subtype in ["Dlbcl"]
                else "Centrocytic B-Cell Lymphoma (Recurrent)"
                if subtype in ["Centrocytic Lymphoma, Recurrent"]
                else "Centrocytic Lymphoma, B Cell"
                if subtype in ["Centrocytic Lymphoma, B Cell"]
                else "NR"
            ),
            pairing_status=(
                "P"
                if any(
                    keyword in subtype
                    for keyword in [
                        "Matched",
                        "Metched",
                        "Metastatic Lung Adenocarcinoma",
                    ]
                )
                and "Normal" not in subtype
                else "P"
                if (
                    diagnosis == "Left Radical Neck Dissection"
                    and subtype == "Squamous Cell Carcinoma (Pharyngeal Primary)"
                )
                else "P"
                if (
                    diagnosis == "Brain" and subtype == "Metastatic Lung Adenocarcinoma"
                )
                else "P"
                if (
                    diagnosis == "Right Neck"
                    and subtype
                    == "Squamous Cell Carcinoma (Unknown Primary, Likely Oropharyngeal)"
                )
                else "P"
                if (
                    diagnosis == "Left Lower Parotid And Left Level Ii"
                    and subtype == "Squamous Cell Carcinoma Maxillia"
                )
                else "P"
                if (
                    diagnosis == "Brain"
                    and subtype
                    == "Metastatic Clear Cell Carcinoma, Consistent With Renal Origin"
                )
                else "P"
                if (
                    diagnosis == "Mucosal Melanoma"
                    and subtype == "Left Parotid And Neck Nodes"
                )
                else "P"
                if (
                    diagnosis == "Pharynx Cancer With Ln Neck Metastasis"
                    and subtype == "-"
                )
                else "P"
                if (subtype == "Metastatic Melanoma, Nos")
                else "No Cancer"
                if (
                    any(
                        keyword in diagnosis
                        for keyword in [
                            "Nor.",
                            "Normal",
                        ]
                    )
                    or any(keyword in subtype for keyword in ["Normal", "Nor."])
                )
                else "NR"
            ),
            pairing_site=(
                "Lung"
                if subtype == "Metastatic Lung Adenocarcinoma"
                else "Pharynx"
                if subtype == "Squamous Cell Carcinoma (Pharyngeal Primary)"
                else "Oropharynx"
                if subtype
                in [
                    "Squamous Cell Carcinoma (Unknown Primary, Likely Oropharyngeal)",
                    "Basaloid Squamous Cell Carcinoma (Oropharyngeal)",
                ]
                else "Maxilla"
                if subtype == "Squamous Cell Carcinoma Maxillia"
                else "Kidney"
                if subtype
                == "Metastatic Clear Cell Carcinoma, Consistent With Renal Origin"
                else "Lymph Nodes"
                if (
                    diagnosis == "Pharynx Cancer With Ln Neck Metastasis"
                    and subtype == "-"
                )
                else "Lymph Nodes (Neck)"
                if subtype == "Left Parotid And Neck Nodes"
                else "No Cancer"
                if (
                    any(
                        keyword in diagnosis
                        for keyword in [
                            "Nor.",
                            "Normal",
                        ]
                    )
                    or any(keyword in subtype for keyword in ["Normal", "Nor."])
                )
                else "Oral Cavity"
                if subtype == "Oral Cavity_Matched Met"
                else "NA"
            ),
            i=i,
        )


# Export to Box & Load it to format the column width
head_and_neck, head_and_neck_inventory = format_and_export("H&N", "Head & Neck")


# In[23]: Gallbladder (Verified)

# Create the combinations dictionary for Gallbladder
create_combinations_dict("Gallbladder")

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Gallbladder":
        assign_values(
            anatomical_location="Gallbladder",
            anatomical_sublocation=(
                "Gallbladder (Invading Beyond Serosa)"
                if subtype
                in [
                    "Biliary Adc Invading Beyond Serosa",
                    "Biliary Adenocarcinoma Invading Beyond Serosa",
                ]
                else "Gallbladder (Invading Perimuscular Connective Tissue - Not Beyond Serosa)"
                if subtype
                in [
                    "Biliary Adenocarcinoma Invading Perimuscular Connective Tissue But Not Beyond Serosa"
                ]
                else "Gallbladder (Infiltrating Muscular And Serous Layers And Surrounding Fat Tissue)"
                if subtype
                in [
                    "Moderately \xa0Differentited Tumor Infiltrating Muscular- And Serous Layers And Surrounding Fat Tissue"
                ]
                else "NR"
            ),
            specimen_type=(
                "N"
                if subtype == "Normal"
                else "NR"
                if diagnosis == "Gallbladder" and subtype == "-"
                else "P"
            ),
            diagnosis=(
                "Adenocarcinoma"
                if subtype
                in [
                    "Adenocarcinoma",
                    "Adc",
                    "Adc Intestinal",
                    "Adc \xa0Biliary Type",
                    "Adc  Biliary Type",
                    "Biliary Adc Invading Beyond Serosa",
                    "Biliary Adenocarcinoma Invading Perimuscular Connective Tissue But Not Beyond Serosa",
                    "Biliary Adenocarcinoma Invading Beyond Serosa",
                    "Adc Biliary Type",
                    "Adc \xa0Biliary Type",
                    "Adenocarcinoma - Billary\n Type",
                    "Adenocarcinoma, Intestinal Type",
                ]
                else "Mixed Adenocarcinoma"
                if subtype == "Adenocarcinoma, Mixed Intestinal/Billary Type"
                else "Tumor"
                if subtype
                == "Moderately \xa0Differentited Tumor Infiltrating Muscular- And Serous Layers And Surrounding Fat Tissue"
                else "Squamous Cell Carcinoma"
                if subtype == "Squamous Cell Carcinoma"
                else "No Cancer"
                if subtype == "Normal"
                else "Carcinoma"
                if subtype
                == "Moderately  Differentited Tumor Infiltrating Muscular- And Serous Layers And Surrounding Fat Tissue"
                else "NR"
            ),
            subtype=(
                "Intestinal Adenocarcinoma"
                if subtype in ["Adc Intestinal", "Adenocarcinoma, Intestinal Type"]
                else "Moderately Differentiated Tumor"
                if subtype
                == "Moderately \xa0Differentited Tumor Infiltrating Muscular- And Serous Layers And Surrounding Fat Tissue"
                else "No Cancer"
                if subtype == "Normal"
                else "Intestinal And Biliary Adenocarcinoma"
                if subtype == "Adenocarcinoma, Mixed Intestinal/Billary Type"
                else "Carcinoma (NOS)"
                if subtype
                == "Moderately  Differentited Tumor Infiltrating Muscular- And Serous Layers And Surrounding Fat Tissue"
                else "Biliary Adenocarcinoma"
                if subtype
                in [
                    "Adc  Biliary Type",
                    "Adc Biliary Type",
                    "Adenocarcinoma - Billary\n Type",
                ]
                else "Biliary Adenocarcinoma (Invading Beyond Serosa)"
                if subtype in ["Biliary Adc Invading Beyond Serosa"]
                else "Biliary Adenocarcinoma (Invading Beyond Perimuscular Connective Tissue)"
                if subtype
                in [
                    "Biliary Adenocarcinoma Invading Perimuscular Connective Tissue But Not Beyond Serosa"
                ]
                else "NR"
            ),
            pairing_status=("No Cancer" if subtype == "Normal" else "NR"),
            pairing_site=("No Cancer" if subtype == "Normal" else "NR"),
            i=i,
        )


# Export to Box & Load it to format the column width
gallbladder, gallbladder_inventory = format_and_export(
    file_name="gallbladder", tissue_source="Gallbladder"
)


# In[24]: Intestines (Verified)

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Intestines":
        assign_values(
            anatomical_location=(
                "Small Intestine"
                if (
                    diagnosis
                    in [
                        "Small Intestine",
                        "Small Bowel",
                        "Ileum",
                        "Ileum Nor.",
                        "Ileum  Nor.",
                        "Moderately Differentiated Adenocarcinoma Duodenum",
                        "Poorly Differentiated Adenocarcinoma Of Duodenum",
                        "Duodenal Adenocarcinoma",
                        "Adenocarcinoma Of The Duodenum",
                        "Poorly Differentiated Focally Mucinous Adenocarcinoma Of The Duodenum",
                        "Duodenum",
                    ]
                    or subtype in ["Ileo-Caecal Valve", "Duodenum"]
                )
                and "Caecal" not in subtype
                else "Small Intestine And Colon"
                if diagnosis in ["Small Intestine And Colon", "Small Bowel And Colon"]
                else "Ileum And Caecum"
                if diagnosis in ["Ileum & Caecum"] or subtype in ["Ileo-Caecal Valve"]
                else "Rectum"
                if diagnosis in ["Rectum"]
                or subtype
                in [
                    "Lower Rectum",
                    "Rectum",
                    "Rectum_Adenocarcinoma",
                    "Rectum_Mucinous Adenocarcinoma",
                    "Rectum_Signet Ring Cell Adenocarcinoma",
                ]
                else "Colon And Rectum"
                if subtype
                in [
                    "Left Colon And Upper Rectum - Adenocarcinoma",
                    "Rectum And Sigmoid",
                    "Sigmoid And Rectum",
                ]
                else "Colon"
                if diagnosis
                in [
                    "Hepatic Flexure",
                    "Ascending Colon",
                    "Colon",
                    "Colon  Nor.",
                    "Colon (Ascending)",
                    "Colon (Non Malignant)",
                    "Colon (Sigmoid)",
                    "Colon (Transverse)",
                    "Colon Descendens",
                    "Colon Pri.",
                    "Colon Pri. And Met.",
                    "Colon Pri. And Met. (Liver)",
                    "Crohns Colon",
                    "Crohn'S Disease",
                    "Left (Descending) Colon",
                    "Normal Colon",
                    "Right (Ascending) Colon",
                    "Right (Ascending) Colon/Transverse Colon",
                    "Sigmoid",
                    "Sigmoid Colon",
                    "Transverse Colon",
                    "Ulcerative Colitis",
                    "Ulcerative Colitis Colon",
                    "Colon Nor.",
                    "Metastatic Renal Cell Carcinoma",
                ]
                or subtype
                in [
                    "Hepatic Flexure",
                    "Splenic Flexure",
                    "Hepatic Flexure_Mucinous Adenocarcinoma",
                    "Splenic Flexure_Adenocarcinoma",
                    "Hepatic Flexure_Adenocarcinoma",
                    "Hepatic Flexure/Ascending Colon_Adenocarcinoma",
                    "Colon",
                    "Ascending Colon",
                    "Right Hemicolectomy",
                    "Mid Transverse Colon",
                    "Sigmoid Colon",
                    "Rectosigmoid",
                    "Transverse Colon",
                    "Sigmoid Colectomy",
                    "Distal Sigmoid Colon",
                    "Mid-Descending Colon",
                    "Sigmoid",
                    "Right Colon",
                    "Mid Sigmoid Colon",
                    "Rectosigmoid Colon",
                    "Descending Colon",
                    "Descending Colon - Adenocarcinoma",
                    "Low Sigmoid Colon",
                    "Proximal Sigmoid Colon",
                    "Right Hemicolectomy Mucinous Carcinoma",
                    "Distal Transverse Colon",
                    "Sigmoid Colon - Adenocarcinoma",
                    "Distal Sigmoid",
                    "Transverse Colon_Adenocarcinoma",
                    "Sigmoid Colon_Adenocarcinoma",
                    "Right (Ascending) Colon_Adenocarcinoma",
                    "Sigmoid Colon & Transverse Colon_Adenocarcinoma",
                    "Transverse Colon_Mucinous Adenocarcinoma",
                    "Sigmoid_Adenocarcinoma",
                    "Left (Descending) Colon_Mucinous Adenocarcinoma",
                    "Left (Descending) Colon_Adenocarcinoma",
                    "Transverse Colon_Mucinous/Signet Ring Cell Adenocarcinoma",
                    "Sigmoid Colon_Mucinous Adenocarcinoma",
                    "Transverse",
                    "Ascending",
                    "Colon",
                    "Descending",
                    "Colon",
                    "Sigmoid_Mucinous Adenocarcinoma",
                    "Proximal Transverse Colon",
                ]
                else "Colon And Uterus"
                if subtype in ["Sigmoid Colon With Adhered Uterus"]
                else "Colon/Rectum"
                if diagnosis in ["Crc", "Colorectal Cancer"]
                and subtype
                in [
                    "-",
                    "Adenocarcinoma",
                    "Mucinous Adenocarcinoma",
                    "Normal Adjacent Tissues",
                    "Adjacent Normal",
                    "Nr",
                    "Adenocarcinoma With Mucinous Features",
                    "Anterior Resection",
                    "Mixed Glandular-Neuroendocrine (Composite) Carcinoma",
                    "Mucinous Carcinoma",
                    "Normal",
                    "Primary / Adenocarcinoma",
                    "Primary / Carcinoma With Medullary Features",
                    "Primary / Medullary Carcinoma",
                    "Primary / Adenocarcinoma With Medullary Features",
                    "Synchronous Adenocarcinomas",
                    "Neuroendocrine Tumour Or A Mixed Adenocarcinoma And Neurotendocrine Tumour",
                    "Synchronous Adenocarcinomas Index Tumour",
                    "",
                ]
                else "Caecum"
                if diagnosis in ["Caecum", "Cecum", "Right Caecum"]
                or subtype
                in [
                    "Caecum",
                    "Caecum Adenocarcinoma",
                    "Caecum_Adenocarcinoma",
                    "Caecum_Mucinous Adenocarcinoma",
                    "Caecum_Signet Ring Cell Adenocarcinoma",
                    "Caecum_Carcinoma",
                    "Adenocarcinoma, Caecal",
                    "Cecum",
                    "Caecal Adenocarcinoma",
                    "Right Caecum",
                ]
                else "Appendix"
                if diagnosis in ["Appendix"]
                else "Peritoneum"
                if subtype == "Peritoneum"
                else "Lung"
                if diagnosis == "Colon Adenocarcinoma Lung Metastase"
                or subtype == "Lung"
                else "Liver"
                if subtype in ["Liver", "Liver Meta"]
                else "Lymph Nodes"
                if subtype == "Ln Meta"
                else "NR"
            ),
            anatomical_sublocation=(
                "Ileum"
                if diagnosis in ["Ileum", "Ileum Nor."]
                else "Duodenum"
                if diagnosis
                in [
                    "Moderately Differentiated Adenocarcinoma Duodenum",
                    "Poorly Differentiated Adenocarcinoma Of Duodenum",
                    "Duodenal Adenocarcinoma",
                    "Adenocarcinoma Of The Duodenum",
                    "Poorly Differentiated Focally Mucinous Adenocarcinoma Of The Duodenum",
                    "Duodenum",
                ]
                or subtype in ["Duodenum"]
                else "Ileo-Caecal Valve"
                if subtype in ["Ileo-Caecal Valve"]
                else "Rectum (Lower)"
                if subtype in ["Lower Rectum"]
                else "Colon (Left) And Rectum (Upper)"
                if subtype in ["Left Colon And Upper Rectum - Adenocarcinoma"]
                else "Colon (Sigmoid) And Rectum"
                if subtype in ["Rectum And Sigmoid", "Sigmoid And Rectum"]
                else "Colon (Hepatic Flexure)"
                if diagnosis in ["Hepatic Flexure"]
                or subtype
                in [
                    "Hepatic Flexure",
                    "Hepatic Flexure_Mucinous Adenocarcinoma",
                    "Hepatic Flexure_Adenocarcinoma",
                ]
                else "Colon (Splenic Flexure)"
                if subtype in ["Splenic Flexure", "Splenic Flexure_Adenocarcinoma"]
                else "Colon (Hepatic Flexure And Ascending)"
                if subtype in ["Hepatic Flexure/Ascending Colon_Adenocarcinoma"]
                else "Colon (Ascending)"
                if diagnosis in ["Ascending Colon", "Colon (Ascending)"]
                or subtype in ["Ascending Colon", "Ascending"]
                else "Colon (Right Ascending And Transverse)"
                if diagnosis in ["Right (Ascending) Colon/Transverse Colon"]
                else "Colon (Right Ascending)"
                if diagnosis in ["Right (Ascending) Colon"]
                or subtype in ["Right (Ascending) Colon_Adenocarcinoma"]
                else "Colon (Left Descending)"
                if diagnosis in ["Left (Descending) Colon", "Colon Descendens"]
                or subtype
                in [
                    "Left (Descending) Colon_Mucinous Adenocarcinoma",
                    "Left (Descending) Colon_Adenocarcinoma",
                ]
                else "Colon (Descending)"
                if subtype
                in [
                    "Descending Colon - Adenocarcinoma",
                    "Descending",
                    "Descending Colon",
                ]
                else "Colon (Middle Descending)"
                if subtype in ["Mid-Descending Colon"]
                else "Colon (Sigmoid)"
                if diagnosis in ["Colon (Sigmoid)", "Sigmoid", "Sigmoid Colon"]
                or subtype
                in [
                    "Sigmoid Colon",
                    "Sigmoid_Adenocarcinoma",
                    "Sigmoid",
                    "Sigmoid Colon - Adenocarcinoma",
                    "Sigmoid_Adenocarcinoma",
                    "Sigmoid Colon_Mucinous Adenocarcinoma",
                    "Sigmoid_Mucinous Adenocarcinoma",
                    "Sigmoid Colon_Adenocarcinoma",
                    "Sigmoid Colectomy",
                ]
                else "Colon (Distal Sigmoid)"
                if subtype in ["Distal Sigmoid", "Distal Sigmoid Colon"]
                else "Colon (Rectosigmoid)"
                if subtype in ["Rectosigmoid", "Rectosigmoid Colon"]
                else "Colon (Low Sigmoid)"
                if subtype in ["Low Sigmoid Colon"]
                else "Colon (Sigmoid And Transverse)"
                if subtype in ["Sigmoid Colon & Transverse Colon_Adenocarcinoma"]
                else "Colon (Middle Sigmoid)"
                if subtype in ["Mid Sigmoid Colon"]
                else "Colon (Proximal Sigmoid)"
                if subtype in ["Proximal Sigmoid Colon"]
                else "Colon (Transverse)"
                if diagnosis in ["Colon (Transverse)", "Transverse Colon"]
                or subtype
                in [
                    "Transverse Colon_Adenocarcinoma",
                    "Transverse Colon_Mucinous Adenocarcinoma",
                    "Transverse Colon_Mucinous/Signet Ring Cell Adenocarcinoma",
                    "Transverse",
                    "Transverse Colon",
                ]
                else "Colon (Distal Transverse)"
                if subtype in ["Distal Transverse Colon"]
                else "Colon (Middle Transverse)"
                if subtype in ["Mid Transverse Colon"]
                else "Colon (Right)"
                if subtype
                in [
                    "Right Hemicolectomy",
                    "Right Colon",
                    "Right Hemicolectomy Mucinous Carcinoma",
                ]
                else "Colon (Proximal Transverse)"
                if subtype in ["Proximal Transverse Colon"]
                else "Colon/Rectum (Anterior Resection)"
                if subtype == "Anterior Resection"
                else "Colon (Sigmoid) And Uterus (Adhered)"
                if subtype == "Sigmoid Colon With Adhered Uterus"
                else "Caecum (Right)"
                if diagnosis == "Right Caecum"
                else "NR"
            ),
            specimen_type=(
                "N"
                if diagnosis
                in ["Normal Colon", "Colon Nor.", "Colon  Nor.", "Ileum Nor."]
                or subtype in ["Normal", "Matched Normal", "Normale", "No Tumor"]
                else "N (Adjacent)"
                if diagnosis in ["Normal Adjancent"]
                or subtype
                in ["Normal Adjacent", "Adjacent Normal", "Normal Adjacent Tissues"]
                else "M"
                if diagnosis in ["Colon Met.", "Colon Adenocarcinoma Lung Metastase"]
                or subtype
                in [
                    "Ln Meta",
                    "Mets",
                    "Matched Met / Adenocarcinoma",
                    "Matched Met / Carcinoma With Medullary Features",
                    "Matched Met / Medullary Carcinoma",
                    "Matched Met",
                    "Matched Met / Adenocarcinoma With Medullary Features",
                    "Liver Meta",
                    "Metastatic Renal Cell Carcinoma",
                    "Crc_Matched Met",
                ]
                else "NR"
                if (
                    diagnosis
                    in [
                        "Crohns Colon",
                        "Crohn'S Disease",
                        "Ulcerative Colitis Colon",
                        "Ulcerative Colitis",
                    ]
                )
                or (
                    diagnosis in ["Small Bowel", "Colon", "Rectum"]
                    and subtype in ["-", "", " ", "Colon"]
                )
                else "P"
            ),
            diagnosis=(
                "Adenocarcinoma"
                if subtype
                in [
                    "Adc",
                    "Adc (Sig. Ring)",
                    "Adc + Kras-Mutated",
                    "Adc Muc.",
                    "Adc/ Metasynchronous",
                    "Adc/ Synchronous",
                    "Adenocarcinoma",
                    "Adenocarcinoma With Mucinous Features",
                    "Adenocarcinoma, Caecal",
                    "Caecal Adenocarcinoma",
                    "Caecum_Adenocarcinoma",
                    "Caecum_Carcinoma",
                    "Caecum_Mucinous Adenocarcinoma",
                    "Caecum_Signet Ring Cell Adenocarcinoma",
                    "Descending Colon - Adenocarcinoma",
                    "Ductal Adenocarcinoma",
                    "Hepatic Flexure/Ascending Colon_Adenocarcinoma",
                    "Hepatic Flexure_Adenocarcinoma",
                    "Hepatic Flexure_Mucinous Adenocarcinoma",
                    "Left (Descending) Colon_Adenocarcinoma",
                    "Left (Descending) Colon_Mucinous Adenocarcinoma",
                    "Left Colon And Upper Rectum - Adenocarcinoma",
                    "Matched Met / Adenocarcinoma",
                    "Matched Met / Adenocarcinoma With Medullary Features",
                    "Matched Met / Carcinoma With Medullary Features",
                    "Matched Met / Medullary Carcinoma",
                    "Mucinous Adenocarcinoma",
                    "Mucinous Carcinoma",
                    "Polyp Adenocarcinoma",
                    "Primary / Adenocarcinoma",
                    "Primary / Adenocarcinoma With Medullary Features",
                    "Primary / Carcinoma With Medullary Features",
                    "Primary / Medullary Carcinoma",
                    "Rectum_Adenocarcinoma",
                    "Rectum_Mucinous Adenocarcinoma",
                    "Rectum_Signet Ring Cell Adenocarcinoma",
                    "Right (Ascending) Colon_Adenocarcinoma",
                    "Serrated Adenocarcinoma",
                    "Sigmoid Colon - Adenocarcinoma",
                    "Sigmoid Colon & Transverse Colon_Adenocarcinoma",
                    "Sigmoid Colon_Adenocarcinoma",
                    "Sigmoid Colon_Mucinous Adenocarcinoma",
                    "Sigmoid_Adenocarcinoma",
                    "Sigmoid_Mucinous Adenocarcinoma",
                    "Signet Ring Cell Carcionoma",
                    "Splenic Flexure_Adenocarcinoma",
                    "Synchronous Adenocarcinomas",
                    "Synchronous Adenocarcinomas Index Tumour",
                    "Transverse Colon_Adenocarcinoma",
                    "Transverse Colon_Mucinous Adenocarcinoma",
                    "Transverse Colon_Mucinous/Signet Ring Cell Adenocarcinoma",
                    "Right Hemicolectomy Mucinous Carcinoma",
                ]
                or diagnosis
                in [
                    "Duodenal Adenocarcinoma",
                    "Moderately Differentiated Adenocarcinoma Duodenum",
                    "Poorly Differentiated Adenocarcinoma Of Duodenum",
                    "Adenocarcinoma Of The Duodenum",
                    "Poorly Differentiated Focally Mucinous Adenocarcinoma Of The Duodenum",
                    "Colon Adenocarcinoma Lung Metastase",
                ]
                else "Adenoma"
                if subtype in ["Adenoma"]
                else "Crohn's Colon"
                if diagnosis in ["Crohns Colon", "Crohn'S Disease"]
                else "Ulcerative Colitis"
                if diagnosis in ["Ulcerative Colitis Colon", "Ulcerative Colitis"]
                else "Gastrointestinal Stromal Tumour"
                if diagnosis in ["Gist"] or subtype in ["Gist"]
                else "Peripheral B-Cell Lymphoma"
                if subtype
                in [
                    "Centrocytic Lymphoma, B Cell",
                    "Mantle Cell",
                    "Dlbcl",
                    "Dlbcl Various Site",
                ]
                else "Enteropathy-Associated T-Cell L"
                if subtype in ["Enteropathy-Associated T-Cell L"]
                else "Extranodal NK/T-Cell Lymphoma"
                if subtype in ["Extranodal Nk/T-Cell L"]
                else "Follicular Lymphoma"
                if subtype in ["Fl"]
                else "Hodgkin's Lymphoma"
                if subtype in ["Hd"]
                else "No Cancer"
                if diagnosis
                in [
                    "Normal Colon",
                    "Colon Nor.",
                    "Colon  Nor.",
                    "Ileum Nor.",
                    "Normal Adjancent",
                ]
                or subtype
                in [
                    "Normal",
                    "Matched Normal",
                    "Normale",
                    "Normal Adjacent",
                    "Adjacent Normal",
                    "Normal Adjacent Tissues",
                    "No Tumor",
                ]
                else "Renal Cell Carcinoma"
                if subtype in ["Metastatic Renal Cell Carcinoma"]
                else "Mixed (Composite) Carcinoma"
                if subtype in ["Mixed Glandular-Neuroendocrine (Composite) Carcinoma"]
                else "Mixed Carcinoma"
                if subtype
                in [
                    "Neuroendocrine Tumour Or A Mixed Adenocarcinoma And Neurotendocrine Tumour"
                ]
                else "Neuroendocrine Tumour"
                if subtype in ["Neuroendocrine Tumour"]
                else "Non-Hodgkin Lymphoma"
                if subtype in ["Nhl"]
                else "T-Cell Lymphoma"
                if subtype in ["Tcl"]
                else "Colorectal Cancer"
                if diagnosis in ["Crc", "Colorectal Cancer"]
                else "Cancer"
                if diagnosis in ["Colon Met."] and subtype in ["Peritoneum"]
                else "NR"
            ),
            subtype=(
                "Signet Ring Cell Adenocarcinoma"
                if subtype
                in [
                    "Adc (Sig. Ring)",
                    "Caecum_Signet Ring Cell Adenocarcinoma",
                    "Rectum_Signet Ring Cell Adenocarcinoma",
                    "Signet Ring Cell Carcionoma",
                ]
                else "Mucinous Adenocarcinoma"
                if subtype
                in [
                    "Adc Muc.",
                    "Caecum_Mucinous Adenocarcinoma",
                    "Hepatic Flexure_Mucinous Adenocarcinoma",
                    "Left (Descending) Colon_Mucinous Adenocarcinoma",
                    "Mucinous Adenocarcinoma",
                    "Mucinous Carcinoma",
                    "Rectum_Mucinous Adenocarcinoma",
                    "Sigmoid Colon_Mucinous Adenocarcinoma",
                    "Sigmoid_Mucinous Adenocarcinoma",
                    "Transverse Colon_Mucinous Adenocarcinoma",
                    "Right Hemicolectomy Mucinous Carcinoma",
                ]
                else "Adenocarcinoma (Metasynchronous)"
                if subtype in ["Adc/ Metasynchronous"]
                else "Adenocarcinoma (Synchronous)"
                if subtype
                in [
                    "Adc/ Synchronous",
                    "Synchronous Adenocarcinomas",
                    "Synchronous Adenocarcinomas Index Tumour",
                ]
                else "Adenocarcinoma With Mucinous Features"
                if subtype in ["Adenocarcinoma With Mucinous Features"]
                else "Ductal Adenocarcinoma"
                if subtype in ["Ductal Adenocarcinoma"]
                else "Adenocarcinoma With Medullary Features"
                if subtype
                in [
                    "Matched Met / Adenocarcinoma With Medullary Features",
                    "Matched Met / Carcinoma With Medullary Features",
                    "Primary / Adenocarcinoma With Medullary Features",
                    "Primary / Carcinoma With Medullary Features",
                ]
                else "Medullary Adenocarcinoma"
                if subtype
                in [
                    "Matched Met / Medullary Carcinoma",
                    "Primary / Medullary Carcinoma",
                ]
                else "Adenocarcinoma (Polyp)"
                if subtype in ["Polyp Adenocarcinoma"]
                else "Serrated Adenocarcinoma"
                if subtype in ["Serrated Adenocarcinoma"]
                else "Mucinous And Signet Ring Cell Adenocarcinoma"
                if subtype
                in ["Transverse Colon_Mucinous/Signet Ring Cell Adenocarcinoma"]
                else "Moderately Differentiated Adenocarcinoma"
                if subtype in ["Moderately Differentiated Adenocarcinoma Duodenum"]
                else "Poorly Differentiated Adenocarcinoma"
                if subtype in ["Poorly Differentiated Adenocarcinoma Of Duodenum"]
                else "Poorly Differentiated Focally Mucinous Adenocarcinoma"
                if subtype
                in [
                    "Poorly Differentiated Focally Mucinous Adenocarcinoma Of The Duodenum"
                ]
                else "Centrocytic Lymphoma (B-Cell)"
                if subtype in ["Centrocytic Lymphoma, B Cell"]
                else "Mantle Cell Lymphoma"
                if subtype in ["Mantle Cell"]
                else "No Cancer"
                if diagnosis
                in [
                    "Normal Colon",
                    "Colon Nor.",
                    "Colon  Nor.",
                    "Ileum Nor.",
                    "Normal Adjancent",
                ]
                or subtype
                in [
                    "Normal",
                    "Matched Normal",
                    "Normale",
                    "Normal Adjacent",
                    "Adjacent Normal",
                    "Normal Adjacent Tissues",
                    "No Tumor",
                ]
                else "Mixed Glandular-Neuroendocrine (Composite) Carcinoma"
                if subtype in ["Mixed Glandular-Neuroendocrine (Composite) Carcinoma"]
                else "Neuroendocrine Tumour Or A Mixed Adenocarcinoma And Neurotendocrine Tumour"
                if subtype
                in [
                    "Neuroendocrine Tumour Or A Mixed Adenocarcinoma And Neurotendocrine Tumour"
                ]
                else "KRAS G12C Only"
                if subtype in ["G12C Only"]
                else "Adenocarcinoma (KRAS-Mutated)"
                if subtype == "Adc + Kras-Mutated"
                else "Diffuse Large B-Cell Lymphopma"
                if subtype in ["Dlbcl", "Dlbcl Various Site"]
                else "Moderately Differentiated Adenocarcinoma"
                if diagnosis in ["Moderately Differentiated Adenocarcinoma Duodenum"]
                else "Poorly Differentiated Adenocarcinoma"
                if diagnosis in ["Poorly Differentiated Adenocarcinoma Of Duodenum"]
                else "Poorly Differentiated Focally Mucinous Adenocarcinoma"
                if diagnosis
                in [
                    "Poorly Differentiated Focally Mucinous Adenocarcinoma Of The Duodenum"
                ]
                else "NR"
            ),
            molecular_information=(
                "KRAS-Mutated"
                if subtype == "Adc + Kras-Mutated"
                else "KRAS G12C Only"
                if subtype == "G12C Only"
                else "NR"
            ),
            pairing_status=(
                "P"
                if subtype
                in [
                    "Matched Met / Adenocarcinoma",
                    "Matched Normal",
                    "Matched Met / Carcinoma With Medullary Features",
                    "Matched Met / Medullary Carcinoma",
                    "Matched Met",
                    "Matched Met / Adenocarcinoma With Medullary Features",
                    "Crc_Matched Met",
                ]
                or diagnosis in ["Colon Pri. And Met.", "Colon Pri. And Met. (Liver)"]
                else "No Cancer"
                if diagnosis
                in [
                    "Normal Colon",
                    "Colon Nor.",
                    "Colon  Nor.",
                    "Ileum Nor.",
                    "Normal Adjancent",
                ]
                or subtype
                in [
                    "Normal",
                    "Matched Normal",
                    "Normale",
                    "Normal Adjacent",
                    "Adjacent Normal",
                    "Normal Adjacent Tissues",
                    "No Tumor",
                ]
                else "NR"
            ),
            pairing_site=(
                "Colon/Rectum"
                if subtype
                in [
                    "Matched Met / Adenocarcinoma",
                    "Matched Normal",
                    "Matched Met / Carcinoma With Medullary Features",
                    "Matched Met / Medullary Carcinoma",
                    "Matched Met",
                    "Matched Met / Adenocarcinoma With Medullary Features",
                    "Crc_Matched Met",
                    "Liver Meta",
                    "Ln Meta",
                    "Matched Met",
                    "Matched Met / Adenocarcinoma",
                    "Matched Met / Adenocarcinoma With Medullary Features",
                    "Matched Met / Carcinoma With Medullary Features",
                    "Matched Met / Medullary Carcinoma",
                ]
                else "Liver"
                if diagnosis
                in ["Colon Pri. And Met. (Liver)", "Metastatic Renal Cell Carcinoma"]
                else "Colon"
                if diagnosis in ["Colon Adenocarcinoma Lung Metastase", "Colon Met."]
                else "No Cancer"
                if diagnosis
                in [
                    "Normal Colon",
                    "Colon Nor.",
                    "Colon  Nor.",
                    "Ileum Nor.",
                    "Normal Adjancent",
                ]
                or subtype
                in [
                    "Normal",
                    "Matched Normal",
                    "Normale",
                    "Normal Adjacent",
                    "Adjacent Normal",
                    "Normal Adjacent Tissues",
                    "No Tumor",
                ]
                else "NA"
            ),
            i=i,
        )


# Export to Box & Load it to format the column width
intestines, intestines_inventory = format_and_export(
    file_name="intestines", tissue_source="Intestines"
)

# In[25]: Liver (Verified)

# Create the combinations dictionary for Liver
create_combinations_dict(tissue_source="Liver")

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Liver":
        assign_values(
            anatomical_location="Liver",
            anatomical_sublocation="NR",
            specimen_type=(
                "M"
                if diagnosis
                in [
                    "Colon Met.",
                    "Liver Metastasis",
                    "Colon Pri. And Met. (Liver)",
                    "Not Hepatocellular Carcinoma",
                    "Matched Met, Liver",
                    "Liver Meta",
                ]
                or subtype in ["Metastasis", "Metastases"]
                else "N"
                if any(keyword in subtype for keyword in ["Normal"])
                else "P"
            ),
            diagnosis=(
                "Hepatocellular Carcinoma"
                if diagnosis
                in [
                    "Hcc",
                    "Hepatocellular Carcinoma",
                    "Recurrent Hepatocellular Carcinoma",
                ]
                or subtype in ["Hepatocarcinoma", "Hcc"]
                else "Hepatocellular Neoplasm"
                if diagnosis in ["Hepatocellular Neoplasm"]
                else "Hepatocellular Adenoma"
                if diagnosis in ["Hepatocellular Adenoma"]
                else "Hepatocellular Adenomas"
                if diagnosis in ["Hepatocellular Adenomas"]
                else "Mixed Carcinoma"
                if diagnosis in ["Hepatocellular Carcinoma And Cholangiocarcinoma"]
                else "Colorectal Cancer"
                if diagnosis in ["Crc"]
                else "No Cancer"
                if any(keyword in subtype for keyword in ["Normal"])
                else "Chronic Active Hepatitis"
                if subtype == "Chronic Active Hepatitis"
                else "Nonalcoholic Steatohepatitis"
                if subtype in ["Nash", "Nash Liver"]
                else "Primary Biliary Cirrhosis"
                if subtype == "Primary Biliary Cirrhosis"
                else "Fatty Change"
                if subtype == "Fatty Change"
                else "Portal Inflammation"
                if subtype == "Portal Inflammation"
                else "Cirrhosis"
                if subtype == "Cirrhosis"
                else "Intrahepatic Cholangiocarcinoma"
                if subtype
                in [
                    "Ihcc (Intrahepatic Cholangiocracinoma)",
                    "Mass-Forming Intrahepatic Cholangiocarcinoma",
                    "Intrahepatic Cholangiocarcinoma (Showing Heterogeneous Glandular And Nested Architecture)",
                    "Intrahep. Cholangiocarcinoma (Intracytoplasmic Mucin Secretion)",
                    "Intrahep. Cholangiocarcinoma",
                    "Intrahep. Cholangiocarcinoma (Primary)",
                    "Intrahep. Cholangiocarcinoma (Mass-Forming)",
                    "Intrahep. Cholangiocarcinoma (Glandular And Nested Architecture)",
                    "Intrahep. Cholangiocarcinoma (Intraductal)",
                    "Intrahep. Cholangiocarcinoma (Cystadenocarcinoma)",
                ]
                else "Intrahepatic Adenocarcinoma"
                if subtype in ["Intrahep. Adenocarcinoma (Pancreatobiliary)"]
                else "Renal Cell Carcinoma"
                if "Renal Cell Carcinoma" in subtype
                else "Primary Biliary Cholangitis"
                if diagnosis == "Pbc Liver"
                else "Adenocarcinoma"
                if subtype in ["Adc/ Synchronous", "Adc/ Metasynchronous"]
                else "NR"
            ),
            subtype=(
                "Hepatocellular Carcinoma And Cholangiocarcinoma"
                if diagnosis in ["Hepatocellular Carcinoma And Cholangiocarcinoma"]
                else "Mass-Forming Intrahepatic Cholangiocarcinoma"
                if subtype
                in [
                    "Mass-Forming Intrahepatic Cholangiocarcinoma",
                    "Intrahep. Cholangiocarcinoma (Mass-Forming)",
                ]
                else "Intrahepatic Cholangiocarcinoma With Heterogeneous Glandular And Nested Architecture"
                if subtype
                in [
                    "Intrahepatic Cholangiocarcinoma (Showing Heterogeneous Glandular And Nested Architecture)",
                    "Intrahep. Cholangiocarcinoma (Glandular And Nested Architecture)",
                ]
                else "Intrahepatic Cholangiocarcinoma With Intracytoplasmic Mucin Secretion"
                if subtype
                == "Intrahep. Cholangiocarcinoma (Intracytoplasmic Mucin Secretion)"
                else "Intrahepatic Adenocarcinoma (Pancreatobiliary)"
                if subtype in ["Intrahep. Adenocarcinoma (Pancreatobiliary)"]
                else "Intrahepatic Cholangiocarcinoma (Intraductal)"
                if subtype in ["Intrahep. Cholangiocarcinoma (Intraductal)"]
                else "Intrahepatic Cholangiocarcinoma (Cystadenocarcinoma)"
                if subtype in ["Intrahep. Cholangiocarcinoma (Cystadenocarcinoma)"]
                else "Metasynchronous"
                if "Metasynchronous" in subtype
                else "Synchronous"
                if "Synchronous" in subtype
                else "No Cancer"
                if any(keyword in subtype for keyword in ["Normal"])
                else "Renal Cell Carcinoma With Clear Cell Areas"
                if "Renal Cell Carcinoma" in subtype
                else "Recurrent Hepatocellular Carcinoma"
                if diagnosis == "Recurrent Hepatocellular Carcinoma"
                else "NR"
            ),
            pairing_status=(
                "P"
                if diagnosis
                in [
                    "Matched Met, Liver",
                    "Colon Met.",
                    "Crc",
                    "Colon Pri. And Met. (Liver)",
                ]
                else "No Cancer"
                if any(keyword in subtype for keyword in ["Normal"])
                else "NR"
            ),
            pairing_site=(
                "Colon/Rectum"
                if diagnosis in ["Colon Met.", "Crc", "Colon Pri. And Met. (Liver)"]
                else "No Cancer"
                if any(keyword in subtype for keyword in ["Normal"])
                else "Kidney"
                if "Renal Cell Carcinoma" in subtype
                else "NA"
            ),
            i=i,
        )


# Export to Box & Load it to format the column width
liver, liver_inventory = format_and_export(file_name="liver", tissue_source="Liver")


# In[26]: Heart (Verified)

# Create the combinations dictionary for Heart
create_combinations_dict("Heart")

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Heart":
        assign_values(
            anatomical_location="Heart",
            anatomical_sublocation="NR",
            specimen_type="N",
            diagnosis="No Cancer",
            subtype="No Cancer",
            pairing_site="No Cancer",
            pairing_status="No Cancer",
            molecular_information="NR",
            i=i,
        )


# Export to Box & Load it to format the column width
heart, heart_inventory = format_and_export(file_name="heart", tissue_source="Heart")

# In[27]: Lymphatic System (Verified)
for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Lymphatic System":
        current_fullDiagnosis = (diagnosis, subtype)
        assign_values(
            anatomical_location=(
                "Lymph Nodes"
                if diagnosis
                in [
                    "Ln",
                    "Axillary Nodes",
                    "Cervical Lymph Node",
                    "Cervical Lymph Nodes",
                    "Inguinal Lymph Node",
                    "Left Cervical Lymph Node",
                    "Left Level 1 Node",
                    "Left Level 2A Node",
                    "Left Level V Node",
                    "Left Neck Level Ii Lymph Node",
                    "Left Neck Lymph Node",
                    "Level 1B Node",
                    "Limph Node",
                    "Limph Node Nor.",
                    "Ln, Axilla",
                    "Ln, Inguinal",
                    "Ln, Neck",
                    "Ln, Retroperioneal",
                    "Ln, Subclavian",
                    "Lymph Node",
                    "Lymph Node Met. (Breast)",
                    "Lymph Node Metastasis",
                    "Lymph Node Nor.",
                    "Lymph Node Pri.",
                    "Matcehd Met, Lymph Node",
                    "Neck (Nodes)",
                    "Neck, Right Level 2 Lymph Nodes",
                    "Nodal Met",
                    "Nodal Met.",
                    "Nodes (Groin)",
                    "Right Inguinal Lymph Node",
                    "Right Level 2A Node",
                    "Right Level V Node",
                    "Right Neck Level Ii Lymph Node",
                    "Right Neck Lymph Node",
                    "Right Neck Node",
                    "Right Submandibular Lymph Node",
                    "Short Level 2A Long Level 4",
                    "Supraclavicular Lymph Node",
                    "Skin + Nodes",
                    "Wedge Of Upper Lobe & Node",
                    "Crc Nodal Met",
                    "Groin Node",
                    "Groin Nodes",
                    "Groin, Lymph Nodes",
                    "Inguinal Nodes",
                    "L Ileo Inguinal Nodes",
                    "Left Groin, Nodes",
                    "Left Superficial Groin, Left Iliac Nodes",
                    "Lymph Node Met. (Melanoma)",
                    "Lymph Nodes",
                    "Lymph Nodes Pri.",
                    "Neck, Left Level I Node",
                    "Nodes",
                    "Nodes (Left Groin, Right External Iliac)",
                    "Nsclc",
                    "Obturator Nodes",
                    "Right Common Iliac Node",
                    "Right Facial Lymph Node",
                    "Right Flank,  Ilio-Inguinal Lymph Node, Skin",
                    "Right Inguinal Nodes",
                    "Right Obturator Node",
                    "Right Obturator Nodes, Right Superficial Groin Scar",
                    "Right Superficial Groin Node",
                    "Sentinel Node Right Groin",
                    "Stomach",
                    "Superficial Inguinal Lymph Node",
                    "Superficial Inguinal Lymph Nodes",
                    "Upper Neck Lymph Nodes",
                    "Right Superficial Groin, Node",
                    "Lung Nodal Met.",
                    "Hen Nodal Met.",
                ]
                else "Bone Marrow"
                if diagnosis in ["Bone", "Bone Marrow"]
                else "Serosa Nodules"
                if diagnosis in ["Representive Examples Of Serosa Nodules"]
                else "Lymph Nodes/Parotid Gland"
                if diagnosis
                in [
                    "Neck, Level 2 Node, Superficial Right Parotid",
                    "Neck, Lymph Node & Partoid Gland",
                ]
                else "Caecum"
                if diagnosis in ["Cecum", "Caecum", "Right Caecum"]
                else "Colon"
                if diagnosis in ["Colon"]
                else "Colon/Ileum/Omentum"
                if diagnosis in ["Colon Ileum Omentum"]
                else "Leg"
                if diagnosis in ["Cute (Leg)"]
                else "Gum"
                if diagnosis in ["Gum"]
                else "Gum/Parotid Gland"
                if diagnosis in ["Gum/Parotid Gland"]
                else "Neck"
                if diagnosis in ["Left Neck Mass", "Neck"]
                else "Tonsil"
                if diagnosis
                in [
                    "Left Tonsil Mass",
                    "Right Tonsil",
                    "Tonsils",
                    "Left Tonsillar Mass",
                ]
                else "Meninges"
                if diagnosis in ["Meninges"]
                else "Lymph Nodes"
                if diagnosis in ["Normal (Non-Malignant) Endometrium"]
                else "Spleen"
                if diagnosis in ["Normal Spleen", "Spleen", "Speen", "Spleen Nor."]
                else "Parotid Gland"
                if diagnosis in ["Parotid Gland"]
                else "Pharynx"
                if diagnosis in ["Pharynx Cancer With Ln Neck Metastasis"]
                else "Skin"
                if diagnosis in ["Skin"]
                else "Small Intestine/Colon"
                if diagnosis in ["Small Bowel And Colon"]
                else "Small Intestine"
                if diagnosis in ["Small Bowel"]
                else "Sternum"
                if diagnosis in ["Sternum"]
                else "Subcutaneous Tissue"
                if diagnosis in ["Subcutis"]
                else "Testis"
                if diagnosis in ["Testis", "Testis Pri."]
                else "Vulva"
                if diagnosis in ["Vulva"]
                else "Lung"
                if diagnosis in ["Wedge Of Upper Lobe & Node"]
                else "NR"
            ),
            anatomical_sublocation=(
                "Lymph Nodes (Axillary)"
                if diagnosis in ["Axillary Nodes", "Ln, Axilla"]
                else "Lymph Nodes (Cervical)"
                if diagnosis in ["Cervical Lymph Node", "Cervical Lymph Nodes"]
                else "Lymph Nodes (Inguinal)"
                if diagnosis
                in ["Inguinal Lymph Node", "Ln, Inguinal", "Inguinal Nodes"]
                else "Lymph Nodes (Left Cervical)"
                if diagnosis in ["Left Cervical Lymph Node"]
                else "Lymph Nodes (Left Level 1)"
                if diagnosis in ["Left Level 1 Node", "Neck, Left Level I Node"]
                else "Lymph Nodes (Left Level 2A)"
                if diagnosis in ["Left Level 2A Node"]
                else "Lymph Nodes (Left Level 2) And Parotid Gland (Superficial Right)"
                if diagnosis in ["Neck, Level 2 Node, Superficial Right Parotid"]
                else "Lymph Nodes (Left Level 5)"
                if diagnosis in ["Left Level V Node"]
                else "Lymph Nodes (Left Level 2)"
                if diagnosis in ["Left Neck Level Ii Lymph Node"]
                else "Lymph Nodes (Left Neck)"
                if diagnosis in ["Left Neck Lymph Node"]
                else "Lymph Nodes (Left Level 1B)"
                if diagnosis in ["Level 1B Node"]
                else "Lymph Nodes (Retroperioneal)"
                if diagnosis in ["Ln, Retroperioneal"]
                else "Lymph Nodes (Subclavian)"
                if diagnosis in ["Ln, Subclavian"]
                else "Lymph Nodes (Neck)"
                if diagnosis in ["Ln, Neck"]
                else "Lymph Nodes (Neck) And Parotid Gland"
                if diagnosis in ["Neck, Lymph Node & Partoid Gland"]
                else "Lymph Nodes (Right Level 2)"
                if diagnosis
                in ["Neck, Right Level 2 Lymph Nodes", "Right Neck Level Ii Lymph Node"]
                else "Lymph Nodes (Right Level 2A)"
                if diagnosis in ["Right Level 2A Node"]
                else "Lymph Nodes (Left Superficial Groin And Left Iliac)"
                if diagnosis in ["Left Superficial Groin, Left Iliac Nodes"]
                else "Lymph Nodes (Groin)"
                if diagnosis
                in ["Nodes (Groin)", "Groin Node", "Groin Nodes", "Groin, Lymph Nodes"]
                else "Lymph Nodes (Left Groin)"
                if diagnosis in ["Left Groin, Nodes"]
                else "Lymph Nodes (Right Inguinal)"
                if diagnosis in ["Right Inguinal Lymph Node", "Right Inguinal Nodes"]
                else "Lymph Nodes (Left Ileo Inguinal)"
                if diagnosis in ["L Ileo Inguinal Nodes"]
                else "Lymph Nodes (Right Level 5)"
                if diagnosis in ["Right Level V Node"]
                else "Lymph Nodes (Right Neck)"
                if diagnosis in ["Right Neck Lymph Node", "Right Neck Node"]
                else "Lymph Nodes (Right Submandibular)"
                if diagnosis in ["Right Submandibular Lymph Node"]
                else "Lymph Nodes (Short Level 2A / Long Level 4)"
                if diagnosis in ["Short Level 2A Long Level 4"]
                else "Lymph Nodes (Supraclavicular)"
                if diagnosis in ["Supraclavicular Lymph Node"]
                else "Lymph Nodes (Pelvic)"
                if subtype in ["Pelvic Node"]
                else "Lymph Nodes (Para-Aortic)"
                if subtype in ["Para-Aortic Node"]
                else "Neck (Left Mass)"
                if diagnosis in ["Left Neck Mass"]
                else "Tonsil (Left Mass)"
                if diagnosis in ["Left Tonsil Mass", "Left Tonsillar Mass"]
                else "Tonsil (Right)"
                if diagnosis in ["Right Tonsil"]
                else "Lung (Upper Lobe And Node)"
                if diagnosis in ["Wedge Of Upper Lobe & Node"]
                else "Lymph Nodes (Neck)"
                if diagnosis in ["Neck (Nodes)"]
                else "Lymph Nodes (Left Groin And Right External Iliac)"
                if diagnosis in ["Nodes (Left Groin, Right External Iliac)"]
                else "Lymph Nodes (Obturator)"
                if diagnosis in ["Obturator Nodes"]
                else "Lymph Nodes (Right Common Iliac)"
                if diagnosis in ["Right Common Iliac Node"]
                else "Lymph Nodes (Right Facial)"
                if diagnosis in ["Right Facial Lymph Node"]
                else "Lymph Nodes (Right Ilio-Inguinal)"
                if diagnosis in ["Right Flank,  Ilio-Inguinal Lymph Node, Skin"]
                else "Lymph Nodes (Right Obturator)"
                if diagnosis in ["Right Obturator Node"]
                else "Lymph Nodes (Right Obturator And Right Superficial Groin Scar)"
                if diagnosis in ["Right Obturator Nodes, Right Superficial Groin Scar"]
                else "Lymph Nodes (Right Superficial Groin)"
                if diagnosis
                in ["Right Superficial Groin Node", "Right Superficial Groin, Node"]
                else "Lymph Nodes (Sentinel Right Groin)"
                if diagnosis in ["Sentinel Node Right Groin"]
                else "Lymph Nodes (Superficial Inguinal)"
                if diagnosis
                in [
                    "Superficial Inguinal Lymph Node",
                    "Superficial Inguinal Lymph Nodes",
                ]
                else "Lymph Nodes (Upper Neck)"
                if diagnosis in ["Upper Neck Lymph Nodes"]
                else "NR"
            ),
            specimen_type=(
                "N"
                if any(keyword in diagnosis for keyword in ["Nor.", "Normal"])
                or any(keyword in subtype for keyword in ["Nor.", "Normal"])
                else "M"
                if diagnosis
                in [
                    "Lymph Node Met. (Breast)",
                    "Lymph Node Metastasis",
                    "Matcehd Met, Lymph Node",
                    "Nodal Met",
                    "Nodal Met.",
                    "Wedge Of Upper Lobe & Node",
                    "Mucosal Melanoma",
                    "Crc Nodal Met",
                    "Lymph Node Met. (Melanoma)",
                    "Lung Nodal Met.",
                    "Matched Met, Left Axilla Node",
                ]
                or subtype
                in [
                    "High Grade Serous",
                    "Metastasis",
                    "Mets",
                    "Oral Cavity_Matched Met",
                    "Scc",
                    "Matched Met (Lh18-37000)",
                    "Small Cell Carcinoma",
                    "Clear Renal Cell Carcinoma",
                    "Papillary Renal Cell Carcinoma",
                    "Squamous Cell Carcinoma (Pharyngeal)",
                    "Clear Cell Renal Cell Carcinoma",
                    "Combined Small Cell And Large Neuroendocrine Carcinoma",
                ]
                or any(keyword in subtype for keyword in ["Metastatic", "Matched Met"])
                else "NR"
                if current_fullDiagnosis
                in {
                    ("Ln", "/"),
                    ("Neck (Nodes)", "-"),
                    ("Skin + Nodes", "-"),
                    ("Nodes (Groin)", "-"),
                    ("Spleen", "-"),
                    ("Lymph Node"),
                    ("Lymph Node", "-"),
                }
                else "P"
            ),
            diagnosis=(
                "Adenocarcinoma"
                if subtype
                in [
                    "Acinar\nAdenocarcinoma",
                    "Acinar Adenocarcinoma",
                    "Acinar Adenocarcinoma With Colloid Component (20%)",
                    "Adc",
                    "Intestinal Adenocarcinoma",
                    "Invasive Acinar Adenocarcinoma",
                    "Invasive Mucinous Adenocarcinoma",
                    "Solid Adenocarcinoma",
                ]
                else "Mixed Adenocarcinoma"
                if subtype
                in [
                    "Acinar (70%) Solid (30%) Adenocarcinoma",
                    "Solid (50%) Acinar (40%) And Micropapillary (10%) Adenocarcinoma",
                    "Solid (90%) Adenocarcinoma With Acinar Component (10%)",
                    "Tubulopapillary And Micropapillary Adenocarcinoma (Lung Primary)",
                ]
                else "Peripheral T-Cell Lymphoma"
                if subtype
                in [
                    "Angioimmunoblastic T Cell",
                    "Angioimmunoblastic T-Cell",
                    "Angioimmunoblastic T-Cell L",
                    "Enteropathy Associated T-Cell",
                    "Enteropathy-Associated T Cell",
                    "Extranodal Nk/T Cell",
                    "Extranodal Nk/T-Cell L",
                    "Peripheral T - Cell",
                    "Ptcl, Nos",
                    "Ptcl, Nos Progressed From Hodgkin Lymphoma, Ns",
                ]
                else "Peripheral B-Cell Lymphoma"
                if subtype
                in [
                    "B-Sll",
                    "Burkitt Type B-Cell",
                    "Burkitt Lymphoma",
                    "Centrocytic Lymphoma, B Cell",
                    "Centrocytic Lymphoma, Recurrent",
                    "Chronic Lymphocytic Leukaemia/Small Lymphocytic Lymphoma",
                    "Cll",
                    "Cll/Sll",
                    "Dlbcl",
                    "Diffuse Large B-Cell",
                    "Dlbcl Various Site",
                    "Fl",
                    "Follicular Lymphoma",
                    "Mantle Cell Lymphoma, Not Stated",
                    "Mantle Cell Lymphoma, Peripheral B-Cell Lymphoma",
                    "Mantle Cell Lymphoma, Recurrent",
                    "Mantle Cell Lymphoma, Staging Marrow",
                    "Mcl",
                    "Sll",
                    "Sll/B-Cll",
                    "Sll/Cll",
                ]
                or diagnosis in ["Fl", "Dlbcl", "Sll"]
                else "Peripheral B-Cell And T-Cell Lymphoma"
                if subtype in ["Ptcl, Nos Mixed With Dlbcl"]
                else "Squamous Cell Carcinoma"
                if subtype
                in [
                    "Basaloid Squamous Cell Carcinoma",
                    "Keratinizing Squamous Cell Carinoma",
                    "Non-Keratinizing Squamous Cell Carinoma",
                    "Scc",
                    "Squamous Cell Carcinoma",
                    "Squamous Cell Carcinoma (Pharyngeal)",
                    "Squamous Cell Carcinoma, H&N",
                ]
                else "Renal Cell Carcinoma"
                if subtype
                in [
                    "Clear Cell Renal Cell Carcinoma",
                    "Papillary Renal Cell Carcinoma",
                    "Rcc Conventiol",
                    "Rcc Papillary Type Ii",
                ]
                else "Combined Carcinoma"
                if subtype in ["Combined Small Cell And Large Neuroendocrine Carcinoma"]
                else "Marginal Zone Lymphoma"
                if subtype in ["Mzl", "Extranodal Marginal Zone", "Marginal Zone"]
                or diagnosis in ["Mzl"]
                else "Hogdkin Lymphoma"
                if subtype
                in [
                    "Hd",
                    "Hd (Lymphocytic Rich)",
                    "Hd (Mixed Cellularity)",
                    "Hd (Nodular Sclerosis)",
                ]
                else "Carcinoma"
                if subtype in ["High Grade Serous"]
                else "Invasive Carcinoma"
                if subtype in ["Idc"]
                else "Melanoma"
                if diagnosis in ["Mucosal Melanoma"]
                or subtype
                in [
                    "Metastatic Superficial Spreading Melanoma",
                    "Metastatic Melanoma, Nos",
                    "Metastatic Acral Lentiginous Melanoma",
                    "Recurrent Melanoma",
                    "Metastatic Melanoma",
                    "Acral Lentiginous Malignant Melanoma",
                    "Superficial Spreading Malignant Melanoma",
                    "Malignant Melanoma",
                    "Metastatic Malignant Melanoma",
                    "Metastatic Melanoma, Nodular",
                    "Malignant Melanoma",
                    "Melanoma",
                ]
                else "No Cancer"
                if diagnosis in ["Normal (Non-Malignant) Endometrium", "Normal Spleen"]
                or subtype in ["Normal"]
                else "Non-Hodgkin Lymphoma"
                if subtype in ["Nhl"]
                else "Non-Small Cell Lung Cancer"
                if subtype in ["Nsclc", "Nsclc (Adc)", "Nsclc (Scc)"]
                else "Neuroendocrine Carcinoma"
                if subtype in ["Small Cell Carcinoma"]
                else "Stomach Lymphoma"
                if subtype in ["Stomach Lymphoma With Infiltration Into The Colon"]
                else "T-Cell Lymphoma"
                if subtype in ["T-Cell", "Tcl"]
                else "Cancer"
                if subtype in ["Cancer"]
                else "Non-Small Cell Lung Cancer"
                if diagnosis in ["Nsclc"]
                else "Colorectal Cancer"
                if diagnosis in ["Crc Nodal Met", "Crc_Matched Met"]
                or subtype in ["Crc_Matched Met"]
                else "NR"
            ),
            subtype=(
                "Acinar Adenocarcinoma"
                if subtype
                in [
                    "Acinar\nAdenocarcinoma",
                    "Acinar Adenocarcinoma",
                    "AcinarAdenocarcinoma",
                ]
                else "Acinar Adenocarcinoma With Colloid Component (20%)"
                if subtype in ["Acinar Adenocarcinoma With Colloid Component (20%)"]
                else "Angioimmunoblastic T-Cell Lymphoma"
                if subtype
                in [
                    "Angioimmunoblastic T Cell",
                    "Angioimmunoblastic T-Cell",
                    "Angioimmunoblastic T-Cell L",
                ]
                else "Small Lymphocytic Lymphoma (B-Cell)"
                if subtype in ["B-Sll", "Sll"] or diagnosis in ["Sll"]
                else "Basaloid Squamous Cell Carcinoma"
                if subtype in ["Basaloid Squamous Cell Carcinoma"]
                else "Burkitt Lymphoma"
                if subtype in ["Burkitt Lymphoma"]
                else "Burkitt-Type Lymphoma (B-Cell)"
                if subtype in ["Burkitt Type B-Cell"]
                else "Centrocytic Lymphoma (B-Cell)"
                if subtype in ["Centrocytic Lymphoma, B Cell"]
                else "Centrocytic Lymphoma (Recurrent)"
                if subtype in ["Centrocytic Lymphoma, Recurrent"]
                else "Chronic Lymphocytic Leukaemia And Small Lymphocytic Lymphoma"
                if subtype
                in [
                    "Chronic Lymphocytic Leukaemia/Small Lymphocytic Lymphoma",
                    "Sll/Cll",
                    "Cll/Sll",
                    "Sll/B-Cll",
                ]
                else "Clear Cell Carcinoma"
                if subtype in ["Clear Cell Renal Cell Carcinoma"]
                else "Chronic Lymphocytic Leukaemia"
                if subtype in ["Cll"]
                else "Diffuse Large B-Cell Lymphoma"
                if subtype in ["Dlbcl", "Diffuse Large B-Cell"]
                or diagnosis in ["Dlbcl"]
                else "Diffuse Large B-Cell Lymphoma (Various Sites)"
                if subtype in ["Dlbcl Various Site"]
                else "Enteropathy Associated T-Cell"
                if subtype
                in ["Enteropathy Associated T-Cell", "Enteropathy-Associated T Cell"]
                else "Extranodal NK/T-Cell Lymphoma"
                if subtype in ["Extranodal Nk/T Cell", "Extranodal Nk/T-Cell L"]
                else "Extranodal Marginal Zone Lymphoma"
                if subtype in ["Extranodal Marginal Zone"]
                else "Follicular Lymphoma"
                if subtype in ["Fl", "Follicular Lymphoma"] or diagnosis in ["Fl"]
                else "Hodgkin Lymphoma (Lymphocytic Rich)"
                if subtype in ["Hd (Lymphocytic Rich)"]
                else "Hodgkin Lymphoma (Mixed Cellularity)"
                if subtype in ["Hd (Mixed Cellularity)"]
                else "Hodgkin Lymphoma (Nodular Sclerosis)"
                if subtype in ["Hd (Nodular Sclerosis)"]
                else "Invasive Ductal Carcinoma"
                if subtype in ["Idc"]
                else "Intestinal Adenocarcinoma"
                if subtype in ["Intestinal Adenocarcinoma"]
                else "Invasive Acinar Adenocarcinoma"
                if subtype in ["Invasive Acinar Adenocarcinoma"]
                else "Invasive Mucinous Adenocarcinoma"
                if subtype in ["Invasive Mucinous Adenocarcinoma"]
                else "Keratinizing Squamous Cell Carinoma"
                if subtype in ["Keratinizing Squamous Cell Carinoma"]
                else "Mantle Cell Lymphoma (NOS)"
                if subtype
                in [
                    "Mantle Cell Lymphoma, Not Stated",
                    "Mantle Cell Lymphoma, Peripheral B-Cell Lymphoma",
                ]
                else "Mantle Cell Lymphoma (Recurrent)"
                if subtype in ["Mantle Cell Lymphoma, Recurrent"]
                else "Mantle Cell Lymphoma (Staging Marrow)"
                if subtype in ["Mantle Cell Lymphoma, Staging Marrow"]
                else "Mantle Cell Lymphoma"
                if subtype in ["Mcl"]
                else "No Cancer"
                if diagnosis in ["Normal (Non-Malignant) Endometrium", "Normal Spleen"]
                or subtype in ["Normal"]
                else "Non-Keratinizing Squamous Cell Carinoma"
                if subtype in ["Non-Keratinizing Squamous Cell Carinoma"]
                else "Non-Small Cell Lung Cancer (Adenocarcinoma)"
                if subtype in ["Nsclc (Adc)"]
                else "Non-Small Cell Lung Cancer (Squamous Cell Carcinoma)"
                if subtype in ["Nsclc (Scc)"]
                else "Papillary Renal Cell Carcinoma"
                if subtype in ["Papillary Renal Cell Carcinoma"]
                else "Peripheral T-Cell Lymphoma (NOS)"
                if subtype in ["Peripheral T - Cell", "Ptcl, Nos"]
                else "Peripheral T-Cell Lymphoma (NOS) And Diffuse Large B-Cell Lymphoma"
                if subtype in ["Ptcl, Nos Mixed With Dlbcl"]
                else "Peripheral T-Celll Lymphoma (Progressed From Hodgkin Lymphoma)"
                if subtype in ["Ptcl, Nos Progressed From Hodgkin Lymphoma, Ns"]
                else "Papillary Renal Cell Carcinoma (Type II)"
                if subtype in ["Rcc Papillary Type Ii"]
                else "Solid (50%), Acinar (40%) And Micropapillary (10%) Adenocarcinoma"
                if subtype
                in ["Solid (50%) Acinar (40%) And Micropapillary (10%) Adenocarcinoma"]
                else "Solid (90%) Adenocarcinoma With Acinar Component (10%)"
                if subtype in ["Solid (90%) Adenocarcinoma With Acinar Component (10%)"]
                else "Stomach Lymphoma (NOS) Infiltrating Into The Colon"
                if subtype in ["Stomach Lymphoma With Infiltration Into The Colon"]
                else "Tubulopapillary And Micropapillary Adenocarcinoma"
                if subtype
                in ["Tubulopapillary And Micropapillary Adenocarcinoma (Lung Primary)"]
                else "Acinar (70%) And Solid (30%) Adenocarcinoma"
                if subtype in ["Acinar (70%) Solid (30%) Adenocarcinoma"]
                else "High Grade Serous Carcinoma"
                if subtype in ["High Grade Serous"]
                else "Solid Adenocarcinoma"
                if subtype in ["Solid Adenocarcinoma"]
                else "Small Cell Carcinoma"
                if subtype in ["Small Cell Carcinoma"]
                else "Superficial Spreading Melanoma"
                if subtype
                in [
                    "Metastatic Superficial Spreading Melanoma",
                    "Superficial Spreading Malignant Melanoma",
                ]
                else "Melanoma (NOS)"
                if subtype in ["Metastatic Melanoma, Nos"]
                else "Acral Lentiginous Melanoma"
                if subtype
                in [
                    "Metastatic Acral Lentiginous Melanoma",
                    "Acral Lentiginous Malignant Melanoma",
                ]
                else "Melanoma (Recurrent)"
                if subtype in ["Recurrent Melanoma"]
                else "Nodular Melanoma"
                if subtype in ["Metastatic Melanoma, Nodular"]
                else "Combined Small Cell And Large Neuroendocrine Carcinoma"
                if subtype in ["Combined Small Cell And Large Neuroendocrine Carcinoma"]
                else "NR"
            ),
            pairing_status=(
                "P"
                if (
                    any(keyword in diagnosis for keyword in ["Matcehd Met"])
                    or any(keyword in subtype for keyword in ["Matched Met"])
                )
                else "No Cancer"
                if diagnosis in ["Normal (Non-Malignant) Endometrium", "Normal Spleen"]
                or subtype in ["Normal"]
                else "NR"
            ),
            pairing_site=(
                "Lung"
                if subtype in ["Nsclc", "Nsclc (Scc)", "Nsclc (Adc)"]
                else "Lung (Upper Lobe)"
                if subtype
                in ["Tubulopapillary And Micropapillary Adenocarcinoma (Lung Primary)"]
                else "Oral Cavity"
                if subtype in ["Oral Cavity_Matched Met"]
                else "No Cancer"
                if diagnosis in ["Normal (Non-Malignant) Endometrium", "Normal Spleen"]
                or subtype in ["Normal"]
                else "Kidney"
                if subtype
                in [
                    "Clear Cell Renal Cell Carcinoma",
                    "Papillary Renal Cell Carcinoma",
                    "Rcc Papillary Type Ii",
                    "Rcc Conventiol",
                ]
                else "Pharynx"
                if diagnosis in ["Squamous Cell Carcinoma (Pharyngeal)"]
                else "Mucosa"
                if diagnosis in ["Mucosal Melanoma"]
                else "Hypopharynx"
                if subtype in ["Matched Met (Lh18-37000)"]
                else "Pharynx"
                if subtype in ["Squamous Cell Carcinoma (Pharyngeal)"]
                else "Breast"
                if diagnosis in ["Lymph Node Met. (Breast)"]
                else "Colon/Rectum"
                if subtype in ["Crc_Matched Met"]
                else "NA"
            ),
            molecular_information="NR",
            i=i,
        )


# Export to Box & Load it to format the column width
lymphatic_system, lymphatic_system_inventory = format_and_export(
    file_name="lymphatic_system", tissue_source="Lymphatic System"
)


# In[28]: GOJ (Verified)

# Create the combinations dictionary for GOJ
create_combinations_dict("GOJ")

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "GOJ":
        assign_values(
            anatomical_location=(
                "Gastroesophageal Junction"
                if diagnosis in ["Gastric Adc & Gej", "Goj"]
                else "Esophagus - Gastroesophageal Junction"
                if diagnosis in ["Esophagus/Gej", "Esophagus/Goj"]
                else "NR"
            ),
            anatomical_sublocation="NR",
            specimen_type=("M" if subtype in ["Matched Met"] else "P"),
            diagnosis=(
                "Gastric Adenocarcinoma"
                if subtype
                in [
                    "Gastric Adc Signet Ring Cell",
                    "Diffuse Gastric Adc",
                    "Gastric Adc",
                    "Diffuse-Type Adenocarcinoma",
                    "Adenocarcinoma Of The Gastro-Oesophageal Junction",
                    "Residual Junctional Adenocarcinoma",
                    "Adenocarcinoma",
                    "Adc",
                    "Goj Adc",
                ]
                else "Oesophageal Carcinoma"
                if subtype in ["Oesophageal Carcinoma"]
                else "Oesophageal Adenocarcinoma"
                if subtype in ["Oesophageal Adenocarcinoma"]
                else "Gastric Adenocarcinoma"
                if diagnosis == "Gastric Adc & Gej" and subtype == "-"
                else "NR"
            ),
            subtype=(
                "Signet Ring Cell Gastric Adenocarcinoma"
                if subtype in ["Gastric Adc Signet Ring Cell"]
                else "Diffuse Gastric Adenocarcinoma"
                if subtype in ["Diffuse Gastric Adc", "Diffuse-Type Adenocarcinoma"]
                else "Adenocarcinoma (Residual)"
                if subtype in ["Residual Junctional Adenocarcinoma"]
                else "NR"
            ),
            pairing_status=("P" if subtype in ["Matched Met"] else "NR"),
            pairing_site="NR",
            i=i,
        )


# Export to Box & Load it to format the column width
GOJ, GOJ_inventory = format_and_export(file_name="GOJ", tissue_source="GOJ")


# In[29]: Lower Body (Verified)

# Create the combinations dictionary for Lower Body
create_combinations_dict("Lower Body")

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Lower Body":
        assign_values(
            anatomical_location=(
                "Leg"
                if (
                    diagnosis
                    in [
                        "Leg",
                        "Left Leg",
                        "Right Lower Leg",
                        "Left Leg Pri. (Sarcoma)",
                        "Left Anterior Leg",
                        "Left Anterior Thigh",
                        "Lower Limb And Hip",
                        "Left Medial Thigh",
                        "Right Foot Pri. (Sarcoma)",
                        "Skin, Right Foot (Matched Met: Groin, Nodes)",
                    ]
                    or subtype in ["Lower Limb And Hip"]
                )
                else "Leg (Soft Tissue)"
                if subtype in ["Right Thigh Pri. (Sarcoma)"]
                else "Leg (Bone)"
                if diagnosis
                in [
                    "Femur",
                    "Right Distal Femur, Third",
                    "Left Distal Femur, Third",
                    "Left Femur",
                    "Left Fibula",
                ]
                else "Knee/Lymph Nodes"
                if diagnosis in ["Left Knee, Right Neck Node"]
                else "NR"
            ),
            anatomical_sublocation=(
                "Leg (Left)"
                if diagnosis in ["Left Leg", "Left Leg Pri. (Sarcoma)"]
                else "Thigh (Left Medial)"
                if diagnosis in ["Left Medial Thigh"]
                else "Thigh (Left Anterior)"
                if diagnosis in ["Left Anterior Thigh"]
                else "Leg (Right Lower)"
                if diagnosis in ["Right Lower Leg"]
                else "Thigh (Right - Soft Tissue)"
                if diagnosis in ["Right Thigh Pri. (Sarcoma)"]
                else "Femur (Right Distal Third)"
                if diagnosis in ["Right Distal Femur, Third"]
                else "Femur (Left Distal Third)"
                if diagnosis in ["Left Distal Femur, Third"]
                else "Fibula (Left)"
                if diagnosis in ["Left Fibula"]
                else "Femur (Left)"
                if diagnosis in ["Left Femur"]
                else "Leg (Left Anterior)"
                if diagnosis in ["Left Anterior Leg"]
                else "Foot (Right)"
                if diagnosis in ["Right Foot Pri."]
                else "Foot (Right - Skin)"
                if diagnosis in ["Skin, Right Foot (Matched Met: Groin, Nodes)"]
                else "Femur"
                if diagnosis in ["Femur"]
                else "Foot (Right)"
                if diagnosis in ["Right Foot Pri. (Sarcoma)"]
                else "Hip"
                if subtype in ["Lower Limb And Hip"]
                else "Knee (Left) And Neck (Right Lymph Nodes)"
                if diagnosis in ["Left Knee, Right Neck Node"]
                else "NR"
            ),
            specimen_type=(
                "M" if subtype in ["Metastatic Malignant Melanoma"] else "P"
            ),
            diagnosis=(
                "Squamous Cell Carcinoma"
                if subtype in ["Scc"] or diagnosis in ["Scc"]
                else "Basal Cell Carcinoma"
                if subtype in ["Bcc"]
                else "Chondrosarcoma"
                if subtype
                in ["High Grade Soft Tissue Chondrosarcoma", "Chondrosarcoma"]
                else "Osteosarcoma"
                if subtype
                in [
                    "High Gradeextraskeletal Osteosarcoma Of Firoblastic/Mfh-Like Subtype",
                    "Osteosarcoma",
                    "Conventional Osteosarcoma - Chondroblastic And Osteoblastic",
                    "Conventional Osteosarcoma",
                ]
                else "Sarcoma"
                if subtype in ["Undifferentiated High Grade Sarcoma"]
                else "T-Cell Lymphoma"
                if subtype in ["Tcl"]
                else "Melanoma"
                if subtype
                in ["Metastatic Malignant Melanoma", "Acral Lentiginous Melanoma"]
                else "NR"
            ),
            subtype=(
                "Chondrosarcoma (High Grade Soft Tissue)"
                if subtype in ["High Grade Soft Tissue Chondrosarcoma"]
                else "Sarcoma (Undifferentiated High Grade)"
                if subtype in ["Undifferentiated High Grade Sarcoma"]
                else "Chondroblastic And Osteoblastic Osteosarcoma"
                if subtype
                in ["Conventional Osteosarcoma - Chondroblastic And Osteoblastic"]
                else "Firoblastic/MFH-Like Osteosarcoma (High Grade Extraskeletal)"
                if subtype
                in [
                    "High Gradeextraskeletal Osteosarcoma Of Firoblastic/Mfh-Like Subtype"
                ]
                else "Acral Lentiginous Melanoma"
                if subtype in ["Acral Lentiginous Melanoma"]
                else "NR"
            ),
            pairing_status=("P" if "Matched Met" in diagnosis else "NR"),
            pairing_site=(
                "Lymph Nodes (Groin)"
                if "Matched Met: Groin, Nodes" in diagnosis
                else "NA"
            ),
            i=i,
        )


# Export to Box & Load it to format the column width
lower_body, lower_body_inventory = format_and_export(
    file_name="lower_body", tissue_source="Lower Body"
)


# In[30]: Kidney (Verified)
for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Kidney":
        assign_values(
            anatomical_location=(
                "Renal Bed"
                if diagnosis in ["Right Renal Bed"]
                else "Kidney, Renal Vein, And Thrombus"
                if diagnosis in ["Right Kidney, Left Renal Vein, Thrombus"]
                else "Kidney"
            ),
            anatomical_sublocation=(
                "Kidney (Left)"
                if diagnosis in ["Left Kidney"]
                else "Kidney (Right)"
                if diagnosis in ["Right Kidney", "Right Kidney (Primary)"]
                else "Kidney (Left And Right)"
                if diagnosis in ["Left And Normal Right Kidney"]
                else "Renal Bed (Right)"
                if diagnosis in ["Right Renal Bed"]
                else "Kidney (Right), Renal Vein (Left), And Thrombus"
                if diagnosis in ["Right Kidney, Left Renal Vein, Thrombus"]
                else "Kidney (Medulla)"
                if diagnosis in ["Kidney, Medulla"]
                else "Kidney (Left - Renal Parenchyma)"
                if "Renal Parenchyma" in subtype and diagnosis in ["Left Kidnety"]
                else "NR"
            ),
            specimen_type=(
                "N"
                if subtype in ["Normal", "Matched Normal"]
                else "N (Adjacent)"
                if subtype in ["Adjacent Normal"]
                else "NR"
                if diagnosis in ["Kidney"] and subtype in ["-"]
                else "P"
            ),
            diagnosis=(
                "Renal Cell Carcinoma"
                if diagnosis in ["Rcc"]
                or subtype
                in [
                    "2X Papillary Renal Cell Carcinomas",
                    "Ccc",
                    "Ccc Sarcomatoid",
                    "Chromophobe",
                    "Chromophobe Carcinoma",
                    "Chromophobe Carcinoma, Confined To Renal  Parenchyma",
                    "Chromophobe Cell Renal Cell Carcinoma",
                    "Chromophobe Renal Cell Carcinoma",
                    "Clear Cell Carcinoma",
                    "Clear Cell Renal Carcinoma",
                    "Clear Cell Renal Carcinoma, Focal Area Showing Rhabdoid Cells",
                    "Clear Cell Renal Cell Carcinoma",
                    "Clear Cell Renal Cell Carcinoma; Foci Of Rhabdoid Change",
                    "Clear Cell Type",
                    "Conventional (Clear Cell) Renal Carcinoma",
                    "Conventional (Clear Cell) Renal Cell Carcinoma",
                    "Conventional Clear Cell Renal Cell Carcinoma",
                    "Conventional Clear Cell Renal Cell Carcinoma With Focal Sarcomatoid Morphology And Coagulative Necrosis",
                    "Conventional Renal Cell Carcinoma",
                    "Conventional Renal Cell Carcinoma,",
                    "Multifocal Papillary Renal Cell Carcinoma",
                    "Papillary Carcinoma",
                    "Papillary Renal Carcinoma",
                    "Papillary Renal Cell Carcinoma",
                    "Papillary Type 2",
                    "Papillary, Type 1",
                    "Papillary, Type 2",
                    "Prcc",
                    "Rc Cromophobe",
                    "Rcc",
                    "Rcc Conventiol",
                    "Rcc Papillary Type Ii",
                    "Rccc",
                    "Renal Cell Carcinoma",
                    "Renal Cell Carcinoma (Type Uncertain)",
                    "Renal Cell Carcinoma (Unclassified; With Clear Cell Areas)",
                    "Renal Cell Carcinoma With Sarcomatoid Differentiation",
                    "Renal Cell Carcinoma, Chromophobe Cell Type",
                    "Renal Cell Carcinoma, Chromophobe Type",
                    "Renal Tubular Cell Carcinoma",
                    "Translocation Associated Renal Cell Carcinoma",
                    "Xp11.2 Translocation",
                ]
                else "Urothelial Carcinoma"
                if subtype in ["Urothelial Carcinoma"]
                else "Renal Cell Carcinoma"
                if subtype
                in [
                    "Parenchymal Carcinoma - Most In Keeping With Papillary Renal Cell Carcinoma"
                ]
                else "Transitional Cell Carcinoma"
                if subtype in ["Trans. Cell Carc."]
                else "Adenocarcinoma"
                if subtype in ["Adc"]
                else "No Cancer"
                if subtype in ["Normal", "Adjacent Normal"]
                else "Cancer"
                if diagnosis in ["Kidney Cancer"] and subtype in ["-"]
                else "NR"
            ),
            subtype=(
                "Papillary Renal Cell Carcinoma"
                if subtype
                in [
                    "Papillary Carcinoma",
                    "Papillary Renal Carcinoma",
                    "Papillary Renal Cell Carcinoma",
                    "Prcc",
                    "Papillary Carcinoma",
                ]
                else "Papillary Renal Cell Carcinoma (2X)"
                if subtype in ["2X Papillary Renal Cell Carcinomas"]
                else "Papillary Renal Cell Carcinoma With Multiple Foci"
                if subtype in ["Multifocal Papillary Renal Cell Carcinoma"]
                else "Papillary Renal Cell Carcinoma (Type 1)"
                if subtype in ["Papillary, Type 1"]
                else "Papillary Renal Cell Carcinoma (Type 2)"
                if subtype
                in ["Papillary Type 2", "Papillary, Type 2", "Rcc Papillary Type Ii"]
                else "Clear Cell Renal Cell Carcinoma"
                if subtype
                in [
                    "Ccc",
                    "Clear Cell Carcinoma",
                    "Clear Cell Renal Carcinoma",
                    "Clear Cell Renal Cell Carcinoma",
                    "Clear Cell Type",
                    "Conventional (Clear Cell) Renal Carcinoma",
                    "Conventional (Clear Cell) Renal Cell Carcinoma",
                    "Conventional Clear Cell Renal Cell Carcinoma",
                ]
                else "Renal Cell Carcinoma (NOS)"
                if subtype
                in [
                    "Conventional Renal Cell Carcinoma",
                    "Conventional Renal Cell Carcinoma",
                    "Rcc Conventiol",
                    "Rcc",
                    "Renal Cell Carcinoma (Type Uncertain)",
                    "Renal Tubular Cell Carcinoma",
                ]
                else "Clear Cell Renal Cell Carcinoma With Foci Of Rhabdoid Change"
                if subtype
                in [
                    "Clear Cell Renal Cell Carcinoma; Foci Of Rhabdoid Change",
                    "Clear Cell Renal Carcinoma, Focal Area Showing Rhabdoid Cells",
                ]
                else "Clear Cell Renal Cell Carcinoma With Focal Sarcomatoid Morphology And Coagulative Necrosis"
                if subtype
                in [
                    "Conventional Clear Cell Renal Cell Carcinoma With Focal Sarcomatoid Morphology And Coagulative Necrosis"
                ]
                else "Clear Cell Renal Cell Carcinoma With Sarcomatoid Differentiation"
                if subtype in ["Ccc Sarcomatoid"]
                else "Chromophobe Renal Cell Carcinoma"
                if subtype
                in [
                    "Chromophobe",
                    "Chromophobe Carcinoma",
                    "Chromophobe Carcinoma, Confined To Renal  Parenchyma",
                    "Chromophobe Cell Renal Cell Carcinoma",
                    "Chromophobe Renal Cell Carcinoma",
                    "Rc Cromophobe",
                    "Renal Cell Carcinoma, Chromophobe Cell Type",
                    "Renal Cell Carcinoma, Chromophobe Type",
                ]
                else "Renal Cell Carcinoma With Clear Cell Areas"
                if subtype
                in ["Renal Cell Carcinoma (Unclassified; With Clear Cell Areas)"]
                else "Renal Cell Carcinoma With Sarcomatoid Differentiation"
                if subtype in ["Renal Cell Carcinoma With Sarcomatoid Differentiation"]
                else "Translocation Associated Renal Cell Carcinoma"
                if subtype in ["Translocation Associated Renal Cell Carcinoma"]
                else "Translocation Associated Renal Cell Carcinoma (XP11.2)"
                if subtype in ["Xp11.2 Translocation"]
                else "No Cancer"
                if subtype in ["Normal", "Adjacent Normal", "Matched Normal"]
                else "Papillary Renal Cell Carcinoma"
                if subtype
                in [
                    "Parenchymal Carcinoma - Most In Keeping With Papillary Renal Cell Carcinoma"
                ]
                else "Adenocarcinoma (NOS)"
                if subtype in ["Adc"]
                else "NR"
            ),
            pairing_status=(
                "No Cancer"
                if subtype in ["Normal", "Adjacent Normal"]
                else "P"
                if subtype == "Matched Normal"
                or diagnosis
                in [
                    "Left And Normal Right Kidney",
                    "Kidney Pri. And Nor. And Met.",
                    "Kidney Pri. And Nor.",
                ]
                else "NR"
            ),
            pairing_site=(
                "No Cancer"
                if subtype in ["Normal", "Adjacent Normal"]
                else "Kidney"
                if subtype == "Matched Normal"
                or diagnosis
                in [
                    "Left And Normal Right Kidney",
                    "Kidney Pri. And Nor. And Met.",
                    "Kidney Pri. And Nor.",
                ]
                else "NA"
            ),
            i=i,
        )


# Export to Box & Load it to format the column width
kidney, kidney_inventory = format_and_export(file_name="kidney", tissue_source="Kidney")


# In[31]: Uterus - Endometrium (Verified)
keywords_femalereprodsystem = [
    "Ovarian",
    "Ovary",
    "Ovaries",
    "Fallopian Tube",
    "Salpinx",
    "Papillary Serous Carcinoma",
    "uterus",
    "Uterus",
    "Myometrium",
    "Placenta",
    "Endometrium",
    "Endometrial",
    "Endometrioid",
    "Cervical",
    "Cervix",
    "cervix",
]


for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Female Reproductive System":
        if diagnosis in ["Endometrium", "Endometrial Cancer", "Endometrium Pri."]:
            if subtype in [
                "Endometrioid Carcinoma",
                "Endometrioid Carcinoma With Squamous Differentiation",
                "Endometrioid Carcinoma With Mucinous Areas",
                "Endometrioid Carcinoma, Squamous Differentiation",
                "Endometrioid Carcinoma, With Areas Of Mucinous Differentiation",
                "Endometrioid Carcinoma; Areas Of Mucinous And Secretory Differentiation",
                "Endometrioid Carcinoma, Prominent Squamous Differentiation",
                "Endometrioid Carcinoma; Areas Of Squamous And Mucinous Differentiation",
            ]:
                inventory.loc[i, "Histological Diagnosis (Original)"] = "Adenocarcinoma"

                assign_values(
                    anatomical_location="Uterus",
                    anatomical_sublocation="Endometrium",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype=(
                        "Endometrioid Adenocarcinoma With Squamous Differentiation"
                        if any(
                            word in subtype
                            for word in [
                                "Carcinoma With Squamous Differentiation",
                                "Prominent Squamous Differentiation",
                                "Squamous Differentiation",
                            ]
                        )
                        else "Endometrioid Adenocarcinoma With Areas Of Mucinous And Secretory Differentiation"
                        if "; Areas Of Mucinous And Secretory Differentiation"
                        in subtype
                        else "Endometrioid Adenocarcinoma With Areas Of Mucinous Differentiation"
                        if any(
                            word in subtype
                            for word in [
                                "With Mucinous Areas",
                                ", With Areas Of Mucinous Differentiation",
                            ]
                        )
                        else "Endometrioid Adenocarcinoma With Areas Of Squamous and Mucinous Differentiation"
                        if "; Areas Of Squamous And Mucinous Differentiation" in subtype
                        else "Endometrioid Adenocarcinoma"
                    ),
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )
            elif subtype in [
                "Endometrioid Carcinoma With Focal Squamous Differentiation",
                "Endometrioid Carcinoma, Focal Squamous Differentiation",
            ]:
                assign_values(
                    anatomical_location="Uterus",
                    anatomical_sublocation="Endometrium",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype="Endometrioid Adenocarcinoma With Focal Squamous Differentiation",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

            elif subtype in [
                "Adc",
                "Endometrioid Adenocarcinoma With Mucinous And Secretory Differentiation",
                "Endometrioid Adenocarcinoma, Extensive Areas Of Squamous Differentiation",
                "Endometrioid Adenocarcinoma, With Several Squamous Morules",
                "Endometrioid Adenocarcinoma",
                "Endometrioid Adenocarcinoma, With Foci Of Mucinous Differentiation",
                "Endometrioid Adenocarcinoma, Foci Of Mucinous Differentiation And Focal Squamous Differentiation",
                "Endometrioid Adenocarcinoma With Focal Secretory Change",
                "Endometrioid Adenocarcinoma; Squamous Morule Formation",
                "Adenocarcinoma, Nos",
                "Endometrioid Adc",
                "Adc Endometrioide",
                "Endometrioid Carcinoma Of The Endometrium",
            ]:
                assign_values(
                    anatomical_location="Uterus",
                    anatomical_sublocation="Endometrium",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype=(
                        "Endometrioid Adenocarcinoma With Mucinous And Secretory Differentiation"
                        if "With Mucinous And Secretory Differentiation" in subtype
                        else "Endometrioid Adenocarcinoma With Areas of Squamous Differentiation"
                        if ", Extensive Areas Of Squamous Differentiation" in subtype
                        else "Endometrioid Adenocarcinoma With Squamous Morule Formation"
                        if any(
                            word in subtype
                            for word in [
                                "With Several Squamous Morules",
                                "Squamous Morule Formation",
                            ]
                        )
                        else "Endometrioid Adenocarcinoma With Foci Of Mucinous And Focal Squamous Differentiation"
                        if "Foci Of Mucinous Differentiation And Focal Squamous Differentiation"
                        in subtype
                        else "Endometrioid Adenocarcinoma With Focal Secretory Change"
                        if "With Focal Secretory Change" in subtype
                        else "Endometrioid Adenocarcinoma With Squamous Morule Formation and Areas Of Secretory Change"
                        if "Squamous Morules, Areas Of Secretory Change" in subtype
                        else "Endometrioid Adenocarcinoma With Foci Of Mucinous Differentiation"
                        if "With Foci Of Mucinous Differentiation" in subtype
                        else "Endometrioid Adenocarcinoma"
                        if subtype
                        in [
                            "Endometrioid Adc",
                            "Adc Endometrioide",
                            "Endometrioid Carcinoma Of The Endometrium",
                        ]
                        else "Adenocarcinoma (NOS)"
                    ),
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )
            elif subtype in [
                "Endometrioid Adenocarcinoma, Nos",
                "Endometrioid Adenocarcinoma, Nos; Squamous Morules, Areas Of Secretory Change",
            ]:
                assign_values(
                    anatomical_location="Uterus",
                    anatomical_sublocation="Endometrium",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype=(
                        "Endometrioid Adenocarcinoma With Squamous Morules And Areas Of Secretory Change"
                        if "Squamous Morules, Areas Of Secretory Change" in subtype
                        else "Endometrioid Adenocarcinoma (NOS)"
                    ),
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )
            elif subtype == "Squamous Cell Carcinoma":
                assign_values(
                    anatomical_location="Uterus",
                    anatomical_sublocation="Endometrium",
                    specimen_type="P",
                    diagnosis="Squamous Cell Carcinoma",
                    subtype="NR",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )
            elif subtype == "Undiffentiated Cancer With Solid Type Of Growth":
                assign_values(
                    anatomical_location="Uterus",
                    anatomical_sublocation="Endometrium",
                    specimen_type="P",
                    diagnosis="Undifferentiated Carcinoma",
                    subtype="Undifferentiated Carcinoma With Solid Type Of Growth",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )
            elif (
                subtype
                == "Malignant Mixed Mullerian Tumour With Rhabdomyoblastic Differentiation"
            ):
                assign_values(
                    anatomical_location="Uterus",
                    anatomical_sublocation="Endometrium",
                    specimen_type="P",
                    diagnosis="Carcinosarcoma (Malignant Mixed Mullerian Tumour)",
                    subtype="Malignant Mixed Mullerian Tumour With Rhabdomyoblastic Differentiation",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )
            elif subtype == "Carcinosarcoma":
                assign_values(
                    anatomical_location="Uterus",
                    anatomical_sublocation="Endometrium",
                    specimen_type="P",
                    diagnosis="Carcinosarcoma (NOS)",
                    subtype="NR",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )
            elif (
                subtype
                == "Dedifferentiated Endometrioid Carcinoma: Undifferentiated Carcinoma And Endometrioid Adenocarcinoma"
            ):
                assign_values(
                    anatomical_location="Uterus",
                    anatomical_sublocation="Endometrium",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype="Undifferentiated Carcinoma And Endometrioid Adenocarcinoma",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )
            elif subtype in [
                "Clear Cell Adenocarcinoma",
                "Clear Cell Carcinoma",
                "Clear Cell Adenocarcinoma, Nos",
            ]:
                assign_values(
                    anatomical_location="Uterus",
                    anatomical_sublocation="Endometrium",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype="Clear Cell Adenocarcinoma",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )
            elif subtype == "Enteropathy-Associated T-Cell L":
                assign_values(
                    anatomical_location="Uterus",
                    anatomical_sublocation="Endometrium",
                    specimen_type="P",
                    diagnosis="Enteropathy-Associated T-Cell Lymphoma",
                    subtype="NR",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )
            elif subtype == "/":
                assign_values(
                    anatomical_location="Uterus",
                    anatomical_sublocation="Endometrium",
                    specimen_type="NR",
                    diagnosis="NR",
                    subtype="NR",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Female Reproductive System":
        if diagnosis in ["Endometrium", "Endometrial Cancer", "Endometrium Pri."]:
            if subtype in [
                "Endometrioid/Hg Serous",
                "Hg Serous",
                "High Grade Serous Carcinoma With Clear Cell Change",
                "High-Grade Serous Adenocarcinoma Of The Endometrium",
                "High-Grade Serous Carcinoma",
                "High Grade Serous Endometrial Carcinoma",
                "High Grade Serous Carcinoma",
                "High Grade Serous Adenocarcinoma",
                "High Grade Endometrioid Carcinoma",
            ]:
                assign_values(
                    anatomical_location="Uterus",
                    anatomical_sublocation="Endometrium",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype=(
                        "High Grade Serous Adenocarcinoma With Clear Cell Change"
                        if "Clear Cell Change" in subtype
                        else "High Grade Serous Endometrioid Adenocarcinoma"
                        if all(
                            keyword in subtype for keyword in ["Serous", "Endometrioid"]
                        )
                        and "Clear Cell Change" not in subtype
                        else "High Grade Serous Adenocarcinoma"
                        if "Serous" in subtype and "Clear Cell Change" not in subtype
                        else "High Grade Adenocarcinoma"
                    ),
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )
            elif subtype in [
                "Serous Carcinoma Of Endometrium",
                "Serous Adenocarcinoma Of The Endometrium",
                "Serous Carcinoma",
                "Serous Adenocarcinoma",
                "Serous Papillary Carcinoma",
                "Papillary Serous Carcinoma",
            ]:
                assign_values(
                    anatomical_location="Uterus",
                    anatomical_sublocation="Endometrium",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype=(
                        "Serous Papillary Adenocarcinoma"
                        if "Papillary" in subtype
                        else "Serous Adenocarcinoma"
                    ),
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

            elif subtype in [
                "Carcinosarcoma With Heterologous Elements; Squamous And Rhabdomyosarcomatous Differentiation",
                "Carcinosarcoma, With Serous Component",
                "Carcinosarcoma With Chondroid Differentiation",
                "Carcinosarcoma With Foci Of Heterologous Differentiation",
                "Carcinosarcoma With Heterologous Differentiation",
                "Carcinosarcoma With Heterologous Elements",
                "Carcinosarcoma Showing Rhabdomyoblastic Differentiation",
                "Carcinosarcoma Of Endometrium, With Heterologous Differentiation",
                "Carcinosarcoma (With Rhabdomyoblastic Differentiation And High Grade Serous Epithelial Component)",
                "Carcinosarcoma Of Homologous Type",
                "Carcinosarcoma With Heterologus Elements",
            ]:
                assign_values(
                    anatomical_location="Uterus",
                    anatomical_sublocation="Endometrium",
                    specimen_type="P",
                    diagnosis=(
                        "Carcinosarcoma (Malignant Mixed Mullerian Tumour)"
                        if subtype
                        in [
                            "Carcinosarcoma With Heterologous Elements; Squamous And Rhabdomyosarcomatous Differentiation",
                            "Carcinosarcoma With Chondroid Differentiation",
                            "Carcinosarcoma With Foci Of Heterologous Differentiation",
                            "Carcinosarcoma With Heterologous Differentiation",
                            "Carcinosarcoma With Heterologous Elements",
                            "Carcinosarcoma Showing Rhabdomyoblastic Differentiation",
                            "Carcinosarcoma Of Endometrium, With Heterologous Differentiation",
                            "Carcinosarcoma (With Rhabdomyoblastic Differentiation And High Grade Serous Epithelial Component)",
                            "Carcinosarcoma With Heterologus Elements",
                        ]
                        else "Carcinosarcoma (NOS)"
                        if subtype
                        in [
                            "Carcinosarcoma Of Homologous Type",
                            "Carcinosarcoma, With Serous Component",
                        ]
                        else "NR"
                    ),
                    subtype=(
                        "Carcinosarcoma With Heterologous Differentiation"
                        if any(
                            word in subtype
                            for word in [
                                "With Heterologous Elements",
                                "With Heterologous Differentiation",
                                "With Heterologus Elements",
                            ]
                        )
                        else "Carcinosarcoma With Serous Component"
                        if "With Serous Component" in subtype
                        else "Carcinosarcoma With Chondroid Differentiation"
                        if "With Chondroid Differentiation" in subtype
                        else "Carcinosarcoma With Foci Of Heterologous Differentiation"
                        if "With Foci Of Heterologous Differentiation" in subtype
                        else "Carcinosarcoma With Rhabdomyoblastic Differentiation"
                        if "Carcinosarcoma Showing Rhabdomyoblastic Differentiation"
                        in subtype
                        else "Carcinosarcoma With Rhabdomyoblastic Differentiation And High Grade Serous Epithelial Component"
                        if "Carcinosarcoma With Rhabdomyoblastic Differentiation And High Grade Serous Epithelial Component"
                        in subtype
                        else "Carcinosarcoma With Homologous Differentiation"
                        if "Carcinosarcoma Of Homologous Type" in subtype
                        else "Carcinosarcoma With Heterologous, Squamous, and Rhabdomyosarcomatous Differentiation"
                        if "Carcinosarcoma With Heterologous Elements; Squamous And Rhabdomyosarcomatous Differentiation"
                        in subtype
                        else "Carcinosarcoma With Rhabdomyoblastic Differentiation And High Grade Serous Epithelial Component"
                        if subtype
                        == "Carcinosarcoma (With Rhabdomyoblastic Differentiation And High Grade Serous Epithelial Component)"
                        else "NRNR"
                    ),
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )
            elif subtype == "High Grade Carcinosarcoma Of Homologous Type":
                assign_values(
                    anatomical_location="Uterus",
                    anatomical_sublocation="Endometrium",
                    specimen_type="P",
                    diagnosis="Carcinosarcoma (Malignant Mixed Mullerian Tumour)",
                    subtype="High Grade Carcinosarcoma Of Homologous Type",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )


for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Female Reproductive System":
        if diagnosis in ["Endometrium", "Endometrial Cancer", "Endometrium Pri."]:
            if subtype in [
                "Mixed Endometrioid And Serous Papillary Carcinoma",
                "Mixed Carcinoma: Clear Cell And Endometrioid",
                "Mixed Serous And Endometrioid Adenocarcinoma",
                "Mixed Endometrioid And Serous Adenocarcinoma",
                "Mixed Endometrioid And Clear Cell Carcinoma; Foci Of Squamous Differentiation",
                "Mixed: High Grade Adenocarcinoma With Features Of Endometrioid And Serous Carcinoma",
                "Mixed High Grade Endometrioid And Serous Adenocarcinoma",
                "Mixed: Clear Cell And Serous Carcinoma",
            ]:
                assign_values(
                    anatomical_location="Uterus",
                    anatomical_sublocation="Endometrium",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype=(
                        "Mixed Endometrioid And Serous Papillary Carcinoma"
                        if "Endometrioid And Serous Papillary Carcinoma" in subtype
                        else "Mixed Endometrioid And Clear Cell Adenocarcinoma"
                        if "Clear Cell And Endometrioid" in subtype
                        else "Mixed High Grade Endometrioid And Serous Carcinoma"
                        if any(
                            word in subtype
                            for word in [
                                "Mixed High Grade Endometrioid And Serous Adenocarcinoma"
                            ]
                        )
                        else "Mixed Endometrioid And Serous Carcinoma"
                        if any(
                            word in subtype
                            for word in [
                                "Serous And Endometrioid Adenocarcinoma",
                                "Endometrioid And Serous Adenocarcinoma",
                            ]
                        )
                        else "Mixed Endometrioid And Clear Cell Adenocarcinoma With Foci Of Squamous Differentiation"
                        if "Mixed Endometrioid And Clear Cell Adenocarcinoma; Foci Of Squamous Differentiation"
                        in subtype
                        else "Mixed High Grade Adenocarcinoma With Features Of Endometrioid and Serous Carcinoma"
                        if "High Grade Adenocarcinoma With Features Of Endometrioid And Serous Carcinoma"
                        in subtype
                        else "Mixed Clear Cell And Serous Carcinoma"
                        if "Clear Cell And Serous Carcinoma" in subtype
                        else "Mixed Endometrioid And Clear Cell Adenocarcinoma With Foci Of Squamous Differentiation"
                        if "Mixed Endometrioid And Clear Cell Carcinoma; Foci Of Squamous Differentiation"
                        in subtype
                        else "NR"
                    ),
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

            elif subtype in [
                "Endometrioid Adenocarcinoma With Foci Of Clear Cell Differentiation",
                "Endometrioid Adenocarcinoma With Secretory Change",
                "Endometrioid Carcinoma Featuring Moderately Pleomorphic Neoplastic Cells",
            ]:
                assign_values(
                    anatomical_location="Uterus",
                    anatomical_sublocation="Endometrium",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype=(
                        "Endometrioid Adenocarcinoma With Foci Of Clear Cell Differentiation"
                        if "With Foci Of Clear Cell Differentiation" in subtype
                        else "Endometrioid Adenocarcinoma With Secretory Change"
                        if "With Secretory Change" in subtype
                        else "Endometrioid Adenocarcinoma With Moderately Pleomorphic Neoplastic Cells"
                        if "Featuring Moderately Pleomorphic Neoplastic Cells"
                        in subtype
                        else "NR"
                    ),
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )
            elif subtype == "Adenocarcinoma. Mixed Clear Cell/Endometriod":
                assign_values(
                    anatomical_location="Uterus",
                    anatomical_sublocation="Endometrium",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype="Mixed Clear Cell And Endometrioid Adenocarcinoma",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )
            elif subtype == "Primary":
                assign_values(
                    anatomical_location="Uterus",
                    anatomical_sublocation="Endometrium",
                    specimen_type="P",
                    diagnosis="Cancer",
                    subtype="NR",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )
            elif subtype == "Serous Carcinoma":
                assign_values(
                    anatomical_location="Uterus",
                    anatomical_sublocation="Endometrium",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype="Serous Adenocarcinoma",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )
            elif subtype == "Clear Cell Carcinoma Of Endometrium":
                assign_values(
                    anatomical_location="Uterus",
                    anatomical_sublocation="Endometrium",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype="Clear Cell Adenocarcinoma",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )
            elif subtype in ["Ca Endometrioide", "Ca. Endometrioide"]:
                assign_values(
                    anatomical_location="Uterus",
                    anatomical_sublocation="Endometrium",
                    specimen_type="P",
                    diagnosis="Adenocarcinoma",
                    subtype="Endometrioid Adenocarcinoma",
                    pairing_status="NR",
                    pairing_site="NA",
                    molecular_information="NR",
                    i=i,
                )

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Female Reproductive System":
        if diagnosis in [
            "Endometrium",
            "Endometrial Cancer",
            "Normal (Non-Malignant) Endometrium",
            "Endometrium Nor.",
        ]:
            if subtype in [
                "Normal Adjacent Tissues",
                "Adjacent Normal",
                "Normal Adjacent Tissues",
                "Normal",
            ]:
                assign_values(
                    anatomical_location="Uterus",
                    anatomical_sublocation="Endometrium",
                    specimen_type=("N (Adjacent)" if "Adjacent" in subtype else "N"),
                    diagnosis="No Cancer",
                    subtype="No Cancer",
                    molecular_information="NR",
                    pairing_status="No Cancer",
                    pairing_site="No Cancer",
                    i=i,
                )
            if diagnosis == "Normal (Non-Malignant) Endometrium" and subtype == "-":
                assign_values(
                    anatomical_location="Uterus",
                    anatomical_sublocation="Endometrium",
                    specimen_type="N",
                    diagnosis="No Cancer",
                    subtype="No Cancer",
                    molecular_information="NR",
                    pairing_status="No Cancer",
                    pairing_site="No Cancer",
                    i=i,
                )

            elif diagnosis == "Endometrial Cancer" and subtype == "-":
                assign_values(
                    anatomical_location="Uterus",
                    anatomical_sublocation="Endometrium",
                    specimen_type="P",
                    diagnosis="Cancer",
                    subtype="NR",
                    molecular_information="NR",
                    pairing_status="NR",
                    pairing_site="NA",
                    i=i,
                )

        elif diagnosis == "Endometrial Stromal Sarcoma":
            assign_values(
                anatomical_location="Uterus",
                anatomical_sublocation="Endometrium",
                specimen_type="P",
                diagnosis="Endometrial Stromal Sarcoma",
                subtype="NR",
                molecular_information="NR",
                pairing_site="NA",
                pairing_status="NR",
                i=i,
            )
    if (
        diagnosis
        == "Endometrioid Carcinoma; Areas Of Mucinous And Secretory Differentiation"
    ):
        if subtype == "-":
            assign_values(
                anatomical_location="Uterus",
                anatomical_sublocation="Endometrium",
                specimen_type="NR",
                diagnosis="Carcinoma",
                subtype="Endometrioid Carcinoma With Areas Of Mucinous And Secretory Differentiation",
                molecular_information="NR",
                pairing_site="NA",
                pairing_status="NR",
                i=i,
            )

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Female Reproductive System":
        if (
            diagnosis
            in [
                "Endometrium",
                "Endometrioid Carcinoma; Areas Of Mucinous And Secretory Differentiation",
            ]
        ) and (subtype == "-"):
            assign_values(
                anatomical_location="Uterus",
                anatomical_sublocation="Endometrium",
                specimen_type=(
                    "P"
                    if diagnosis
                    == "Endometrioid Carcinoma; Areas Of Mucinous And Secretory Differentiation"
                    else "NR"
                ),
                diagnosis=(
                    "Adenocarcinoma"
                    if diagnosis
                    == "Endometrioid Carcinoma; Areas Of Mucinous And Secretory Differentiation"
                    else "NR"
                ),
                subtype=(
                    "Endometrioid Adenocarcinoma With Areas Of Mucinous And Secretory Differentiation"
                    if diagnosis
                    == "Endometrioid Carcinoma; Areas Of Mucinous And Secretory Differentiation"
                    else "NR"
                ),
                molecular_information="NR",
                pairing_site="NA",
                pairing_status="NR",
                i=i,
            )

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Female Reproductive System":
        if diagnosis in ["Uterus", "Uterus Pri.", "Uterus  Pri."] and any(
            keyword in subtype for keyword in ["Endometrioid", "Endometrial"]
        ):
            assign_values(
                anatomical_location="Uterus",
                anatomical_sublocation="Endometrium",
                specimen_type="P",
                diagnosis=(
                    "Sarcoma"
                    if "Sarcoma" in subtype
                    else "Adenocarcinoma"
                    if any(
                        keyword in subtype
                        for keyword in [
                            "Adenocarcinoma",
                            "Adenocarinoma",
                            "Carinoma",
                            "Adc",
                            "Carcinoma",
                        ]
                    )
                    else "Mixed Carcinoma"
                    if subtype
                    in [
                        "Endometrial Carcinosarcoma Plus Small Cell (Neuroendocrine) Carcinoma"
                    ]
                    else "NR"
                ),
                subtype=(
                    "High Grade Endometrioid Sarcoma"
                    if "Endometrioid Sarcoma" in subtype
                    else "Endometrial Carcinosarcoma And Small Cell (Neuroendocrine) Carcinoma"
                    if subtype
                    in [
                        "Endometrial Carcinosarcoma Plus Small Cell (Neuroendocrine) Carcinoma"
                    ]
                    else "Endometrioid Adenocarcinoma"
                    if "Endometrioid" in subtype
                    else "NR"
                ),
                pairing_status="NR",
                pairing_site="NA",
                i=i,
            )


# In[32]: Uterus - Myometrium (Verified)

keys_to_keep = [
    "Uterus- Myometrium",
    "Uterus, Myometrium",
    "Myometrium Nor.",
    "Uterus-Myometrium",
    "Uterus Myometrium",
]

my_dict = create_combinations_dict("Female Reproductive System")
filtered_dict = {}
for key in keys_to_keep:
    if key in my_dict:
        filtered_dict[key] = my_dict[key]
filtered_dict


for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Female Reproductive System":
        if "Myometrium" in diagnosis:
            assign_values(
                anatomical_location="Uterus",
                anatomical_sublocation="Myometrium",
                specimen_type=(
                    "N"
                    if subtype in ["Normal"]
                    else "N (Adjacent)"
                    if subtype in ["Normal Adjacent"]
                    else "NR"
                ),
                diagnosis="No Cancer",
                subtype="No Cancer",
                pairing_status="No Cancer",
                pairing_site="No Cancer",
                i=i,
            )

# In[33]: Uterus - Cervix (Verified)

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Female Reproductive System":
        if any(
            keyword in diagnosis
            for keyword in ["Cervix", "cervix", "Cervical", "cervical"]
        ):
            assign_values(
                anatomical_location="Uterus",
                anatomical_sublocation=(
                    "Cervix"
                    if diagnosis in ["Exocervix", "Esocervix Nor."]
                    else "Cervix (Endocervix)"
                    if diagnosis in ["Endocervix"]
                    else "Cervix"
                ),
                specimen_type=(
                    "N"
                    if subtype in ["Normal"]
                    else "N (Adjacent)"
                    if subtype in ["Normal Adjacent", "Normal Adjacent Tissues"]
                    else "P"
                ),
                diagnosis=(
                    "Squamous Cell Carcinoma"
                    if (
                        (subtype in ["Squamous Carcinoma", "Scc"])
                        or (diagnosis in ["Cervical Scc"] and subtype in ["-"])
                    )
                    else "Adenocarcinoma"
                    if subtype
                    in [
                        "Endometrioid Adenocarcinoma",
                        "Endocervical Carcinoma",
                        "Adenocarcinoma",
                        "Cervical Adcendometrioid Type",
                        "Cervical Adc Endometrioid Type",
                        "Clear Cell Cervical Adc",
                        "Cervical Adc Endocervical Type",
                        "Cervical Adc",
                        "Cervical Adc Invasive",
                        "Cervical Adcendocervical Type",
                    ]
                    else "Squamous Cell Carcinoma"
                    if subtype
                    in [
                        "Cervical Carcinomaepidermioid",
                        "Cervical Carcinoma Epidermioid",
                    ]
                    else "Neuroendocrine Carcinoma"
                    if subtype in ["Cervical Carcinoma Neurendocrine"]
                    else "Carcinoma"
                    if diagnosis in ["Cervical Cancer"]
                    and all(
                        keyword not in subtype
                        for keyword in ["Carcinoma", "Adenocarcinoma"]
                    )
                    else "No Cancer"
                    if subtype
                    in ["Normal", "Normal Adjacent", "Normal Adjacent Tissues"]
                    else "Cancer"
                    if diagnosis in ["Cervical Pri"] and subtype in ["-"]
                    else "NR"
                ),
                subtype=(
                    "Endometrioid Adenocarcinoma"
                    if subtype
                    in [
                        "Endometrioid Adenocarcinoma",
                        "Cervical Adcendometrioid Type",
                        "Cervical Adc Endometrioid Type",
                    ]
                    else "Clear Cell Adenocarcinoma"
                    if subtype in ["Clear Cell Cervical Adc"]
                    else "Endocervical Adenocarcinoma"
                    if subtype
                    in [
                        "Cervical Adc Endocervical Type",
                        "Cervical Adcendocervical Type",
                        "Endocervical Carcinoma",
                    ]
                    else "Epidermioid Adenocarcinoma"
                    if subtype
                    in [
                        "Cervical Carcinomaepidermioid",
                        "Cervical Carcinoma Epidermioid",
                    ]
                    else "Invasive Adenocarcinoma"
                    if subtype in ["Cervical Adc Invasive"]
                    else "No Cancer"
                    if subtype
                    in ["Normal", "Normal Adjacent", "Normal Adjacent Tissues"]
                    else "NR"
                ),
                pairing_status=(
                    "No Cancer"
                    if subtype
                    in ["Normal", "Normal Adjacent", "Normal Adjacent Tissues"]
                    else "NR"
                ),
                pairing_site=(
                    "No Cancer"
                    if subtype
                    in ["Normal", "Normal Adjacent", "Normal Adjacent Tissues"]
                    else "NA"
                ),
                i=i,
            )


# In[34]: Uterus (Specific Location Not Provided)
for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Female Reproductive System":
        if diagnosis in [
            "Uterus",
            "Uterus Cancer",
            "Uterus (Non Malignant)",
        ] and subtype in [
            "Adjacent Normal",
            "-",
            "Serous Adenocarcinoma",
            "Carcinosarcoma",
            "Normal Adjacent Tissues",
            "Cancer",
            "Normal Adjacent",
        ]:
            assign_values(
                anatomical_location="Uterus",
                anatomical_sublocation="NR",
                specimen_type=(
                    "N (Adjacent)"
                    if subtype
                    in ["Adjacent Normal", "Normal Adjacent", "Normal Adjacent Tissues"]
                    else "P"
                ),
                diagnosis=(
                    "Adenocarcinoma"
                    if subtype in ["Serous Adenocarcinoma"]
                    else "Carcinosarcoma"
                    if subtype in ["Carcinosarcoma"]
                    else "No Cancer"
                    if subtype
                    in ["Adjacent Normal", "Normal Adjacent", "Normal Adjacent Tissues"]
                    else "Cancer"
                    if subtype in ["-", "Cancer"]
                    else "NR"
                ),
                subtype=(
                    "Serous Adenocarcinoma"
                    if subtype in ["Serous Adenocarcinoma"]
                    else "No Cancer"
                    if subtype
                    in ["Adjacent Normal", "Normal Adjacent", "Normal Adjacent Tissues"]
                    else "NR"
                ),
                pairing_status=(
                    "No Cancer"
                    if subtype
                    in ["Adjacent Normal", "Normal Adjacent", "Normal Adjacent Tissues"]
                    else "NR"
                ),
                pairing_site=(
                    "No Cancer"
                    if subtype
                    in ["Adjacent Normal", "Normal Adjacent", "Normal Adjacent Tissues"]
                    else "NA"
                ),
                i=i,
            )

# In[35]: Placenta (Verified)
for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Female Reproductive System":
        if "Placenta" in diagnosis:
            assign_values(
                anatomical_location="Placenta",
                anatomical_sublocation=(
                    "Placenta (Early)" if "Early" in diagnosis else "NR"
                ),
                specimen_type="N",
                diagnosis="No Cancer",
                subtype="No Cancer",
                pairing_status="No Cancer",
                pairing_site="No Cancer",
                i=i,
            )


# In[36]: Ovaries (Verified)
for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Female Reproductive System":
        if (
            any(
                keyword in diagnosis
                for keyword in [
                    "Ovary",
                    "Ovaries",
                    "Ovarian",
                    "ovary",
                    "ovaries",
                    "ovarian",
                ]
            )
            or "Ovary" in subtype
        ):
            assign_values(
                anatomical_location=(
                    "Ovary/Peritoneum"
                    if diagnosis in ["Ovary/Peritoneal"]
                    else "Ovary/Fallopian Tube"
                    if diagnosis in ["Left Ovary + Tube", "Left Tube And Ovary"]
                    else "Ovary"
                    if "Ovary" in diagnosis
                    else "Ovaries"
                ),
                anatomical_sublocation=(
                    "Ovary (Left)"
                    if diagnosis in ["Ovary (Left)", "Left Ovary"]
                    else "Ovary (Right)"
                    if diagnosis in ["Right Ovary"]
                    else "Ovaries (Bilateral)"
                    if diagnosis in ["Bilateral Ovaries"]
                    else "Ovaries (Biopsy)"
                    if diagnosis in ["Ovaries-Biopsy"]
                    else "Ovary (Left) And Fallopian Tube"
                    if diagnosis in ["Left Ovary + Tube", "Left Tube And Ovary"]
                    else "Ovary And Peritoneum"
                    if diagnosis in ["Ovary/Peritoneal"]
                    else "NR"
                ),
                specimen_type=(
                    "N (Adjacent)"
                    if "Adjacent" in subtype
                    else "N"
                    if "Normal" in subtype
                    else "M"
                    if any(
                        keyword in subtype
                        for keyword in ["Metastatic", "Metastasis", "Matched Met"]
                    )
                    else "NR"
                    if diagnosis in ["Ovaries", "Ovarian"] and subtype in ["-"]
                    else "P"
                ),
                diagnosis=(
                    "Adenocarcinoma"
                    if (
                        any(
                            keyword in subtype
                            for keyword in [
                                "Adenocarcinoma",
                                "Adc",
                                "Cysto Adc",
                                "Cystadenocarcinoma",
                                "Hg",
                                "High Grade",
                                "Lg",
                                "Low Grade",
                                "Papillary Serous Carcinoma",
                            ]
                        )
                        or subtype
                        in [
                            "Endometrioid Carcinoma",
                            "Serous Carcinoma",
                            "Carcinoma (Clear Cell)",
                            "Carcinoma (Papillary Serous)",
                            "Carcinoma (Serous Endometrioid)",
                            "Carcinoma (Serous)",
                            "Clear Cell",
                            "Endometrioid",
                            "High  Grade Serous",
                            "Mixed Endometrioid And Serous",
                            "Mixed Serous/Clear Cell",
                            "Mixed Serous/Endometrioid",
                            "Mucinous",
                            "Serous",
                            "Serous/Endometrioid",
                            "Solid Type Carcinoma",
                        ]
                    )
                    and "Neuroendocrine" not in subtype
                    else "Adenocarcinoma"
                    if diagnosis in ["Papillary Serous Carcinoma"]
                    else "Carcinoma"
                    if subtype in ["Ovarian Carcinoma", "Carcinoma"]
                    else "Renal Cell Carcinoma"
                    if subtype in ["Metastatic Renal Cell Carcinoma"]
                    else "No Cancer"
                    if any(keyword in subtype for keyword in ["Adjacent", "Normal"])
                    else "Carcinosarcoma"
                    if subtype in ["Carcinosarcoma"]
                    else "Cancer"
                    if subtype in ["Mixed Cell"]
                    else "Brenner Tumour"
                    if subtype
                    in ["Malignant Brenner Tumour", "Benign Brenners", "Brenner"]
                    else "Colorectal Cancer"
                    if subtype in ["Crc_Matched Met"]
                    else "Neuroendocrine Carcinoma"
                    if subtype
                    in [
                        "Primary High Grade Neuroendocrine Carcinoma Of The Fallopian Tube"
                    ]
                    else "Cancer"
                    if diagnosis in ["Ovarian Cancer"] and subtype in ["-"]
                    else "Borderline Tumour"
                    if subtype in ["Serous Borderline"]
                    else "Cancer"
                    if diagnosis == "Ovarian Cancer"
                    and subtype in ["No Information", "Nr"]
                    else "NR"
                ),
                subtype=(
                    "Mucinous Cystadenocarcinoma"
                    if subtype in ["Mucinous Cystadenocarcinoma"]
                    else "Cystadenocarcinoma"
                    if subtype
                    in ["Cystadenocarcinoma", "Cysto Adc", "Ovarian Cystadenocarcinoma"]
                    else "Serous Cystadenocarcinoma"
                    if subtype in ["Serous Cystadenocarcinoma", "Cysto Adc Serous"]
                    else "Clear Cell Adenocarcinoma"
                    if subtype
                    in ["Adc Clear Cell", "Carcinoma (Clear Cell)", "Clear Cell"]
                    else "Serous Papillary Adenocarcinoma"
                    if subtype
                    in [
                        "Adc Papillary Serous",
                        "Adc Sp",
                        "Carcinoma (Papillary Serous)",
                        "Serous Papillary Adenocarcinoma",
                    ]
                    or diagnosis in ["Papillary Serous Carcinoma"]
                    else "Papillary Solid Trabecular Adenocarcinoma"
                    if subtype in ["Adc Papillary Solid Trabecular"]
                    else "Serous Adenocarcinoma"
                    if subtype
                    in [
                        "Adc Serous",
                        "Carcinoma (Serous)",
                        "Serous",
                        "Serous Adenocarcinoma",
                        "Serous Carcinoma",
                    ]
                    else "Brenner Tumour (Benign)"
                    if subtype in ["Benign Brenners"]
                    else "Mixed Serous And Endometrioid Adenocarcinoma"
                    if subtype
                    in [
                        "Carcinoma (Serous Endometrioid)",
                        "Mixed Endometrioid And Serous",
                        "Mixed Serous/Endometrioid",
                        "Serous/Endometrioid",
                    ]
                    else "Serous Papillary Cystadenocarcinoma"
                    if subtype
                    in [
                        "Cysto Adc Papillary Serous",
                        "Serous Papillary Cystadenocarcinoma",
                    ]
                    else "Endometrioid Adenocarcinoma"
                    if subtype
                    in [
                        "Endometrioid",
                        "Endometrioid Adenocarcinoma",
                        "Endometrioid Carcinoma",
                    ]
                    else "Endometrioid Cystadenocarcinoma"
                    if subtype in ["Endometrioid Cystadenocarcinoma"]
                    else "High Grade Clear Cell Adenocarcinoma"
                    if subtype in ["Hg Clear Cell", "High Grade Clear Cell"]
                    else "High Grade Serous Adenocarcinoma"
                    if subtype
                    in [
                        "Hg Serous",
                        "High Grade Serous",
                        "High  Grade Serous",
                        "High Grades Serous",
                    ]
                    else "High Grade Endometrioid Adenocarcinoma"
                    if subtype in ["Hgendometrioid"]
                    else "High Grade Serous Clear Cell Adenocarcinoma"
                    if subtype in ["High Grade Serous/Clear Cell"]
                    else "High Grade Serous Endometrioid Adenocarcinoma"
                    if subtype in ["High Grade Serous/Endometrioid"]
                    else "Low Grade Serous Adenocarcinoma"
                    if subtype in ["Lg Serous", "Low Grade Serous", "Low Grade Seorus"]
                    else "Brenner Tumour (Malignant)"
                    if subtype in ["Malignant Brenner Tumour"]
                    else "Adenocarcinoma (Krukenberg Tumour)"
                    if subtype in ["Metastatic Adenocarcinoma (Krukenberg Tumour)"]
                    else "Mixed High Grade Serous And Endometrioid Adenocarcinoma"
                    if subtype in ["Mixed Hg Serous/Endometrioid"]
                    else "Mixed Serous And Clear Cell Adenocarcinoma"
                    if subtype in ["Mixed Serous/Clear Cell"]
                    else "Mucinous Adenocarcinoma"
                    if subtype in ["Mucinous", "Mucinous Adenocarcinoma"]
                    else "No Cancer"
                    if any(keyword in subtype for keyword in ["Adjacent", "Normal"])
                    else "High Grade Neuroendocrine Carcinoma Of The Fallopian Tube"
                    if subtype
                    in [
                        "Primary High Grade Neuroendocrine Carcinoma Of The Fallopian Tube"
                    ]
                    else "High Grade Serous Adenocarcinoma"
                    if subtype in ["Serous Carcinoma, High Grade"]
                    else "Solid Type Adenocarcinoma"
                    if subtype in ["Solid Type Carcinoma"]
                    else "Serous Borderline Tumour"
                    if subtype in ["Serous Borderline"]
                    else "Mixed Cell"
                    if subtype in ["Mixed Cell"]
                    else "NR"
                ),
                pairing_status=(
                    "P"
                    if diagnosis in ["Ovarian Pri. And Met."]
                    or subtype in ["Crc_Matched Met"]
                    else "No Cancer"
                    if any(keyword in subtype for keyword in ["Adjacent", "Normal"])
                    else "NR"
                ),
                pairing_site=(
                    "Colon/Rectum"
                    if subtype in ["Crc_Matched Met"]
                    else "Ovaries"
                    if diagnosis in ["Ovarian Pri. And Met."]
                    else "Kidney"
                    if subtype in ["Metastatic Renal Cell Carcinoma"]
                    else "No Cancer"
                    if any(keyword in subtype for keyword in ["Adjacent", "Normal"])
                    else "NA"
                ),
                i=i,
            )


# In[37]: Salpinx/Fallopian Tube (Verified)
for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Female Reproductive System":
        if (
            any(
                keyword in diagnosis
                for keyword in ["Saplinx", "Fallopian Tube", "Salpinx"]
            )
            or "Salpinx" in subtype
        ):
            assign_values(
                anatomical_location="Fallopian Tube",
                anatomical_sublocation="NR",
                specimen_type=(
                    "N (Adjacent)"
                    if "Adjacent" in subtype
                    else "N"
                    if "Normal" in subtype
                    else "P"
                ),
                diagnosis="No Cancer" if "Normal" in subtype else "Carcinoma",
                subtype=(
                    "No Cancer"
                    if "Normal" in subtype
                    else "Serous Papillary Adenocarcinoma"
                    if diagnosis in ["Papillary Serous Carcinoma"]
                    else "High Grade Serous Carcinoma"
                ),
                pairing_status="No Cancer" if "Normal" in subtype else "NR",
                pairing_site="No Cancer" if "Normal" in subtype else "NR",
                i=i,
            )


# Export to Box & Load it to format the column width
female_reproductive_system, female_reproductive_system_inventory = format_and_export(
    file_name="female_reproductive_system", tissue_source="Female Reproductive System"
)


# In[38]: Male Reproductive System (Verified)

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Male Reproductive System":
        assign_values(
            anatomical_location=("Penis" if diagnosis == "Penis" else "Testis"),
            anatomical_sublocation="NR",
            specimen_type="N"
            if "Nor." in diagnosis or any(keyword in subtype for keyword in ["Normal"])
            else "P",
            diagnosis=(
                "Neuroendocrine Carcinoma"
                if subtype
                in [
                    "Neuroendocrine Carcinoma (Large Cell); Presumed Progression Of Urothelial Carcinoma Of Bladder"
                ]
                else "Peripheral B-Cell Lymphoma"
                if subtype in ["Dlbcl"]
                else "Germ Cell Tumour"
                if subtype in ["Ca.Embrionale", "Ca. Germinale", "Seminoma"]
                else "No Cancer"
                if "Nor." in diagnosis
                or any(keyword in subtype for keyword in ["Normal"])
                else "Mixed Germ Cell Tumour"
                if subtype in ["Tum. Misto"]
                else "NR"
            ),
            subtype=(
                "Large Cell Neuroendocrine Carcinoma (Presumed Progression Of Bladder Urothelial Carcinoma)"
                if subtype
                in [
                    "Neuroendocrine Carcinoma (Large Cell); Presumed Progression Of Urothelial Carcinoma Of Bladder"
                ]
                else "Diffuse Large B-Cell Lymphoma"
                if subtype in ["Dlbcl"]
                else "Embryonal Carcinoma"
                if subtype in ["Ca.Embrionale"]
                else "Seminoma"
                if subtype in ["Seminoma"]
                else "No Cancer"
                if "Nor." in diagnosis
                or any(keyword in subtype for keyword in ["Normal"])
                else "NR"
            ),
            pairing_status="No Cancer"
            if "Nor." in diagnosis or any(keyword in subtype for keyword in ["Normal"])
            else "NR",
            pairing_site="No Cancer"
            if "Nor." in diagnosis or any(keyword in subtype for keyword in ["Normal"])
            else "NA",
            i=i,
        )


# Export to Box & Load it to format the column width
malereprod, malereprod_inventory = format_and_export(
    file_name="male_reproductive_system", tissue_source="Male Reproductive System"
)


# In[39]: Omentum (Verified)
for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Omentum":
        assign_values(
            anatomical_location="Omentum",
            anatomical_sublocation="NR",
            specimen_type=(
                "N"
                if subtype in ["Normal Adjacent"]
                else "M"
                if subtype
                in [
                    "Metastatic High Grade Neuroendocrine Carcinoma Of Small Cell Type (Unknown Primary)",
                    "Metastasis",
                ]
                else "P"
            ),
            diagnosis=(
                "Adenocarcinoma"
                if subtype in ["Hg Serous", "High Grade Serous"]
                else "Neuroendocrine Carcinoma"
                if subtype
                in [
                    "Metastatic High Grade Neuroendocrine Carcinoma Of Small Cell Type (Unknown Primary)"
                ]
                else "No Cancer"
                if subtype in ["Normal Adjacent"]
                else "Peripheral B-Cell Lymphoma"
                if subtype in ["Dlbcl Various Site"]
                else "NR"
            ),
            subtype=(
                "High Grade Serous Adenocarcinoma"
                if subtype in ["Hg Serous", "High Grade Serous"]
                else "Small Cell Neuroendocrine Carcinoma"
                if subtype
                in [
                    "Metastatic High Grade Neuroendocrine Carcinoma Of Small Cell Type (Unknown Primary)"
                ]
                else "No Cancer"
                if subtype in ["Normal Adjacent"]
                else "Diffuse Large B-Cell Lymphoma"
                if subtype in ["Dlbcl Various Site"]
                else "NR"
            ),
            pairing_status=("No Cancer" if subtype in ["Normal Adjacent"] else "NR"),
            pairing_site=("No Cancer" if subtype in ["Normal Adjacent"] else "NR"),
            i=i,
        )


# Export to Box & Load it to format the column width
omentum, omentum_inventory = format_and_export(
    file_name="omentum", tissue_source="Omentum"
)

# In[40]: Urinary System (Verified)

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Urinary System":
        assign_values(
            anatomical_location=(
                "Urinary Bladder"
                if diagnosis
                in [
                    "Bladder",
                    "Bladder Carcinoma",
                    "Urinary Bladder",
                    "Urinary Bladder,Uroth.",
                    "Bladder And Met.",
                    "Bladder Diverticulum And Met.",
                    "Bladder Pri.",
                    "Bladder Nor.",
                    "Urinary Bladder, Uroth",
                    "Uriry Bladder,Uroth.",
                ]
                else "Ureter"
                if diagnosis
                in [
                    "Normal Ureter",
                    "Right Proximal Ureter",
                    "Ureter",
                    "Ureter And Met.",
                    "Ureter Nor.",
                ]
                else "Urothelium Tract"
                if diagnosis in ["Upper Urotelium Tract"]
                else "Prostate/Bladder"
                if diagnosis == "Prostate/Bladder"
                else "NR"
            ),
            anatomical_sublocation=(
                "Ureter (Right Proximal)"
                if diagnosis in ["Right Proximal Ureter"]
                else "Urinary Bladder (Diverticulum)"
                if diagnosis in ["Bladder Diverticulum And Met."]
                else "Urothelium Tract (Upper)"
                if diagnosis in ["Upper Urotelium Tract"]
                else "NR"
            ),
            specimen_type=(
                "N"
                if diagnosis in ["Normal Ureter", "Bladder Nor.", "Ureter Nor."]
                or subtype in ["Normal"]
                else "NR"
                if (subtype in ["Cui"] and "Pri" not in diagnosis)
                or (diagnosis == "Bladder" and subtype == "-")
                else "P"
            ),
            diagnosis=(
                "Urothelial Carcinoma"
                if subtype
                in [
                    "Urothelial Papillar Carcinoma",
                    "Uc",
                    "Urothelial Carcinoma",
                    "Met Blocks/Urothelial Carcinoma",
                    "Papillary Urothelial Carcinoma",
                    "Invasive Uc",
                    "Non Invasive Uc",
                    "Urothelial Carcinoma With Sarcomatoid Change",
                    "Urinary Bladder, Urothelium",
                ]
                and diagnosis not in ["Bladder Nor."]
                else "Sarcomatoid Malignant Tumour"
                if subtype in ["Sarcomatoid Malignant Tumour"]
                else "Urothelial (Transitional Cell) Carcinoma"
                if subtype
                in [
                    "Papillary Transitional Cell Carcinoma, With Focal Lamina Propria Invasion",
                    "Papillary Transitional Cell Carcinoma (Minority Squamous)",
                    "Transitional Cell Carcinoma",
                    "Trans. Cell Carc.",
                    "Tcc",
                    "Tcc With Squamous Cell Differentiation Invading Left Prostate",
                ]
                else "Neuroendocrine Carcinoma"
                if subtype
                in [
                    "Small Cell Bladder Carcinoma",
                    "Neuroendocrine Carcinoma, With Focal Features Suggestive Of Small Cell Carcinoma",
                    "Neuroendocrine Carcinoma Of The Urinary Bladder With Features In Places Suggestive Of Small Cell Carcinoma",
                    "Small Cell Carcinoma",
                ]
                else "Carcinoma"
                if subtype
                in [
                    "High Grade Invasive Carcinoma; Neuroendocrine Differentiation (Small Cell Carcinoma)",
                    "Carcinoma",
                ]
                or (diagnosis == "Bladder Pri." and subtype in ["/", "Cui"])
                else "No Cancer"
                if diagnosis in ["Normal Ureter", "Bladder Nor.", "Ureter Nor."]
                or subtype in ["Normal"]
                else "Squamous Cell Carcinoma"
                if subtype in ["Scc"]
                else "Urothelial Carcinoma"
                if subtype in ["Tcc With Cis"]
                else "Continuous Urinary Incontinence"
                if subtype in ["Cui"]
                else "Carcinoma"
                if diagnosis == "Bladder Carcinoma" and subtype == "-"
                else "Adenocarcinoma"
                if subtype
                in [
                    "Prostatic Carcinoma With Areas Of Neuroendocrine Differentiation (No More Than 30%)",
                    "Prostatic Carcinoma With Widespread Neuroendocrine Differentiation (Gleason Grade Provided For Small Acinar Type Area)",
                ]
                else "NR"
            ),
            subtype=(
                "Papillary Urothelial Carcinoma"
                if subtype
                in ["Urothelial Papillar Carcinoma", "Papillary Urothelial Carcinoma"]
                else "Papillary Urothelial (Transitional Cell) Carcinoma With Focal Lamina Propria Invasion"
                if subtype
                in [
                    "Papillary Transitional Cell Carcinoma, With Focal Lamina Propria Invasion"
                ]
                else "Papillary Urothelial (Transitional Cell) Carcinoma With Minor Squamous Differentiation"
                if subtype
                in ["Papillary Transitional Cell Carcinoma (Minority Squamous)"]
                else "Invasive Urothelial Carcinoma"
                if subtype in ["Invasive Uc"]
                else "Non-Invasive Urothelial Carcinoma"
                if subtype in ["Non Invasive Uc"]
                else "Urothelial Carcinoma With Sarcomatoid Change"
                if subtype in ["Urothelial Carcinoma With Sarcomatoid Change"]
                else "Small Cell Carcinoma"
                if subtype in ["Small Cell Bladder Carcinoma", "Small Cell Carcinoma"]
                else "Neuroendocrine Carcinoma With Features Of Small Cell Carcinoma"
                if subtype
                in [
                    "Neuroendocrine Carcinoma, With Focal Features Suggestive Of Small Cell Carcinoma",
                    "Neuroendocrine Carcinoma Of The Urinary Bladder With Features In Places Suggestive Of Small Cell Carcinoma",
                ]
                else "High-Grade Invasive Carcinoma With Neuroendocrine Differentiation (Small Cell Carcinoma)"
                if subtype
                in [
                    "High Grade Invasive Carcinoma; Neuroendocrine Differentiation (Small Cell Carcinoma)"
                ]
                else "No Cancer"
                if diagnosis in ["Normal Ureter", "Bladder Nor.", "Ureter Nor."]
                or subtype in ["Normal"]
                else "Urothelial (Transitional Cell) Carcinoma And Carcinoma In Situ"
                if subtype in ["Tcc With Cis"]
                else "Urothelial (Transitional Cell) Carcinoma With Squamous Cell Differentiation Invading Left Prostate"
                if subtype
                in ["Tcc With Squamous Cell Differentiation Invading Left Prostate"]
                else "Adenocarcinoma With Areas Of Neuroendocrine Differentiation (No More Than 30%)"
                if subtype
                in [
                    "Prostatic Carcinoma With Areas Of Neuroendocrine Differentiation (No More Than 30%)"
                ]
                else "Adenocarcinoma With Widespread Neuroendocrine Differentiation"
                if subtype
                in [
                    "Prostatic Carcinoma With Widespread Neuroendocrine Differentiation (Gleason Grade Provided For Small Acinar Type Area)"
                ]
                else "NR"
            ),
            pairing_status=(
                "No Cancer"
                if diagnosis in ["Normal Ureter", "Bladder Nor.", "Ureter Nor."]
                or subtype in ["Normal"]
                else "P"
                if diagnosis
                in [
                    "Bladder And Met.",
                    "Bladder Diverticulum And Met.",
                    "Ureter And Met.",
                ]
                else "NR"
            ),
            pairing_site=(
                "No Cancer"
                if diagnosis in ["Normal Ureter", "Bladder Nor.", "Ureter Nor."]
                or subtype in ["Normal"]
                else "NA"
            ),
            i=i,
        )


# Export to Box & Load it to format the column width
urinary_system, urinary_system_inventory = format_and_export(
    file_name="urinary_system", tissue_source="Urinary System"
)


# In[41]: Pelvis (Verified)
for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Pelvis":
        assign_values(
            anatomical_location=(
                "Seminal Vesicle" if "Seminal Vesicle" in diagnosis else "Pelvis"
            ),
            anatomical_sublocation=(
                "Pelvis (Mass Biopsy)"
                if diagnosis in ["Pelvic Mass Biopsy", "Pelvis"]
                else "NR"
            ),
            specimen_type=(
                "N"
                if subtype in ["Normal"]
                or diagnosis in ["Seminal Vesicle Normal (Non-Malignant Blocks)"]
                else "N (Adjacent)"
                if subtype in ["Normal Adjacent Tissues", "Adjacent Normal"]
                else "P"
            ),
            diagnosis=(
                "No Cancer"
                if subtype in ["Normal Adjacent Tissues", "Adjacent Normal", "Normal"]
                or diagnosis in ["Seminal Vesicle Normal (Non-Malignant Blocks)"]
                else "Clear Cell Carcinoma"
                if subtype in ["Clear Cell"]
                else "High Grade Serous Carcinoma"
                if subtype in ["High Grade Serous"]
                else "Neuroblastoma"
                if subtype in ["Neuroblastoma"]
                else "NR"
            ),
            subtype=(
                "No Cancer"
                if subtype in ["Normal Adjacent Tissues", "Adjacent Normal", "Normal"]
                or diagnosis in ["Seminal Vesicle Normal (Non-Malignant Blocks)"]
                else "NR"
            ),
            pairing_status=(
                "No Cancer"
                if subtype in ["Normal", "Normal Adjacent Tissues", "Adjacent Normal"]
                or diagnosis in ["Seminal Vesicle Normal (Non-Malignant Blocks)"]
                else "NR"
            ),
            pairing_site=(
                "No Cancer"
                if subtype in ["Normal", "Normal Adjacent Tissues", "Adjacent Normal"]
                or diagnosis in ["Seminal Vesicle Normal (Non-Malignant Blocks)"]
                else "NR"
            ),
            i=i,
        )


# Export to Box & Load it to format the column width
pelvis, pelvis_inventory = format_and_export(file_name="pelvis", tissue_source="Pelvis")


# In[42]: Chest (Verified)

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Chest":
        assign_values(
            anatomical_location=(
                "Chest"
                if diagnosis
                in ["Basal Cell Carcinoma - Left Chest", "Chest", "Left Chest"]
                else "Mediastinum"
                if diagnosis in ["Mediastinal", "Mediastinum"]
                else "Thymus"
                if diagnosis in ["Thymus"]
                else "Rib"
                if diagnosis in ["Left Rib"]
                else "Lymph Nodes"
                if diagnosis in ["Chest Wall, Node"]
                else "NR"
            ),
            anatomical_sublocation=(
                "Chest (Left)"
                if diagnosis in ["Basal Cell Carcinoma - Left Chest", "Left Chest"]
                else "Rib (Left)"
                if diagnosis in ["Left Rib"]
                else "Lymph Nodes (Chest Wall)"
                if diagnosis in ["Chest Wall, Node"]
                else "NR"
            ),
            specimen_type=(
                "N"
                if subtype in ["Normal"]
                else "M"
                if subtype in ["Metastatic Renal Cell Carcinoma"]
                else "P"
            ),
            diagnosis=(
                "Neuroblastoma"
                if subtype in ["Neuroblastoma"]
                else "Basal Cell Carcinoma"
                if subtype in ["Bcc"]
                else "Squamous Cell Carcinoma"
                if subtype in ["Scc"]
                else "Renal Cell Carcinoma"
                if subtype in ["Metastatic Renal Cell Carcinoma"]
                else "T-Lymphoblastic Lymphoma"
                if subtype in ["T Lymphoblastic L"]
                else "No Cancer"
                if subtype in ["Normal"]
                else "Melanoma"
                if subtype in ["Malignant Melanoma"]
                else "NR"
            ),
            subtype=("No Cancer" if subtype in ["Normal"] else "NR"),
            pairing_status=("No Cancer" if subtype in ["Normal"] else "NR"),
            pairing_site=(
                "No Cancer"
                if subtype in ["Normal"]
                else "Liver"
                if subtype in ["Metastatic Renal Cell Carcinoma"]
                else "NA"
            ),
            i=i,
        )


# Export to Box & Load it to format the column width
chest, chest_inventory = format_and_export(file_name="chest", tissue_source="Chest")

# In[43]: Skin (Verified)

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Skin":
        assign_values(
            anatomical_location=(
                "Skin"
                if any(
                    word in diagnosis
                    for word in [
                        "Cutaneous Melanoma",
                        "Cute (Leg)",
                        "Lesion",
                        "Excision",
                        "Skin",
                        "Cscc",
                        "Superficial",
                        "Right Nasojugal Fold",
                        "Cutaneous",
                        "Cute Pri.",
                    ]
                )
                or any(
                    word in subtype
                    for word in [
                        "Acral Lentiginous",
                        "Acral Lentiginous/Subungual Melanoma",
                        "Lesion",
                        "Excision",
                        "Skin",
                        "Cscc",
                        "Superficial",
                        "Cutaneous",
                        "Cutaneus",
                        "Lentigo Maligna",
                    ]
                )
                else (
                    "Skeletal Muscle"
                    if diagnosis in ["Left Pec Major", "Left Triceps"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Vulva"
                    if diagnosis in ["Right Vulva"]
                    or subtype in ["Vulva"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Eye"
                    if diagnosis in ["Uveal Melanoma"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Anal Canal"
                    if (subtype in ["Anal Canal"] or diagnosis in ["Anal Canal"])
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Small Intestine"
                    if subtype in ["Jejunum"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "External Jugular Vein"
                    if diagnosis in ["External Jugular Vein"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Lung"
                    if diagnosis
                    in ["Lung", "Lung (Right, Lower Lobe)", "Right Lung Lower Lobe"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Neck"
                    if diagnosis in ["Left Neck", "Right Neck"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Abdominal Wall"
                    if diagnosis in ["Abdominal Wall"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Adrenal Gland"
                    if diagnosis in ["Adrenal Gland (Left)"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Abdominal Wall"
                    if diagnosis in ["Anterior Abdominal Wall"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Intestines"
                    if diagnosis in ["Bowel"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Brain"
                    if diagnosis
                    in ["Brain", "Brain (Left Frontal Space )", "Right Frontal Lobe"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Chest"
                    if diagnosis in ["Chest", "Left Chest"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Gallbladder"
                    if diagnosis in ["Gallbladder"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Small Intestine"
                    if diagnosis
                    in [
                        "Ileum Pri.",
                        "Ileum Met. (Melanoma)",
                        "Small Bowel",
                        "Ileum  Pri.",
                    ]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Arm"
                    if diagnosis in ["Left Arm"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Femur"
                    if diagnosis in ["Left Femur"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Forearm"
                    if diagnosis in ["Left Forearm"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Iliac Crest/Wrist"
                    if diagnosis in ["Left Iliac Crest, Right Wrist"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Groin"
                    if diagnosis in ["Left Inguinal Groin"]
                    or subtype in ["Right Groin"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Leg"
                    if diagnosis in ["Left Leg"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Nasal Cavity"
                    if diagnosis in ["Left Nasal Cavity", "Right Nasal Cavity"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Vertex"
                    if diagnosis in ["Left Vertex"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Liver"
                    if diagnosis in ["Liver"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Upper Limb"
                    if subtype in ["Upper Limb"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Large Intestine"
                    if diagnosis in ["Matched Met, Large Bowel"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Medial (Nodule)"
                    if diagnosis in ["Medial Nodule"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Vagina"
                    if subtype in ["Vagina"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Groin (Nodule)"
                    if diagnosis in ["Nodule Left Groin"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Oesophagus"
                    if diagnosis in ["Oesophagus"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Rectum"
                    if diagnosis in ["Rectum"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Rectum/Anus"
                    if diagnosis in ["Rectum/Anus"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Retroperitoneum"
                    if diagnosis in ["Retroperitoneum"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Breast"
                    if diagnosis in ["Right Breast", "Right Breast, Subcutaneous"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Kidney"
                    if diagnosis in ["Right Kidney"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Nasal Septum"
                    if diagnosis in ["Right Nasal Septum"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Nasojugal Fold"
                    if diagnosis in ["Right Nasojugal Fold"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Small Bowel/Colon/Ileum"
                    if diagnosis
                    in [
                        "Small Bowel Resection And Linited Right Ehemicolectomy With End Ileostomy"
                    ]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Spleen"
                    if diagnosis in ["Spleen"]
                    else "Vagina"
                    if diagnosis in ["Vagina"]
                    else "Back"
                    if (
                        diagnosis in ["Right Upper Back", "Matched Met, Upper Back"]
                        or subtype in ["Right Back", "Back"]
                    )
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Thigh"
                    if diagnosis
                    in [
                        "Right Thigh",
                        "Left Thigh",
                        "Right Thigh, Infraspinatus",
                        "Deposit, Right Thigh",
                    ]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "Thigh/Abdomen"
                    if diagnosis in ["Thigh, Abdomen"]
                    else "Sinonasal Mucosa"
                    if diagnosis in ["Sino Nasal Mucosa"]
                    else "Abdomen"
                    if subtype in ["Left Abdomen", "Lower Abdomen"]
                    and ("Skin" not in diagnosis)
                    and (
                        not any(
                            word in subtype
                            for word in ["Cutaneous", "Cutaneus", "Superficial", "Skin"]
                        )
                    )
                    else "NR"
                )
            ),
            anatomical_sublocation=(
                "Acral"
                if subtype
                in ["Acral Lentiginous", "Acral Lentiginous/Subungual Melanoma"]
                else "Leg"
                if diagnosis in ["Cute (Leg)", "Skin, Leg"]
                else "Cheek (Left)"
                if diagnosis in ["Skin (Left Cheek)"]
                else "Jejunum"
                if subtype in ["Jejunum"]
                else "Abdomen (Left)"
                if subtype in ["Left Abdomen", "Skin, Left Abdomen"]
                or diagnosis in ["Skin, Left Abdomen"]
                else "Abdomen (Right)"
                if diagnosis in ["Skin-R Abdo"]
                else "Scalp"
                if diagnosis in ["Skin, Scalp", "Scalp, Skin"]
                else "External Jugular Vein"
                if diagnosis in ["External Jugular Vein"]
                else "Little Finger"
                if subtype in ["Little Finger"]
                else "Abdomen (Lower)"
                if subtype in ["Lower Abdomen"]
                else "Back (Right Upper)"
                if diagnosis
                in ["Right Upper Back", "Right Upper Back, Skin", "Right Upper Back"]
                else "Brain (Right Frontal Lobe)"
                if diagnosis in ["Right Frontal Lobe"]
                else "Chest (Left)"
                if diagnosis in ["Left Chest"]
                else "Forearm (Left)"
                if diagnosis in ["Left Forearm", "Skin, Left Forearm"]
                else "Iliac Crest (Left)/ Wrist (Right)"
                if diagnosis in ["Left Iliac Crest, Right Wrist"]
                else "Neck (Left)"
                if diagnosis in ["Left Neck"]
                else "Thigh (Left)"
                if diagnosis in ["Left Thigh, Skin", "Left Thigh"]
                else "Chest (Wall)"
                if diagnosis in ["Chest Wall, Skin"]
                else "Heel (Right)"
                if diagnosis in ["Right Heel, Skin"]
                else "Arm (Left)"
                if diagnosis in ["Left Arm, Skin", "Left Arm"]
                else "Axilla (Left)"
                if diagnosis in ["Left Axilla, Skin"]
                else "Arm (Left Upper)"
                if diagnosis in ["Left Upper Arm, Skin"]
                else "Back"
                if diagnosis in ["Back, Skin", "Skin, Back", "Skin-Back"]
                else "Neck (Right)"
                if diagnosis in ["Right Neck, Skin", "Right Neck"]
                else "Flank (Left)"
                if diagnosis in ["Left Flank, Skin"]
                else "Elbow (Left)"
                if diagnosis in ["Left Elbow, Skin"]
                else "Leg (Right Lower)"
                if diagnosis in ["Right Lower Leg, Skin"]
                else "Knee (Left)"
                if diagnosis in ["Left Knee, Skin"]
                else "Back (Right)"
                if diagnosis in ["Skin, Right Back", "Right Back"]
                or subtype in ["Right Back"]
                else "Ileum"
                if diagnosis in ["Ileum Met. (Melanoma)", "Ileum Pri.", "Ileum  Pri."]
                else "Thigh (Right)"
                if diagnosis in ["Right Thigh", "Right Thigh, Infraspinatus"]
                else "Thigh (Right - Infraspinatus)"
                if diagnosis in ["Right Thigh, Infraspinatus"]
                else "Back (Left Mid)"
                if diagnosis in ["Skin: Left, Mid Back"]
                else "Leg (Lower Right)"
                if diagnosis in ["Skin, Lower Right Leg"]
                else "Abdominal Wall (Anterior)"
                if diagnosis in ["Anterior Abdominal Wall"]
                else "Back (Upper)"
                if diagnosis in ["Matched Met, Upper Back"]
                else "Brain (Left Frontal Space)"
                if diagnosis in ["Brain (Left Frontal Space )"]
                else "Breast (Right - Subcutaneous)"
                if diagnosis in ["Right Breast, Subcutaneous"]
                else "Femur (Left)"
                if diagnosis in ["Left Femur"]
                else "Groin (Left Nodule)"
                if diagnosis in ["Nodule Left Groin"]
                else "Kidney (Right)"
                if diagnosis in ["Right Kidney"]
                else "Lung (Right Lower Lobe)"
                if diagnosis in ["Right Lung Lower Lobe", "Lung (Right, Lower Lobe)"]
                else "Left Pec Major"
                if diagnosis in ["Left Pec Major"]
                else "Leg (Right)"
                if diagnosis in ["Right Leg, Skin"]
                else "Thigh (Right - Deposit)"
                if diagnosis in ["Deposit, Right Thigh"]
                else "Leg (Left)"
                if diagnosis in ["Left Leg", "Skin, Left Leg"]
                else "Adrenal Gland (Left)"
                if diagnosis in ["Adrenal Gland (Left)"]
                else "Breast (Right)"
                if diagnosis in ["Right Breast"]
                else "Groin (Left Inguinal)"
                if diagnosis in ["Left Inguinal Groin"]
                else "Groin (Right)"
                if diagnosis in ["Skin, Right Groin", "Right Groin"]
                or subtype in ["Right Groin"]
                else "Nasal Cavity (Left)"
                if diagnosis in ["Left Nasal Cavity"]
                else "Nasal Cavity (Right)"
                if diagnosis in ["Right Nasal Cavity"]
                else "Abdominal Wall"
                if diagnosis in ["Skin, Abdominal Wall"]
                else "Triceps (Left)"
                if diagnosis in ["Left Triceps"]
                else "Buttock (Right)/ Leg (Right)"
                if diagnosis in ["Right Buttock And Right Leg"]
                else "Nasal Septum (Right)"
                if diagnosis in ["Right Nasal Septum"]
                else "Vulva"
                if diagnosis in ["Skin Of Vulva"]
                else "Vertex (Left)"
                if diagnosis in ["Left Vertex"]
                else "Vulva (Right)"
                if diagnosis in ["Right Vulva"]
                else "Medial (Nodule)"
                if diagnosis in ["Medial Nodule"]
                else "Breast"
                if diagnosis in ["Breast"]
                else "Glutes"
                if diagnosis in ["Gluteus"]
                else "Nasojugal Fold (Right)"
                if diagnosis in ["Right Nasojugal Fold"]
                else "Thigh/Groin"
                if diagnosis in ["Groin/Thigh"]
                else "Axilla/Chest"
                if diagnosis in ["Skin, Axilla/Chest"]
                else "NR"
            ),
            specimen_type=(
                "M"
                if (
                    diagnosis
                    in [
                        "Matched Met, Upper Back",
                        "Matched Met, Large Bowel",
                        "Ileum Met. (Melanoma)",
                        "Skin Metastasis",
                    ]
                    or "Metastatic" in subtype
                    or diagnosis in ["Lung", "Right Frontal Lobe"]
                )
                and ("Superficial Spreading" not in subtype)
                else "N"
                if diagnosis in ["Skin Nor."] or subtype in ["Normal", "Normal Skin"]
                else "N (Adjacent)"
                if subtype in ["Normal Adjacent Skin"]
                else "NR"
                if (diagnosis == "Skin" and subtype == "-")
                or (diagnosis == ["Not Merkel Cell Carcinoma"])
                else "P"
            ),
            diagnosis=(
                "Squamous Cell Carcinoma"
                if diagnosis in ["Cscc", "Cutaneous Scc"] or subtype in ["Scc"]
                else "Melanoma"
                if diagnosis
                in [
                    "Cutaneous Melanoma",
                    "Malignant Melanoma",
                    "Melanoma",
                    "Mucosal Melanoma",
                    "Uveal Melanoma",
                ]
                or subtype
                in [
                    "Malignant Melanoma",
                    "Cutaneous Melanoma",
                    "Cutaneus Melanoma",
                    "Malignant Melanoma (Nodular)",
                    "Malignant Melanoma (Superficial Spreading)",
                    "Malignt Melanoma",
                    "Metastatic Malignant Melanoma",
                    "Metastatic Malignant Melanoma (Lh18-08858)",
                    "Metastatic Malignant Melanoma (Unknown Origin)",
                    "Metastatic Melanoma (On Path Report); However Is Likely Primary Site (Nasal Mucosal Melanoma)",
                    "Metastatic Melanoma, Nos",
                    "Metastatic Mucosal Melanoma",
                    "Metastatic Nodular Malignant Melanoma (Lh17-42508)",
                    "Metastatic Poorly Differentiated Malignancy (Presumed Melanoma Given Pmhx; S100-)",
                    "Metastatic Superficial Spreading Melanoma",
                    "Metastatic Superficial Spreading Melanoma (Lh17-721)",
                    "Metastatic Superficial Spreading Melanoma (Lh17-721)",
                    "Mucosal Malignant Melanoma",
                    "Mucosal Melanoma",
                    "Nodular Malignant Melanoma",
                    "Recurrent Malignant Melanoma",
                    "Recurrent Malignant Melanoma (Nodular)",
                    "Recurrent Melanoma",
                    "Recurrent Metastatic Malignant Melanoma",
                    "Small Cell Variant, Malignant Melanoma (Mucosal)",
                    "Superficial Spreading Melanoma",
                    "Acral Lentiginous/Subungual Melanoma",
                    "Metastatic Melanoma",
                    "Melanoma",
                    "Lentigo Maligna Melanoma",
                    "Lentigo Maligna",
                ]
                else "Merkel Cell Carcinoma"
                if diagnosis in ["Merkel Cell Carcinoma"]
                else "T-Cell Lymphoma"
                if subtype
                in ["C84.0 Mycosis Fungoides"]  # subtype is Mycosis Fungoides (MF)
                else "Peripheral B-Cell Lymphoma"
                if subtype in ["Dlbcl Various Site"]
                else "Peripheral B-Cell Lymphoma"
                if subtype in ["Follicular Lymphoma"]
                else "Basal Cell Carcinoma"
                if subtype
                in [
                    "Infiltrative Basal Cell Carcinoma",
                    "Infiltrative Basal Cell Carcinoma (Insufficient Tissue) Scrap",
                    "Nodular And Infiltrating Basal Cell Carcinoma",
                    "Nodular Basal Cell Carcinoma",
                    "Nodular Basal Cell Carcinoma Get Block A5 On Lh17-20390",
                ]
                else "No Cancer"
                if diagnosis in ["Skin Nor."]
                or subtype in ["Normal", "Normal Skin", "Normal Adjacent Skin"]
                else "T-Cell Lymphoma"
                if subtype in ["Tcl"]
                else "NR"
            ),
            subtype=(
                "Cutaneous Squamous Cell Carcinoma"
                if diagnosis in ["Cscc", "Cutaneous Scc"]
                else "Cutaneous Melanoma"
                if diagnosis in ["Cutaneous Melanoma"]
                or subtype in ["Cutaneous Melanoma", "Cutaneus Melanoma"]
                else "Mucosal Melanoma"
                if diagnosis in ["Mucosal Melanoma"]
                or subtype
                in [
                    "Metastatic Mucosal Melanoma",
                    "Mucosal Malignant Melanoma",
                    "Mucosal Melanoma",
                ]
                else "Uveal Melanoma"
                if diagnosis in ["Uveal Melanoma"]
                else "Nodular Melanoma"
                if subtype
                in [
                    "Malignant Melanoma (Nodular)",
                    "Metastatic Nodular Malignant Melanoma (Lh17-42508)",
                    "Nodular Malignant Melanoma",
                ]
                else "Poorly Differentiated Melanoma"
                if subtype
                in [
                    "Metastatic Poorly Differentiated Malignancy (Presumed Melanoma Given Pmhx; S100-)"
                ]
                else "Superficial Spreading Melanoma"
                if subtype
                in [
                    "Malignant Melanoma (Superficial Spreading)",
                    "Metastatic Superficial Spreading Melanoma",
                    "Superficial Spreading Melanoma",
                    "Metastatic Superficial Spreading Melanoma (Lh17-721)",
                ]
                else "Melanoma (Recurrent)"
                if subtype
                in [
                    "Recurrent Malignant Melanoma",
                    "Recurrent Melanoma",
                    "Recurrent Metastatic Malignant Melanoma",
                ]
                else "Nodular Melanoma (Recurrent)"
                if subtype in ["Recurrent Malignant Melanoma (Nodular)"]
                else "Mucosal Melanoma Small Cell Variant"
                if subtype in ["Small Cell Variant, Malignant Melanoma (Mucosal)"]
                else "Mycosis Fungoides"
                if subtype in ["C84.0 Mycosis Fungoides"]
                else "Diffuse Large B-Cell Lymphoma"
                if subtype in ["Dlbcl Various Site"]
                else "Follicular Lymphoma"
                if subtype in ["Follicular Lymphoma"]
                else "Infiltrative Basal Cell Carcinoma"
                if subtype
                in [
                    "Infiltrative Basal Cell Carcinoma",
                    "Infiltrative Basal Cell Carcinoma (Insufficient Tissue) Scrap",
                ]
                else "Infiltrative Basal Cell Carcinoma (Insufficient Tissue - Scrap)"
                if subtype
                in ["Infiltrative Basal Cell Carcinoma (Insufficient Tissue) Scrap"]
                else "Nodular And Infiltrative Basal Cell Carcinoma"
                if subtype in ["Nodular And Infiltrating Basal Cell Carcinoma"]
                else "Nodular Basal Cell Carcinoma"
                if subtype
                in [
                    "Nodular Basal Cell Carcinoma Get Block A5 On Lh17-20390",
                    "Nodular Basal Cell Carcinoma",
                ]
                else "Lentigo Maligna"
                if subtype in ["Lentigo Maligna", "Lentigo Maligna Melanoma"]
                else "Acral Lentiginous"
                if subtype in ["Acral Lentiginous"]
                else "Acral And Subungual Melanoma"
                if subtype in ["Acral Lentiginous/Subungual Melanoma"]
                else "Poorly Differentiated Melanoma With Rhabdoid Morphology"
                if subtype in ["Rhabdoid Morphology, Poorly Differentiated"]
                else "No Cancer"
                if diagnosis in ["Skin Nor."]
                or subtype in ["Normal", "Normal Skin", "Normal Adjacent Skin"]
                else "NR"
            ),
            pairing_status=(
                "P"
                if "Matched Met" in diagnosis
                else "No Cancer"
                if diagnosis in ["Skin Nor."]
                or subtype in ["Normal", "Normal Skin", "Normal Adjacent Skin"]
                else "NR"
            ),
            pairing_site=(
                "No Cancer"
                if diagnosis in ["Skin Nor."]
                or subtype in ["Normal", "Normal Skin", "Normal Adjacent Skin"]
                else "NA"
            ),
            i=i,
        )


# Export to Box & Load it to format the column width
skin, skin_inventory = format_and_export(file_name="skin", tissue_source="Skin")


# In[44]: Digestive Tract (Verified)

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Digestive Tract":
        assign_values(
            anatomical_location=(
                "Stomach"
                if subtype == "Stomach"
                else "Colon"
                if subtype == "Colon"
                else "Small Intestine"
                if subtype == "Duodenum"
                else "Gastrointestinal Tract"
            ),
            anatomical_sublocation=("Duodenum" if subtype == "Duodenum" else "NR"),
            specimen_type="P",
            diagnosis="Gastrointestinal Stromal Tumour",
            subtype="NR",
            pairing_status="NR",
            pairing_site="NA",
            i=i,
        )


# Export to Box & Load it to format the column width
digestive_tract, digestive_tract_inventory = format_and_export(
    file_name="digestive_tract", tissue_source="Digestive Tract"
)


# In[45]:Abscess (Verified)
for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Abscess":
        assign_values(
            anatomical_location="Abscess",
            anatomical_sublocation="NR",
            specimen_type="NA",
            diagnosis="NR",
            subtype="NR",
            i=i,
        )


# Export to Box & Load it to format the column width
abscess, abscess_inventory = format_and_export(
    file_name="abscess", tissue_source="Abscess"
)


# In[46]: Peritoneum (Verified)

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Peritoneum":
        assign_values(
            anatomical_location="Peritoneum",
            anatomical_sublocation="NR",
            specimen_type=("M" if subtype == "Metastasis" else "P"),
            diagnosis="Carcinoma",
            subtype=(
                "High Grade Serous Carcinoma"
                if subtype == "Hg Serous"
                else "Low Grade Serous Carcinoma"
                if subtype == "Lg Serous"
                else "NR"
            ),
            i=i,
        )


# Export to Box & Load it to format the column width
peritoneum, peritoneum_inventory = format_and_export(
    file_name="peritoneum", tissue_source="Peritoneum"
)

# In[47]: Lymphatic/Blood Vessels (Verified)
for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Lymphatic/Blood Vessels":
        assign_values(
            anatomical_location="Lymphatic/Blood Vessels",
            anatomical_sublocation="Lymphatic/Blood Vessels (Wall Lining)",
            specimen_type="M" if subtype == "Crc_Matched Met" else "P",
            diagnosis="Soft Tissue Sarcoma",
            subtype="Angiosarcoma",
            pairing_status="P" if subtype == "Crc_Matched Met" else "NR",
            pairing_site="Colon/Rectum" if subtype == "Crc_Matched Met" else "NA",
            i=i,
        )


# Export to Box & Load it to format the column width
lymphatic_blood_vessels, lymphatic_blood_vessels_inventory = format_and_export(
    file_name="lymphatic_blood_vessels", tissue_source="Lymphatic/Blood Vessels"
)


# In[48]: Oesophagus (Verified)
for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Oesophagus":
        assign_values(
            anatomical_location=(
                "Lymph Nodes"
                if subtype in ["Normal (Non-Malignant) Lymph Noder"]
                else "Oesophagus"
            ),
            anatomical_sublocation="Lymph Nodes (Oesophagus)"
            if subtype in ["Normal (Non-Malignant) Lymph Noder"]
            else "NR",
            specimen_type=(
                "M"
                if diagnosis in ["Esophagus Met."] or subtype in ["Matched Met"]
                else "N"
                if diagnosis in ["Oesophagus Nor."]
                or subtype in ["Normal (Non-Malignant) Lymph Noder", "Normal"]
                else "NR"
                if (subtype in ["Not Scc"])
                or (diagnosis in ["Esophagus"] and subtype in ["-"])
                else "P"
            ),
            diagnosis=(
                "Squamous Cell Carcinoma"
                if subtype
                in [
                    "Squamous Cell Carcinoma",
                    "Basaloid Squamous Cell Carcinoma",
                    "Matched Met",
                    "Scc",
                    "Squamous",
                ]
                else "Adenosquamous Carcinoma"
                if subtype in ["Adenosquamous"]
                else "Adenocarcinoma"
                if (diagnosis == "Esophageal Adc" and subtype == "Primary")
                or subtype in ["Adc", "Adenocarcinoma", "Colloid Adenocarcinoma"]
                else "Mixed Adenoneuroendocrine Carcinoma"
                if subtype
                in [
                    "Mixed Adenoneuroendocrine Carcinoma (Manec) Arising On A Background Of Barrett'S Oesophagus"
                ]
                else "No Cancer"
                if diagnosis in ["Oesophagus Nor."]
                or subtype in ["Normal (Non-Malignant) Lymph Noder", "Normal"]
                else "NR"
            ),
            subtype=(
                "Basaloid Squamous Cell Carcinoma"
                if subtype in ["Basaloid Squamous Cell Carcinoma"]
                else "Colloid Adenocarcinoma"
                if subtype in ["Colloid Adenocarcinoma"]
                else "Adenocarcinoma And Neuroendocrine Carcinoma"
                if subtype
                in [
                    "Mixed Adenoneuroendocrine Carcinoma (Manec) Arising On A Background Of Barrett'S Oesophagus"
                ]
                else "No Cancer"
                if diagnosis in ["Oesophagus Nor."]
                or subtype in ["Normal (Non-Malignant) Lymph Noder", "Normal"]
                else "NR"
            ),
            pairing_status=(
                "P"
                if subtype in ["Matched Met"]
                or diagnosis in ["Esophagus Pri. And Met."]
                else "No Cancer"
                if diagnosis in ["Oesophagus Nor."]
                or subtype in ["Normal (Non-Malignant) Lymph Noder", "Normal"]
                else "NR"
            ),
            pairing_site=(
                "No Cancer"
                if diagnosis in ["Oesophagus Nor."]
                or subtype in ["Normal (Non-Malignant) Lymph Noder", "Normal"]
                else "NA"
            ),
            i=i,
        )


# Export to Box & Load it to format the column width
oesophagus, oesophagus_inventory = format_and_export(
    file_name="oesophagus", tissue_source="Oesophagus"
)


# In[49]: Connective Tissue (Verified)

# Create combinations dictionary for Connective Tissue
create_combinations_dict("Connective Tissue")

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Connective Tissue":
        assign_values(
            anatomical_location="Connective Tissue",
            anatomical_sublocation="NR",
            specimen_type="P",
            diagnosis="Soft Tissue Sarcoma",
            subtype="Fibrosarcoma"
            if subtype == "Fibrosarcoma"
            else "Undifferentiated Sarcoma",
            i=i,
        )


# Export to Box & Load it to format the column width
connective_tissue, connective_tissue_inventory = format_and_export(
    file_name="fibrous_connective_tissue", tissue_source="Fibrous Connective Tissue"
)


# In[50]: Adrenal Gland (Verified)

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Adrenal Gland":
        assign_values(
            anatomical_location=(
                "Adrenal Gland/Kidney"
                if diagnosis in ["Right Kidney + Adrenal"]
                else "Adrenal Gland"
            ),
            anatomical_sublocation=(
                "Adrenal Gland (Left)"
                if diagnosis in ["Left Adrenal Gland"]
                else "Adrenal Gland/ Kidney (Right)"
                if diagnosis in ["Right Kidney + Adrenal"]
                else "NR"
            ),
            specimen_type=(
                "N"
                if subtype in ["Normal", "Normal (Non Malignant)"]
                else "M"
                if any(keyword in subtype for keyword in ["Metastatic", "Matched Met"])
                or diagnosis in ["Adrenal Gland Met. (Breast)"]
                else "P"
            ),
            diagnosis=(
                "Adenocarcinoma"
                if "Adenocarcinoma" in subtype
                else "Renal Cell Carcinoma"
                if "Renal Cell Carcinoma" in subtype
                else "Neuroblastoma"
                if subtype in ["Neuroblastoma"]
                else "Invasive Carcinoma"
                if subtype in ["Idc"]
                else "No Cancer"
                if subtype in ["Normal", "Normal (Non Malignant)"]
                else "NR"
            ),
            subtype=(
                "Clear Cell Carcinoma"
                if subtype
                in [
                    "Metastatic Conventional (Clear Cell) Renal Cell Carcinoma",
                    "Clear Cell Renal Cell Carcinoma",
                    "Metastatic Clear Cell Renal Cell Carcinoma",
                ]
                else "Invasive Ductal Carcinoma"
                if subtype in ["Idc"]
                else "No Cancer"
                if subtype in ["Normal", "Normal (Non Malignant)"]
                else "NR"
            ),
            pairing_status=(
                "P"
                if subtype in ["Matched Met"]
                else "No Cancer"
                if subtype in ["Normal", "Normal (Non Malignant)"]
                else "NR"
            ),
            pairing_site=(
                "Lung"
                if subtype in ["Metastatic Lung Adenocarcinoma"]
                else "Kidney"
                if subtype
                in [
                    "Metastatic Conventional (Clear Cell) Renal Cell Carcinoma",
                    "Metastatic Clear Cell Renal Cell Carcinoma",
                    "Clear Cell Renal Cell Carcinoma",
                ]
                else "Breast"
                if diagnosis in ["Adrenal Gland Met. (Breast)"]
                else "No Cancer"
                if subtype in ["Normal", "Normal (Non Malignant)"]
                else "NA"
            ),
            i=i,
        )


# Export to Box & Load it to format the column width
adrenal_gland, adrenal_gland_inventory = format_and_export(
    file_name="adrenal_gland", tissue_source="Adrenal Gland"
)


# In[51]: Pleura (Verified)
for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Pleura":
        assign_values(
            anatomical_location=(
                "Pleura/Chest"
                if diagnosis in ["Pleura/Chest Wall"]
                else "Pleura"
                if diagnosis
                in [
                    "Pleura Pri.",
                    "Parietal Pleura  Pri.",
                    "Parietal- And Visceral Pleura  Pri.",
                    "Parietal- And Visceral Pleura \xa0Pri.",
                    "Visceral Pleura  Pri.",
                ]
                else "NR"
            ),
            anatomical_sublocation=(
                "Pleura/Chest (Wall)"
                if diagnosis in ["Pleura/Chest Wall"]
                else "Pleura (Parietal)"
                if diagnosis in ["Parietal Pleura  Pri."]
                else "Pleura (Parietal And Visceral)"
                if diagnosis
                in [
                    "Parietal- And Visceral Pleura  Pri.",
                    "Parietal- And Visceral Pleura \xa0Pri.",
                ]
                else "Pleura (Visceral)"
                if diagnosis in ["Visceral Pleura  Pri."]
                else "NR"
            ),
            specimen_type=("M" if "Metastatic" in subtype else "P"),
            diagnosis=(
                "Renal Cell Carcinoma"
                if subtype in ["Met Blocks/Metastatic Renal Cell Carcinoma (Multiple)"]
                else "Mesothelioma"
            ),
            subtype=(
                "Biphasic Mesothelioma"
                if subtype in ["Biphasic Malignant Mesothelioma"]
                else "Epithelioid Mesothelioma"
                if subtype in ["Epithelioid Malignant Mesothelioma"]
                else "NR"
            ),
            pairing_status="NR",
            pairing_site="Kidney"
            if subtype in ["Met Blocks/Metastatic Renal Cell Carcinoma (Multiple)"]
            else "NA",
            i=i,
        )


# Export to Box & Load it to format the column width
pleura, pleura_inventory = format_and_export(file_name="pleura", tissue_source="Pleura")

# In[52]: Bile Duct (Verified)

# create combinations dictionary for Bile Duct
create_combinations_dict("Bile Duct")

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Bile Duct":
        assign_values(
            anatomical_location="Bile Duct",
            anatomical_sublocation=(
                "Bile Duct (Hilar)"
                if any(
                    keyword in subtype for keyword in ["Perihilar", "Perhilar", "Hilar"]
                )
                else "Bile Duct (Intrahepatic)"
                if any(keyword in subtype for keyword in ["Intrahepatic"])
                else "NR"
            ),
            specimen_type=(
                "M"
                if subtype in ["Matched Mets"]
                else "NR"
                if diagnosis == "Bile Duct" and subtype == "-"
                else "P"
            ),
            diagnosis=(
                "Cholangiocarcinoma"
                if "Cholangiocarcinoma" in subtype or "Cholangiocarcinoma" in diagnosis
                else "Adenocarcinoma"
                if subtype == "Adc"
                else "Intraductal Tubopapillary Neoplasm"
                if subtype == "Intraductal Tubopapillary Neoplasm"
                else "NR"
            ),
            subtype=(
                "Ductal Cholangiocarcinoma"
                if "Ductal" in diagnosis
                else "Cholangiocarcinoma With Squamous Cell Differentiation"
                if subtype
                in ["Squamous Cell Carcinoma Intrahepatic Cholangiocarcinoma"]
                else "NR"
            ),
            pairing_status=("P" if subtype in ["Matched Mets"] else "NR"),
            pairing_site="NA",
            i=i,
        )


# Export to Box & Load it to format the column width
bile_duct, bile_duct_inventory = format_and_export(
    file_name="bile_duct", tissue_source="Bile Duct"
)

# In[53]: Connective Tissue (Verified)
for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Connective Tissue":
        assign_values(
            anatomical_location="Connective Tissue",
            anatomical_sublocation="NR",
            specimen_type="P",
            diagnosis="Sarcoma",
            subtype="Undifferentiated Sarcoma",
            i=i,
        )


# Export to Box & Load it to format the column width
connective_tissue, connective_tissue_inventory = format_and_export(
    file_name="connective_tissue", tissue_source="Connective Tissue"
)


# In[54]: Back (Verified)
for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Back":
        assign_values(
            anatomical_location=(
                "Skin"
                if diagnosis in ["Back, Skin/Node"]
                else "Bone"
                if any(keyword in diagnosis for keyword in ["Vertebrae", "Spine"])
                else "Skin"
                if diagnosis in ["Left Upper Back"]
                else "NR"
            ),
            anatomical_sublocation=(
                "Spine (T7)"
                if diagnosis in ["T7 Vertebrae"]
                else "Spine (L4)"
                if diagnosis in ["L4 Vertebrae"]
                else "Spine (T12)"
                if diagnosis in ["Spine (T12)"]
                else "Back (Left Upper)"
                if diagnosis in ["Left Upper Back"]
                else "NR"
            ),
            specimen_type=("M" if "Metastatic" in subtype else "P"),
            diagnosis=(
                "Melanoma"
                if subtype in ["Superficial Spreading Melanoma"]
                else "Thymus Cancer"
                if subtype in ["Metastatic Thymus Cancer"]
                else "Adenocarcinoma"
                if "Adenocarcinoma" in subtype
                else "Renal Cell Carcinoma"
                if subtype
                in [
                    "Metastatic Carcinoma Consistent With Kidney As Primary Site (Clear Cell Renal Cell Carcinoma)",
                    "Metastatic Renal Cell Carcinoma",
                    "Metastatic Renal Clear Cell Carcinoma",
                    "Metastatic Clear Cell Carcinoma (In Keeping With Conventional Rcc)",
                ]
                else "Squamous Cell Carcinoma"
                if subtype in ["Scc"]
                else "NR"
            ),
            subtype=(
                "Superficial Spreading Melanoma"
                if subtype in ["Superficial Spreading Melanoma"]
                else "Renal Cell Carcinoma (NOS)"
                if subtype
                in [
                    "Metastatic Renal Cell Carcinoma",
                    "Metastatic Clear Cell Carcinoma (In Keeping With Conventional Rcc)",
                ]
                else "Clear Cell Carcinoma"
                if subtype
                in [
                    "Metastatic Carcinoma Consistent With Kidney As Primary Site (Clear Cell Renal Cell Carcinoma)",
                    "Metastatic Renal Clear Cell Carcinoma",
                ]
                else "NR"
            ),
            pairing_status="NR",
            pairing_site=(
                "Lung"
                if "Lung" in subtype
                else "Thymus"
                if "Thymus" in subtype
                else "Kidney"
                if any(keyword in subtype for keyword in ["Kidney", "Rcc", "Renal"])
                else "NA"
            ),
            i=i,
        )


# Export to Box & Load it to format the column width
back, back_inventory = format_and_export(file_name="back", tissue_source="Back")

# In[55]: Cortex (Verified)
for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Cortex":
        assign_values(
            anatomical_location=(
                "Kidney"
                if "Kidney" in diagnosis
                else "Brain"
                if diagnosis
                in [
                    "Cerebellum, Cortex",
                    "Normal Cerebral Cortex",
                    "Cortex Nor. (Cerebrum)",
                    "Cerebellum, Cortex Nor.",
                ]
                else "Cortex"
            ),
            anatomical_sublocation=(
                "Kidney (Medulla And Cortex)"
                if diagnosis in ["Kidney Medulla & Cortex"]
                else "Kidney (Cortex)"
                if diagnosis in ["Kidney Cortex", "Kidney, Cortex"]
                else "Cerebellum (Cortex)"
                if diagnosis in ["Cerebellum, Cortex", "Cerebellum, Cortex Nor."]
                else "Cerebrum (Cortex)"
                if diagnosis in ["Cortex Nor. (Cerebrum)", "Normal Cerebral Cortex"]
                else "NR"
            ),
            specimen_type="N",
            diagnosis="No Cancer",
            subtype="No Cancer",
            pairing_status="No Cancer",
            pairing_site="No Cancer",
            i=i,
        )


# Export to Box & Load it to format the column width
cortex, cortex_inventory = format_and_export(file_name="cortex", tissue_source="Cortex")


# In[56]: Mesothelium (Verified)
for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Mesothelium":
        assign_values(
            anatomical_location=(
                "Tunica Vaginalis Testis"
                if diagnosis
                in ["Testicular Vaginal Biphasic Mesothelioma Lung Metastasis"]
                else "Peritoneum"
                if subtype in ["Peritoneal"]
                else "Pleura"
                if subtype in ["Pleural"]
                else "NR"
            ),
            anatomical_sublocation="NR",
            specimen_type="M" if "Lung Metastasis" in diagnosis else "P",
            diagnosis="Mesothelioma",
            subtype=(
                "Epithelioid Mesothelioma"
                if subtype in ["Epithelioid", "Epithelioid Mesothelioma"]
                or diagnosis in ["Epitheloid Mesothelioma"]
                else "Multicystic Mesothelioma"
                if subtype
                in [
                    "Multicystic Mesothelioma",
                ]
                else "Sarcomatoid Mesothelioma"
                if subtype in ["Sarcomatoid"]
                or diagnosis in ["Sarcomatoid Mesothelioma"]
                else "Biphasic Mesothelioma"
                if subtype in ["Biphasic"]
                or diagnosis
                in [
                    "Testicular Vaginal Biphasic Mesothelioma Lung Metastasis",
                    "Biphasic Mesothelioma",
                ]
                else "Epithelioid Neoplasia Mesothelioma"
                if subtype in ["Epithelioid Neoplasia Mesothelioma"]
                else "Desmoplastic Mesothelioma"
                if diagnosis in ["Desmoplasic Mesothelioma"]
                else "NR"
            ),
            pairing_status="NR",
            pairing_site=("Lung" if "Lung" in diagnosis else "NR"),
            i=i,
        )


# Export to Box & Load it to format the column width
mesothelium, mesothelium_inventory = format_and_export(
    file_name="mesothelium", tissue_source="Mesothelium"
)

# In[57]: Arms (Verified)

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Arms":
        assign_values(
            anatomical_location=(
                "Lymph Nodes"
                if any(keyword in diagnosis for keyword in ["Ln", "Node", "Nodes"])
                else "Lymph Nodes/Skin"
                if diagnosis in ["Left Axillary Nodes, Left Arm Skin"]
                else "Skin"
                if any(
                    keyword in diagnosis
                    for keyword in ["Scc", "Cscc"] or "Scc" in subtype
                )
                else "Arm"
                if diagnosis in ["Left Arm Pri. (Sarcoma)", "Left Humerus"]
                else "NR"
            ),
            anatomical_sublocation=(
                "Lymph Nodes (Axillary/Arm)"
                if diagnosis in ["Node, Axilla/Arm"]
                else "Lymph Nodes (Axillary)"
                if diagnosis
                in [
                    "Axilla Node",
                    "Axillary Nodes",
                    "Axillary Lymph Node",
                    "Ln, Axilla",
                ]
                else "Lymph Nodes (Sentinel Right Axillary)"
                if diagnosis in ["Sentinel Node Right Axilla"]
                else "Lymph Nodes (Right Axillary)"
                if diagnosis
                in [
                    "Right Axillary Lymph Node",
                    "Right Axillary Node",
                    "Nodes (Right Axilla)",
                ]
                else "Lymph Nodes (Left Axillary)"
                if diagnosis
                in [
                    "Left Axillary Nodes",
                    "Left Axillary Node",
                    "Matched Met, Left Axilla Node",
                    "Left Axillary Lymph Node",
                    "Nodes (Left Axilla)",
                ]
                else "Lymph Nodes (Left Axillary) And Skin (Left Arm)"
                if diagnosis in ["Left Axillary Nodes, Left Arm Skin"]
                else "Lymph Nodes (Axillary/Intraclavical)"
                if diagnosis in ["Axillary/Intraclavical Nodes"]
                else "Forearm (Left)"
                if diagnosis in ["Left Forearm"]
                else "Hand (Left Dorsum)"
                if subtype in ["Dorsum Lefrt Hand"]
                else "Hand (Left)"
                if diagnosis in ["Left Hand"]
                else "Chest (1-Sternum)/Hand (3-Hand)"
                if subtype in ["1-Sternum 3-Hand"]
                else "Forearm (Right)"
                if diagnosis in ["Right Forearm"]
                else "Arm (Left)"
                if diagnosis in ["Left Arm Pri. (Sarcoma)"]
                else "Arm (Left Humerus)"
                if diagnosis in ["Left Humerus"]
                else "Hand (Left Thumb)"
                if diagnosis in ["Left Thumb"]
                else "Wrist"
                if diagnosis in ["Wrist"]
                else "NR"
            ),
            specimen_type=(
                "M"
                if any(
                    keyword in subtype
                    for keyword in ["Metastatic", "Mets", "Matched Met"]
                )
                or (
                    any(keyword in diagnosis for keyword in ["Node"])
                    and "Melanoma" in subtype
                )
                else "N (Adjacent)"
                if any(keyword in subtype for keyword in ["Normal Adjacent"])
                else "P"
            ),
            diagnosis=(
                "Melanoma"
                if any(keyword in subtype for keyword in ["Melanoma"])
                else "Squamous Cell Carcinoma"
                if diagnosis in ["Scc", "Cscc"] or subtype in ["Scc"]
                else "Peripheral T-Cell Lymphoma"
                if subtype in ["Ptcl, Nos", "Angioimmunoblastic T-Cell L"]
                else "Sarcoma"
                if "Sarcoma" in diagnosis
                else "Osteosarcoma"
                if subtype in ["Conventional Osteosarcoma"]
                else "No Cancer"
                if any(keyword in subtype for keyword in ["Normal Adjacent"])
                else "NR"
            ),
            subtype=(
                "Superficial Spreading Melanoma"
                if subtype in ["Superficial Spreading Melanoma"]
                else "Nodular Melanoma"
                if subtype in ["Metastatic Nodular Melanoma"]
                else "Angioimmunoblastic T-Cell Lymphoma"
                if subtype in ["Angioimmunoblastic T-Cell L"]
                else "Undifferentiated Sarcoma"
                if subtype in ["Undifferentiated Sarcoma Of Left Arm"]
                else "Conventional Osteosarcoma"
                if subtype in ["Conventional Osteosarcoma"]
                else "Cutaneous Squamous Cell Carcinoma"
                if diagnosis in ["Cscc"]
                else "No Cancer"
                if any(keyword in subtype for keyword in ["Normal Adjacent"])
                else "NR"
            ),
            pairing_status=(
                "P"
                if "Matched Met" in diagnosis or "Matched Met" in subtype
                else "No Cancer"
                if any(keyword in subtype for keyword in ["Normal Adjacent"])
                else "NR"
            ),
            pairing_site=(
                "No Cancer"
                if any(keyword in subtype for keyword in ["Normal Adjacent"])
                else "NA"
            ),
            i=i,
        )


# Export to Box & Load it to format the column width
arms, arms_inventory = format_and_export(file_name="arms", tissue_source="Arms")

# In[58]: Smooth Muscle Tissue (Verified)

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Smooth Muscle Tissue":
        assign_values(
            anatomical_location=(
                "Leg"
                if "Leg" in diagnosis
                else "Buttock"
                if "Buttock" in diagnosis
                else "Pelvis"
                if "Pelvis" in diagnosis
                else "Shoulder"
                if "Shoulder" in diagnosis
                else "NR"
            ),
            anatomical_sublocation=(
                "Leg (Left - Soft Tissue)"
                if "Leg" in diagnosis
                else "Buttock (Right - Soft Tissue)"
                if "Buttock" in diagnosis
                else "Pelvis (Soft Tissue)"
                if "Pelvis" in diagnosis
                else "Shoulder (Left - Soft Tissue)"
                if "Shoulder" in diagnosis
                else "NR"
            ),
            specimen_type="P",
            diagnosis="Soft Tissue Sarcoma",
            subtype=(
                "Leiomyosarcoma"
                if subtype in ["Leiomyosarcoma", "-"]
                else "High Grade Leiomyosarcoma"
                if subtype in ["High Grade Leiomyosarcoma"]
                else "NR"
            ),
            pairing_status="NR",
            pairing_site="NA",
            i=i,
        )


# Export to Box & Load it to format the column width
smooth_muscle_tissue, smooth_muscle_tissue_inventory = format_and_export(
    file_name="smooth_muscle_tissue", tissue_source="Smooth Muscle Tissue"
)


# In[59]: Ampulla (Verified)

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Ampulla":
        assign_values(
            anatomical_location="Duodenuum",
            anatomical_sublocation=(
                "Ampulla (Surrounding Area)" if "Peri" in diagnosis else "Ampulla"
            ),
            specimen_type="P",
            diagnosis="Adenocarcinoma",
            subtype=(
                "Well Differentiated Adenocarcinoma"
                if diagnosis in ["Well Differentiated Ampullary Adenocarcinoma"]
                else "Intestinal Adenocarcinoma"
                if diagnosis in ["Intestinal Type Carcinoma Of The Ampulla"]
                else "Moderately Differentiated Adenocarcinoma"
                if diagnosis
                in ["Moderately Differentiated Peri Ampullary Adenocarcinoma"]
                else "NR"
            ),
            pairing_status="NR",
            pairing_site="NA",
            i=i,
        )


# Export to Box & Load it to format the column width
ampulla, ampulla_inventory = format_and_export(
    file_name="ampulla", tissue_source="Ampulla"
)

# In[60]: Eye (Verified)

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Eye":
        assign_values(
            anatomical_location="Eye",
            anatomical_sublocation="NR",
            specimen_type="P",
            diagnosis="Choroidal Melanoma",
            subtype="NR",
            i=i,
        )


# Export to Box & Load it to format the column width
eye, eye_inventory = format_and_export(file_name="eye", tissue_source="Eye")

# In[61]: Shoulder (Verified)
for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Shoulder":
        assign_values(
            anatomical_location=(
                "Soft Tissue"
                if "Subscapular Soft Tissue" in diagnosis
                else "Shoulder/Lymph Nodes"
                if diagnosis in ["Groin, Nodes, Shoulder"]
                else "NR"
            ),
            anatomical_sublocation=(
                "Subscapular Soft Tissue"
                if "Subscapular Soft Tissue" in diagnosis
                else "Shoulder/Lymph Nodes (Groin)"
                if diagnosis in ["Groin, Nodes, Shoulder"]
                else "NR"
            ),
            specimen_type="M" if "Met Blocks" in subtype else "P",
            diagnosis="Melanoma" if "Melanoma" in subtype else "Mesothelioma",
            subtype="Superficial Spreading Melanoma" if "Melanoma" in subtype else "NR",
            pairing_status="NR",
            pairing_site="NA",
            i=i,
        )


# Export to Box & Load it to format the column width
shoulder, shoulder_inventory = format_and_export(
    file_name="shoulder", tissue_source="Shoulder"
)

# In[62]: Retroperitoneum (Verified)

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Retroperitoneum":
        assign_values(
            anatomical_location="Retroperitoneum",
            anatomical_sublocation="NR",
            specimen_type="P",
            diagnosis="Neuroblastoma",
            subtype="NR",
            i=i,
        )


# Export to Box & Load it to format the column width
retroperitoneum, retroperitoneum_inventory = format_and_export(
    file_name="retroperitoneum", tissue_source="Retroperitoneum"
)

# In[63]: Plasma (Verified)

# Create combinations dictionary for Plasma
create_combinations_dict("Plasma")

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Plasma":
        assign_values(
            anatomical_location="Blood",
            anatomical_sublocation="Plasma",
            specimen_type="NR",
            diagnosis=(
                "Ovarian Cancer"
                if subtype in ["Ovarian Cancer"]
                else "Prostate Cancer"
                if subtype in ["Prostate Cancer"]
                else "Colorectal Cancer"
                if subtype in ["Colorectal Cancer"]
                else "NR"
            ),
            subtype="NR",
            pairing_status="NR",
            pairing_site="NA",
            i=i,
        )


# Export to Box & Load it to format the column width
plasma, plasma_inventory = format_and_export(file_name="plasma", tissue_source="Plasma")

# In[64]: Muscle (Verified)

# Create combinations dictionary for Muscle
create_combinations_dict("Muscle")


for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Muscle":
        assign_values(
            anatomical_location="Muscle",
            anatomical_sublocation=(
                "Muscle (Skeletal)"
                if "Skeletal" in diagnosis
                else "Muscle (Lower Back)"
                if "Lower Back" in diagnosis
                else "NR"
            ),
            specimen_type=(
                "N"
                if "Normal" in diagnosis or "Normal" in subtype
                else "M"
                if subtype in ["Metastatic Renal Cell Carcinoma"]
                else "P"
            ),
            diagnosis=(
                "No Cancer"
                if "Normal" in diagnosis or "Normal" in subtype
                else "Renal Cell Carcinoma"
                if subtype in ["Metastatic Renal Cell Carcinoma"]
                else "NR"
            ),
            subtype=(
                "No Cancer" if "Normal" in diagnosis or "Normal" in subtype else "NR"
            ),
            pairing_status=(
                "No Cancer" if "Normal" in diagnosis or "Normal" in subtype else "NR"
            ),
            pairing_site=(
                "No Cancer" if "Normal" in diagnosis or "Normal" in subtype else "NA"
            ),
            i=i,
        )


# Export to Box & Load it to format the column width
muscle, muscle_inventory = format_and_export(file_name="muscle", tissue_source="Muscle")

# In[65]: Multiple Sites (Pending Final Verification)

# Create combinations dictionary for Multiple Sites
create_combinations_dict("Multiple Sites")

for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "Multiple Sites":
        assign_values(
            anatomical_location="Adrenal, Thymus And Bladder",
            anatomical_sublocation="NR",
            specimen_type="N",
            diagnosis="No Cancer",
            subtype="No Cancer",
            pairing_status="No Cancer",
            pairing_site="No Cancer",
            i=i,
        )


# Export to Box & Load it to format the column width
multiple, multiple_inventory = format_and_export(
    file_name="multiple_sites", tissue_source="Multiple Sites"
)


# In[66]: Unknown Tissue Source (Pending Final Verification)

# Create combinations dictionary for Unknown Tissue Source
create_combinations_dict("NR")


for i, (tissue_src, diagnosis, subtype) in enumerate(
    zip(
        inventory["Tissue Source"],
        inventory["Specimen_Histological_Diagnosis"],
        inventory["Sub_Type"],
    )
):
    if tissue_src == "NR":
        assign_values(
            anatomical_location="Skeletal Muscle" if "Skeletal" in diagnosis else "NR",
            anatomical_sublocation="NR",
            specimen_type=(
                "N"
                if subtype in ["No Cancer", "Normal"] or "Normal" in diagnosis
                else "P"
                if subtype in ["Basaloid Squamous Cell Carcinoma"]
                else "NR"
            ),
            diagnosis=(
                "Squamous Cell Carcinoma"
                if subtype in ["Basaloid Squamous Cell Carcinoma"]
                else "No Cancer"
                if subtype in ["No Cancer", "Normal"] or "Normal" in diagnosis
                else "Liposarcoma"
                if diagnosis in ["Liposarcoma"] or subtype in ["Myxoid Liposarcoma"]
                else "NR"
            ),
            subtype=(
                "Basaloid Squamous Cell Carcinoma"
                if subtype in ["Basaloid Squamous Cell Carcinoma"]
                else "No Cancer"
                if subtype in ["No Cancer", "Normal"] or "Normal" in diagnosis
                else "Dedifferentiated Liposarcoma"
                if subtype in ["Dedifferentiated"]
                else "Myxoid Liposarcoma"
                if subtype in ["Myxoid", "Myxoid Liposarcoma"]
                else "NR"
            ),
            pairing_status=(
                "No Cancer"
                if subtype in ["No Cancer", "Normal"] or "Normal" in diagnosis
                else "NR"
            ),
            pairing_site=(
                "No Cancer"
                if subtype in ["No Cancer", "Normal"] or "Normal" in diagnosis
                else "NR"
            ),
            i=i,
        )


# Export to Box & Load it to format the column width
unknown, unknown_inventory = format_and_export(file_name="Unknown", tissue_source="NR")

# Create the Combined Inventory (Full Inventory and Unique Combinations)
combined_inventory_unique_combos = pd.concat(
    [
        pancreatic,
        gastric,
        breast,
        lung,
        prostate,
        head_and_neck,
        gallbladder,
        intestines,
        liver,
        heart,
        lymphatic_system,
        GOJ,
        lower_body,
        kidney,
        female_reproductive_system,
        omentum,
        urinary_system,
        pelvis,
        chest,
        skin,
        digestive_tract,
        abscess,
        peritoneum,
        lymphatic_blood_vessels,
        oesophagus,
        connective_tissue,
        adrenal_gland,
        pleura,
        bile_duct,
        back,
        cortex,
        mesothelium,
        arms,
        smooth_muscle_tissue,
        ampulla,
        eye,
        shoulder,
        retroperitoneum,
        unknown,
        malereprod,
        muscle,
        multiple,
        plasma,
    ]
)
combined_inventory_unique_combos.shape


# Notes on errors in the combined inventory:
# - Rows from Stomach tissue source are all data (SOLVED - verify)
# - Rows of normal and normal adjacent specimens have incorrect or missing data in other columns (SOLVED - verify)
# - specimens of the primary type missing remaining data (SOLVED - verify)

# Strip duplicate columns from each DataFrame before concatenating
inventory_dfs = [
    df.loc[:, ~df.columns.duplicated()]
    for name, df in globals().items()
    if name.endswith("_inventory") and isinstance(df, pd.DataFrame)
]

combined_inventory = pd.concat(inventory_dfs, ignore_index=True)


# In[67]: A function to process the Others column
def process_inventory_row(row):
    # Initialize all potentially updated columns from the row, defaulting to 'NR' if missing
    updates = {
        "Molecular Information": row.get("Molecular Information", "NR"),
        "Grade": row.get("Grade", "NR"),
        "Paired Sample ID": row.get("Paired Sample ID", "NR"),
        "Paired Sample": row.get("Paired Sample", "NR"),
        "Anatomical Location": row.get("Anatomical Location", "NR"),
        "Anatomical Sub-Location": row.get("Anatomical Sub-Location", "NR"),
        "Specimen Type": row.get("Specimen Type", "NR"),
        "Histological Diagnosis (Original)": row.get(
            "Histological Diagnosis (Original)", "NR"
        ),
        "Histological Subtype": row.get("Histological Subtype", "NR"),
        "Pairing Site / Primary Tumour Site": row.get(
            "Pairing Site / Primary Tumour Site", "NR"
        ),
        "Others": str(row["Others"]),  # Start with the current 'Others' value
        "Tissue Source": row.get("Tissue Source", "NR"),
        "Tumor Size": row.get("Tumor Size", "NR"),
        "Age": row.get("Age", "NR"),
        "FollowUp_Result": row.get("FollowUp_Result", "NR"),
        "IHC_Molecular_Data": row.get("IHC_Molecular_Data", "NR"),
        "Chemo_Targeted_Tx": row.get("Chemo_Targeted_Tx", "NR"),
        "Year_Collection": row.get("Year_Collection", "NR"),
        "T": row.get("T", "NR"),
        "N": row.get("N", "NR"),
        "M": row.get("M", "NR"),
    }

    x = updates["Others"]  # Use the string value

    # Define keywords lists from both original functions
    HN_others_col = [
        "Larynx",
        "Left lateral oropharynx",
        "Pharynx",
        "Right oropharynx",
        "Left oropharynx",
        "Right oropharynx (right base of tongue)",
        "Right lateral oropharynx",
        "Left posterior tongue base",
        "Right tongue/right tonsil",
        "Oral cavity",
        "Anterior floor of mouth + anterior partial glossectomy",
        "Left retromolar trigone",
        "Anterior floor of mouth",
        "Left bucca/maxillary sulcus",
        "Right tongue",
        "Left tongue",
        "Tongue",
        "Soft palate",
        "Right lateral tongue",
        "Anterior left floor of mouth",
        "Left retromolar area",
        "Left mandible",
        "Left tongue base",
        "Anterior mandible",
        "Base of tongue",
        "Left base of tongue",
        "Right mandible",
        "Right posterior floor of mouth",
        "Right posterior mandible",
        "Right anterior/ventral tongue",
        "Left postero-lateral tongue",
        "Right ventral/lateral tongue",
        "Left tonsil",
        "Right tonsil",
        "Right oropharynx (tonsil)",
        "Left lateral oropharynx (tonsil)",
        "Right ethmoidal sinus",
        "Left maxilla",
        "Maxilla",
        "Right anterior maxilla",
        "Upper nose and ethmoid",
        "Buccal mucosa",
        "Left bucca",
        "Right pterygomandibular area",
        "left retromolar/pterygmaxillary area",
        "LARYNX CARCINOMA",
        "Larynx/vocal cord",
        "Larynx,Eroding throughEpiglottic cartilage and involving right carotid",
        "Total Larynx and Bilateral Selective Neck Dissections II-IV",
        "Total laryngectomy, left hemithyroidectomy & bilateral neck dissection II-V, with a small amount of left level I included",
        "Left Supraglottis",
        "Resection, Supra-cricoid partial laryngectomy",
        "Tonsil - right",
        "Tonsil - left",
        "Left Tonsil",
        "Lateral oropharyngeal wall",
        "Hypopharynx",
        "Pharynx and larynx",
        "Tissue, left lateral orapharynx",
        "Larynx and pharynx",
        "Right piriform fossa involving bilateral posterior larynx",
        "Right piriform fossa",
        "Piriform fossa, Uvula",
        "Left supraglottic involving left piriform fossa, left paraglottis, right posterior supraglottis and cricoid / arytenoid",
        "mouth floor",
        "Right partial glossectomy",
        "Right Mandible",
        "Glossectomy",
        "Left partial Glossectomy",
        "Left partial glossectomy",
        "Partial Glossectomy",
        "Floor of mouth",
        "Left Partial Glossectomy",
        "Left hemi-glossectomy",
        "Left floor of mouth",
        "Resection - left selective neck dissection levels",
        "Neck - LN dissection, selective lt lvl2a",
        "Left neck",
        "Mass, right level 2",
        "Neck levels II-IV",
        "Resection, left modified radical neck dissection, level II",
        "Resection, left selective neck dissection, levels I-III",
        "Neck - LN dissection - radical, lev 1/2/3/4/5b",
        "Tongue, partial right",
        "Tongue - partial right glossectomy",
        "Tongue - left hemi-glossectomy with bilateral neck dissections (levels I-IV)",
        "Tongue - right hemiglossectomy",
        "Tongue ",
        "Lesion from Right ventral tongue",
        "Tongue - partial, left",
        "Wide localExcision, cT2 left lateral tongue squamous carcinoma",
        "Tongue - Partial right lateral",
        "Resection, Left lateral tongue, left floor of mouth & left alveolar mucosa, with attached left selective neck dissection, levels I-IV",
        "Tongue - partial left",
        "Tongue/piriform fossa",
        "Posterior / base of tongueExtending onto lingual aspect ofEpiglottic",
        "tongue",
        "TONGUE LESION",
        "TONGUE",
        "Partial Mandibulectomy with floor of mouth",
        "LEFT EYEBROW",
        "NIVO 2020",
        "PEMBRO 2020",
        "Pembro/Nivo 2017",
        "Pembro/Nivo 2018",
        "Pembro/Nivo 2019",
        "Pembro 2021",
        "Nivo 2021",
        "Immunotherapy",
    ]  # Added all relevant strings from both original functions

    # Define all regex expressions (using e1 to e48)
    e1 = re.search(r"^T(\d) N(\d) (R|M)(\d)", x)
    e2 = re.search(r"^T(\d) N(.*)\s*(M)?(.*)?", x)
    e3 = re.search(r"^Her2\(-\),?ER\(-\),?\s*PR\(-\)", x)
    e4 = re.search(r"^MSI-([H|L])\s*(/Kras=(.*))?", x)
    e5 = re.search(
        r"^KRAS[:|=]([\s*A-Za-z0-9.>]+);?([\s*A-Za-z0-9.>]+)?[\/|,]EGFR:\s*([A-Za-z0-9.>\s*]+)?[:|;]?\s*([A-Za-z0-9.>]+)?",
        x,
        flags=re.IGNORECASE,
    )
    e6 = re.search(r"^KRAS=?(.*)", x, flags=re.IGNORECASE)
    e7 = re.search(r"MSI: ([A-Za-z\-]+)\s*\/\s*KRAS: \s*([A-Za-z0-9-]+)", x, flags=re.IGNORECASE)
    e8 = re.search(r"PDL1\(%\)\s*:\s*(\d{1,2}),\s*(.*)", x, flags=re.IGNORECASE)
    e9 = re.search(r"Gleason\s*(\d{1,2})\s*\((\d)\+(\d)\)", x, flags=re.IGNORECASE)
    e10 = re.search(r"Gleason:?\s*(\d)\+(\d)=(\d{1,2})", x, flags=re.IGNORECASE)
    e11 = re.search(r"Gleason\s*(?:(\d{1,2}$)|((\d)\+(\d)))", x, flags=re.IGNORECASE)
    e12 = re.search(r"(\d)\s*\+\s*(\d)\s*=?\s*\(?(\d)\)?", x, flags=re.IGNORECASE)
    e13 = re.search(r"locate:\s*(.*), size:\s*(\d*\.\d*cm)", x, flags=re.IGNORECASE)
    e14 = re.search(
        r"ER:(\d{1,2}) /PR:(\d{1,2}) /KI67:(\d{1,2}) /HER2:(\d{1,2})$", x, flags=re.IGNORECASE
    )
    e15 = re.search(
        r"ER:(\d{1,2}),PR:(\d{1,2}),KI-67:(\d{1,2}),HER2:(\d{1,2})$", x, flags=re.IGNORECASE
    )
    e16 = re.search(
        r"ER:(\d{1,2}),\s*PR:(\d{1,2}),\s*KI-67:(\d{1,2}),\s*HER2:([0,/])$",
        x,
        flags=re.IGNORECASE,
    )
    e17 = re.search(r"PDL1\s*(%)?\s*:\s*(\d{1,3})(-\d{1,3})?", x, flags=re.IGNORECASE)
    e18 = re.search(r"PDL1\s*\(\%\):(\d{1,2})$", x, flags=re.IGNORECASE)
    e19 = re.search(r"Gleason\s* (\d{1})(\(\d{1}\+\d{1}\))", x, flags=re.IGNORECASE)
    e20 = re.search(r"(\d)\s*\+\s*(\d)\s*[=(](\d)\)?", x, flags=re.IGNORECASE)
    e21 = re.search(r"(\d{1,2})(\(\d\+\d\))", x, flags=re.IGNORECASE)
    e22 = re.search(r"Matched with (.+)", x, flags=re.IGNORECASE)
    e23 = re.search(r"mets\((.+)\)", x, flags=re.IGNORECASE)
    e24 = re.search(r"KRAS:\s*([A-Z]+)[,|/]\s*EGFR:\s*([A-Z]+\s*[A-Z]+)", x, flags=re.IGNORECASE)
    e25 = re.search(r"Gleason (\d$)?(\d \((\d\+\d\)))?((\d)\+(\d))?", x, flags=re.IGNORECASE)
    e26 = re.search(r"met disease: (Yes)?(.+)?", x, flags=re.IGNORECASE)
    e27 = re.search(
        r"^(Left|Right) ((pulmonectomy|pneuemectomy|pulmectomy)|(lobectomy))",
        x,
        flags=re.IGNORECASE,
    )
    e28 = re.search(
        r"^(MSI$|MSI-H|MSI-L|MSI:MSS|MSI: MSI-H)(\s*/KRAS(:|=)\s*(wt|G\d{1,2}[A-Z]))?",
        x,
        flags=re.IGNORECASE,
    )
    e29 = re.search(r"^(MSI: MSS|MSS)\s*/KRAS: (wt|[A-Z]\d{1,2}[A-Z])$", x, flags=re.IGNORECASE)
    e30 = re.search(r"ALK\+, (\d)\+\/ Production year:\s*(\d+)", x)
    e31 = re.search(r"(?:PD-L1|PDL1)\s*(\(%\))?\s*:\s*(\d{1,3})", x, flags=re.IGNORECASE)
    e33 = re.search(
        r"^Exon 21 Mutation ([A-Za-z0-9>.]+); p.\(([A-Za-z0-9]+)\), (([A-Za-z0-9]+)\/(EGFREXON 18 MUTATION ([A-Za-z0-9>.]+); p.\(([A-Za-z0-9]+)\), ([A-Za-z0-9]+)))?",
        x,
        flags=re.IGNORECASE,
    )
    e34 = re.search(r"Exon 19 Deletion \(?([A-Za-z0-9]+) (\([A-Za-z0-9_]+\))\)?", x)
    e35 = re.search(r"Pre$|Pre[-|\s]treatment|Pre ", x, flags=re.IGNORECASE)
    e36 = re.search(
        r"Post$|Post treatment|Post ([A-Za-z]+) treatment$|Post-treatment|Post\xa0|Post ",
        x,
        flags=re.IGNORECASE,
    )
    e37 = re.search(
        r"Ln. Mets|Metastatic - Lymph Node|Metastatic|meta to Liver\(([A-Za-z0-9\-]+)\)|meta to T56|mets",
        x,
        flags=re.IGNORECASE,
    )
    e38 = re.search(r"^([0-9.]+)\s*cm$", x, flags=re.IGNORECASE)
    e39 = re.search(
        r"^T cell;\s*LN, neck\/([0-9.]+ cm)(?:\/(?:EBV\([+-]\)))?(?:,\s*alive:\s*(\d+))?(?:[\/|,]?(?:EBV\([+-]\)))?(?:,\s*alive:\s*(\d+))?$",
        x,
    )
    e40 = re.search(
        r"^T cell\s*;(EBV\([\+|\-]\))\s*[,|\/]\s*(alive|death): (\d{1,2})",
        x,
        flags=re.IGNORECASE,
    )
    e41 = re.search(
        r"^T cell\s*;\s*EBV\(([\+|\-])\)(\/|;)?\s*LN,\s*([A-Za-z]+)\/(\d.\d cm)\/\s*(alive|death):\s*(\d{1,2})",
        x,
        flags=re.IGNORECASE,
    )
    e42 = re.search(
        r"LN,\s*neck\/([0-9.]+ cm)(?:[\/|,]\s*(?:(T cell;EBV\([+-]\))|((?:(death|alive)):\s*(\d+))))*(?:[\/|,]\s*(?:(T cell;EBV\([+-]\))|((?:(death|alive)):\s*(\d+))))*",
        x,
        flags=re.IGNORECASE,
    )
    e43 = re.search(
        r"T cell;EBV\((\+|\-)\)\/LN,\s*(axilla|subclavian)\/(\d cm)\/\s*(alive|death)\s*:\s*(\d{1,2})",
        x,
        flags=re.IGNORECASE,
    )
    e44 = re.search(r"^EBV, ([a-z]+): (\d{1,2})", x, flags=re.IGNORECASE)
    e45 = re.search(r"^MSS \/KRAS=([A-Za-z0-9]+)", x, flags=re.IGNORECASE)
    e46 = re.search(r"KRAS", x, flags=re.IGNORECASE)
    e47 = re.search(r"^(\d.\d) cm, ([a-z]+): (\d{1,2})", x, flags=re.IGNORECASE)
    e48 = re.search(r"lymph node", x, flags=re.IGNORECASE)
    e49 = re.search(r"Gleason\s+(\d)\s*\+\s*(\d)", x, flags=re.IGNORECASE)
    e50 = re.search(r"(\d)\s*\+\s*(\d)\s*=\s*(\d)", x, flags=re.IGNORECASE)

    # Molecular/Grade/TNM/Paired Sample (Mutually Exclusive or Specific Matches)

    # TNM Staging
    if e1:
        updates["T"] = f"T{e1.group(1)}"
        updates["N"] = f"N{e1.group(2)}"
        updates["M"] = f"M{e1.group(4)}" if e1.group(3) == "M" else "NR"
        updates["Others"] = "NR"
    elif e2:
        updates["T"] = f"T{e2.group(1)}"
        updates["N"] = f"N{e2.group(2)}"
        updates["M"] = f"M{e2.group(3).lower()}" if e2.group(3) else "NR"
        updates["Others"] = "NR"

    # Molecular Information (Hormone/MSI/KRAS/EGFR/PDL1)
    elif e7:  # MSI & KRAS
        updates["Molecular Information"] = (
            f"MSI:{e7.group(1).replace('MSI-H', 'High').replace('unknown', 'Unknown')} KRAS:{e7.group(2).replace('unknown', 'Unknown')}"
        )
        updates["Others"] = "NR"
    elif e5:  # KRAS & EGFR
        kras_full = f"{e5.group(1).strip()} {e5.group(2).strip() if e5.group(2) else ''}".strip()
        egfr_full = f"{e5.group(3).strip() if e5.group(3) else ''} {e5.group(4).strip() if e5.group(4) else ''}".strip()
        updates["Molecular Information"] = f"KRAS:{kras_full} ; EGFR:{egfr_full}"
        updates["Others"] = "NR"
    elif e24:  # KRAS & EGFR (alternative format)
        updates["Molecular Information"] = (
            f"KRAS:{e24.group(1).strip()} ; EGFR:{e24.group(2).strip()}"
        )
        updates["Others"] = "NR"
    elif e45:  # MSS / KRAS
        updates["Molecular Information"] = f"MSI:MSS KRAS:{e45.group(1)}"
        updates["Specimen Type"] = "P"
        updates["Others"] = "NR"
    elif e28:  # MSI-H/L/MSS
        if e28.group(1) in ["MSI: MSI-H", "MSI-H"]:
            updates["Molecular Information"] = (
                f"MSI:High KRAS:{e28.group(4).upper()}" if e28.group(3) else "MSI:High"
            )
            updates["Others"] = "NR"
        elif e28.group(1) == "MSI-L":
            updates["Molecular Information"] = "MSI:Low"
            updates["Others"] = "NR"
    elif e29:  # MSI: MSS / KRAS
        updates["Molecular Information"] = f"MSI:MSS ; KRAS:{e29.group(2).upper()}"
        updates["Others"] = "NR"
    elif x == "MSS":
        updates["Molecular Information"] = "MSS"
        updates["Others"] = "NR"
    elif "MSS/ EBV" in x:
        updates["Molecular Information"] = "MSS EBV"
        updates["Others"] = "NR"
    elif e14:  # Hormone Receptor (ER/PR/KI67/HER2) - format 1
        updates["Molecular Information"] = (
            f"ER:{e14.group(1)} ; PR:{e14.group(2)} ; KI67:{e14.group(3)} HER2:{e14.group(4)}"
        )
        updates["Others"] = "NR"
    elif e15:  # Hormone Receptor - format 2
        updates["Molecular Information"] = (
            f"ER:{e15.group(1)} ; PR:{e15.group(2)} ; KI67:{e15.group(3)} ; HER2:{e15.group(4)}"
        )
        updates["Others"] = "NR"
    elif e16:  # Hormone Receptor - format 3
        updates["Molecular Information"] = (
            f"ER:{e16.group(1)} ; PR:{e16.group(2)} ; KI67:{e16.group(3)} ; HER2:{e16.group(4)}"
        )
        updates["Others"] = "NR"
    elif e3:  # HER2- ER- PR-
        updates["Molecular Information"] = "HER2- ER- PR-"
        updates["Others"] = "NR"
    elif e33:  # EGFR Exon 21/18
        if e33.group(3):
            updates["Molecular Information"] = (
                f"Exon 21 Mutation: {e33.group(1)} p.({e33.group(2)}) ; {e33.group(4)}/EGFR Exon 18 Mutation: {e33.group(6)} p.({e33.group(7)}) {e33.group(8)}"
            )
            updates["Others"] = "NR"
        else:
            updates["Molecular Information"] = (
                f"Exon 21 Mutation: {e33.group(1)} p.({e33.group(2)})"
            )
            updates["Others"] = "NR"
    elif e34:  # EGFR Exon 19 Deletion
        updates["Molecular Information"] = (
            f"Exon 19 Deletion: {e34.group(1)} {e34.group(2)}"
        )
        updates["Others"] = "NR"
    elif x == "EGFREXON 21 MUTATION c.2573T>G; p.(Leu858Arg), L858R ":
        updates["Molecular Information"] = "EGFR Exon 21 Mutation c.2573T>G; p.(L858R)"
        updates["Others"] = "NR"
    elif x == "EGFR mutant":
        updates["Molecular Information"] = "EGFR+"
        updates["Others"] = "NR"
    elif x == "EGFR wt":
        updates["Molecular Information"] = "EGFR:WT"
        updates["Others"] = "NR"
    elif e8:  # PDL1 & viable cells
        updates["Molecular Information"] = (
            f"PDL1:{e8.group(1)} ({e8.group(2).strip().title().replace('Vialbe', 'Viable').replace('Mm', 'mm')})"
        )
        updates["Others"] = "NR"
    elif e17:  # PDL1 range
        updates["Molecular Information"] = (
            f"PDL1:{e17.group(2)}{e17.group(3) if e17.group(3) else ''}"
        )
        updates["Others"] = "NR"
    elif e18:  # PDL1 percentage
        updates["Molecular Information"] = f"PDL1:{e18.group(1)}"
        updates["Others"] = "NR"
    elif e31:  # PDL1 (generic)
        updates["Molecular Information"] = f"PDL1:{e31.group(2)}"
        updates["Others"] = "NR"
    elif e30:  # ALK+
        updates["Molecular Information"] = f"ALK+, {e30.group(1)}+"
        updates["Year_Diagnosis"] = e30.group(2)
        updates["Others"] = "NR"
    elif x == "BRAF=mt":
        updates["Molecular Information"] = "BRAF:mt"
        updates["Others"] = "NR"
    elif x == "BRAF_mt":
        updates["Molecular Information"] = "BRAF-Mutated"
        updates["Others"] = "NR"
    elif e6:  # General KRAS Match (last check)
        updates["Molecular Information"] = (
            f"KRAS:{e6.group(1).strip().replace('-', '')}" if e6.group(1) else "KRAS"
        )
        updates["Others"] = "NR"
    elif x == "KRAS":
        updates["Molecular Information"] = "KRAS:WT"
        updates["Others"] = "NR"

    # Gleason Grade
    if e9:
        updates["Grade"] = f"Gleason {e9.group(1)} ({e9.group(2)}+{e9.group(3)})"
        updates["Others"] = "NR"
    elif e10:
        updates["Grade"] = (
            f"Gleason {e10.group(3).strip()} ({e10.group(1).strip()}+{e10.group(2).strip()})"
        )
        updates["Others"] = "NR"
    elif e12:
        updates["Grade"] = f"Gleason {e12.group(3)} ({e12.group(1)}+{e12.group(2)})"
        updates["Others"] = "NR"
    elif e19:
        updates["Grade"] = f"Gleason {e19.group(1)} {e19.group(2)}"
        updates["Others"] = "NR"
    elif e20:
        updates["Grade"] = f"Gleason {e20.group(3)} ({e20.group(1)}+{e20.group(2)})"
        updates["Others"] = "NR"
    elif e21:
        updates["Grade"] = f"Gleason {e21.group(1)} {e21.group(2)}"
        updates["Others"] = "NR"
    elif e25:
        if e25.group(1):
            updates["Grade"] = f"Gleason {e25.group(1).strip()}"
            updates["Others"] = "NR"
        elif e25.group(2):
            updates["Grade"] = f"Gleason {e25.group(2).strip()}"
            updates["Others"] = "NR"
        elif e25.group(4):
            updates["Grade"] = (
                f"Gleason {int(e25.group(5)) + int(e25.group(6))} ({e25.group(4).strip()})"
            )
            updates["Others"] = "NR"
    elif e50:
        primary = e50.group(1)
        secondary = e50.group(2)
        updates["Grade"] = f"Gleason {int(e50.group(3))} ({primary}+{secondary})"
        updates["Others"] = "NR"
    elif e49:
        primary = int(e49.group(1))
        secondary = int(e49.group(2))
        updates["Grade"] = f"Gleason {primary + secondary} ({primary}+{secondary})"
        updates["Others"] = "NR"

    elif x in ["WHO G4 ICD-O9401/3"]:
        updates["Grade"] = 4
        updates["Specimen Type"] = "P"
        updates["Histological Diagnosis (Original)"] = "Astrocytoma"
        updates["Histological Subtype"] = "Anaplastic Astrocytoma"
        updates["Others"] = "NR"
    elif x == 4:
        updates["Grade"] = 4
        updates["Others"] = "NR"

    # Paired Sample Info
    if e22:
        updates["Paired Sample ID"] = e22.group(1).strip()
        updates["Paired Sample"] = "P"
        updates["Others"] = "NR"
    elif e23:
        updates["Paired Sample ID"] = e23.group(1).strip()
        updates["Paired Sample"] = "P"
        updates["Others"] = "NR"
    elif e26:
        if e26.group(2):
            updates["Pairing Site / Primary Tumour Site"] = e26.group(2).strip().title()
            updates["Others"] = "NR"
        elif e26.group(1):
            updates["Paired Sample"] = "NR"
            updates["Others"] = "NR"

    # General Anatomical/Size/Follow-up (T-Cell Lymphoma, Lung Resection, etc.)
    elif e13:
        loc = e13.group(1).strip().title()
        updates["Tumor Size"] = f"{e13.group(2)}"
        if loc == "Retroperitoneum":
            updates["Tissue Source"] = "Retroperitoneum"
            updates["Anatomical Location"] = "Retroperitoneum"
            updates["Others"] = "NR"
        elif loc in ["Abdominal Cavity", "Trumk"]:
            updates["Tissue Source"] = "Trunk"
            updates["Anatomical Location"] = (
                "Abdominal Cavity" if loc == "Abdominal Cavity" else "NR"
            )
            updates["Others"] = "NR"
        elif loc == "Extremity":
            updates["Tissue Source"] = "Extremity"
            updates["Anatomical Location"] = "NR"
            updates["Others"] = "NR"
    elif e27:
        updates["Anatomical Location"] = "Lung"
        updates["Others"] = "NR"
        if e27.group(3):
            updates["Anatomical Sub-Location"] = (
                f"Lung ({e27.group(1).strip().title()})"
            )
            updates["Others"] = "NR"
        elif e27.group(4):
            updates["Anatomical Sub-Location"] = (
                f"Lung ({e27.group(1).strip().title()} Lobe)"
            )
            updates["Others"] = "NR"
    elif x == "supraglo TTic":
        updates["Anatomical Sub-Location"] = "Supraglottis"
        updates["Specimen Type"] = "P"
        updates["Histological Diagnosis (Original)"] = "Cancer"
        updates["Histological Subtype"] = "Supraglottic Cancer"
        updates["Others"] = "NR"
    elif x == "Plasma":
        updates["Anatomical Sub-Location"] = (
            "Plasma"
            if updates["Anatomical Sub-Location"] == "NR"
            else updates["Anatomical Sub-Location"] + " (Plasma)"
        )
        updates["Others"] = "NR"
    elif e37:
        if x in ["Ln. Mets.", "Metastatic - Lymph node"]:
            updates["Pairing Site / Primary Tumour Site"] = "Lymph Nodes"
            updates["Others"] = "NR"
        elif "meta to Liver" in x:
            updates["Pairing Site / Primary Tumour Site"] = "Liver"
            updates["Others"] = "NR"
            if e37.group(1):
                updates["Paired Sample"] = "P"
        elif x in ["meta to T56"]:
            updates["Paired Sample"] = "P"
            updates["Others"] = "NR"
        elif x in ["Metastatic"]:
            updates["Specimen Type"] = "P"
            updates["Others"] = "NR"
        elif x in ["Lung Mets."]:
            updates["Pairing Site / Primary Tumour Site"] = "Lung"
            updates["Others"] = "NR"
        elif x in ["M1 Surrenal Mets."]:
            updates["Pairing Site / Primary Tumour Site"] = "Adrenal Gland"
            updates["Others"] = "NR"
    elif e38:
        updates["Tumor Size"] = f"{e38.group(1)} cm"
        updates["Others"] = "NR"

    # T-Cell Lymphoma Group (e39-e44)
    elif e39:
        updates.update(
            {
                "Histological Diagnosis (Original)": "T-Cell Lymphoma",
                "Tumor Size": e39.group(1),
                "FollowUp_Result": "alive",
            }
        )
        if "LN" in x:
            updates["Anatomical Location"] = "Lymph Nodes"
        if "neck" in x:
            updates["Anatomical Sub-Location"] = "Lymph Nodes (Neck)"
        if "EBV(+)" in x:
            updates["Molecular Information"] = "EBV+"
        if "EBV(-)" in x:
            updates["Molecular Information"] = "EBV-"
        updates["Others"] = "NR"
    elif e40:
        updates.update(
            {
                "Histological Diagnosis (Original)": "T-Cell Lymphoma",
                "FollowUp_Result": e40.group(2),
                "Molecular Information": "EBV+",
            }
        )
        updates["Others"] = "NR"
    elif e41:
        loc_map = {
            "retroperitoneal": "Retroperitoneum (Lymph Nodes)",
            "inguinal": "Lymph Nodes (Inguinal)",
        }
        updates.update(
            {
                "Histological Diagnosis (Original)": "T-Cell Lymphoma",
                "Molecular Information": f"EBV{e41.group(1)}",
                "Tumor Size": e41.group(4),
                "FollowUp_Result": e41.group(5),
            }
        )
        updates["Anatomical Location"] = (
            "Retroperitoneum"
            if e41.group(3) == "retroperitoneal"
            else "Lymph Nodes"
            if e41.group(3) == "inguinal"
            else "NR"
        )
        updates["Anatomical Sub-Location"] = loc_map.get(e41.group(3), "NR")
        updates["Others"] = "NR"
    elif e42:
        updates.update(
            {
                "Anatomical Location": "Lymph Nodes",
                "Anatomical Sub-Location": "Lymph Nodes (Neck)",
                "Tumor Size": e42.group(1),
                "Histological Diagnosis (Original)": "T-Cell Lymphoma",
                "Molecular Information": "EBV+",
                "FollowUp_Result": e42.group(5),
            }
        )
        updates["Others"] = "NR"
    elif e43:
        updates.update(
            {
                "Histological Diagnosis (Original)": "T-Cell Lymphoma",
                "Molecular Information": f"EBV{e43.group(1)}",
                "Anatomical Location": "Lymph Nodes",
                "Tumor Size": e43.group(3),
                "FollowUp_Result": e43.group(4),
            }
        )
        updates["Anatomical Sub-Location"] = (
            "Lymph Nodes (Axilla)"
            if e43.group(2) == "axilla"
            else "Lymph Nodes (Subclavian)"
        )
        updates["Others"] = "NR"
    elif e44:
        updates.update(
            {"Molecular Information": "EBV", "FollowUp_Result": e44.group(1)}
        )
        updates["Others"] = "NR"
    elif e47:
        updates.update({"Tumor Size": e47.group(1), "FollowUp_Result": e47.group(2)})
    elif x in ["T cell; 3 cm/EBV(+), death: 9"]:
        updates.update(
            {
                "Histological Diagnosis (Original)": "T-Cell Lymphoma",
                "Tumor Size": "3 cm",
                "Molecular Information": "EBV+",
                "FollowUp_Result": "death",
            }
        )
        updates["Others"] = "NR"

    # Pre/Post Treatment
    if e35:
        updates["Histological Subtype"] = (
            updates["Histological Subtype"] + " (Pre-Treatment)"
        )
        updates["Others"] = "NR"
    elif e36:
        updates["Histological Subtype"] = (
            updates["Histological Subtype"] + " (Post Neoadjuvant Treatment)"
            if e36.group(1)
            else updates["Histological Subtype"] + " (Post-Treatment)"
        )
        updates["Others"] = "NR"

    # Head & Neck Classification
    if x in HN_others_col:
        updates["Tissue Source"] = "Head & Neck"
        updates["Others"] = "NR"  # Mark as processed
        if x in ["Larynx", "supraglo TTic"]:
            updates["Anatomical Location"] = "Larynx"
        elif x in [
            "Left tonsil",
            "Right tonsil",
            "Right oropharynx (tonsil)",
            "Left lateral oropharynx (tonsil)",
        ]:
            updates["Anatomical Location"] = "Tonsil"
        elif x in [
            "Left lateral oropharynx",
            "Pharynx",
            "Right oropharynx",
            "Left oropharynx",
            "Right oropharynx (right base of tongue)",
            "Right lateral oropharynx",
            "Left posterior tongue base",
            "Right tongue/right tonsil",
            "Left tongue base",
            "Base of tongue",
            "Left base of tongue",
        ]:
            updates["Anatomical Location"] = "Pharynx"
        elif x in [
            "Oral cavity",
            "Anterior floor of mouth + anterior partial glossectomy",
            "Left retromolar trigone",
            "Anterior floor of mouth",
            "Left bucca/maxillary sulcus",
            "Right tongue",
            "Left tongue",
            "Tongue",
            "Soft palate",
            "Right lateral tongue",
            "Anterior left floor of mouth",
            "Left retromolar area",
            "Left mandible",
            "Anterior mandible",
            "Right mandible",
            "Right posterior floor of mouth",
            "Right posterior mandible",
            "Right anterior/ventral tongue",
            "Left postero-lateral tongue",
            "Right ventral/lateral tongue",
            "mouth floor",
            "Right partial glossectomy",
            "Glossectomy",
            "Left partial Glossectomy",
            "Left partial glossectomy",
            "Partial Glossectomy",
            "Floor of mouth",
            "Left Partial Glossectomy",
            "Left hemi-glossectomy",
            "Left floor of mouth",
            "Tongue, partial right",
            "Tongue - partial right glossectomy",
            "Tongue - left hemi-glossectomy with bilateral neck dissections (levels I-IV)",
            "Tongue - right hemiglossectomy",
            "Tongue ",
            "Lesion from Right ventral tongue",
            "Tongue - partial, left",
            "Wide localExcision, cT2 left lateral tongue squamous carcinoma",
            "Tongue - Partial right lateral",
            "Resection, Left lateral tongue, left floor of mouth & left alveolar mucosa, with attached left selective neck dissection, levels I-IV",
            "Tongue - partial left",
            "TONGUE LESION",
            "TONGUE",
        ]:
            updates["Anatomical Location"] = "Oral Cavity"
        if x == "Anterior floor of mouth + anterior partial glossectomy":
            updates["Anatomical Sub-Location"] = (
                "Mouth (Anterior Floor) And Tongue (Anterior)"
            )
        elif x == "Right posterior floor of mouth":
            updates["Anatomical Sub-Location"] = "Mouth (Right Posterior Floor)"
        elif x == "Left bucca/maxillary sulcus":
            updates["Anatomical Sub-Location"] = "Mouth (Left Buccal Maxillary Sulcus)"
        elif x == "Right oropharynx (right base of tongue)":
            updates["Anatomical Sub-Location"] = "Oropharynx (Right Tongue Base)"

        if x in [
            "LARYNX CARCINOMA",
            "Larynx/vocal cord",
            "Larynx,Eroding throughEpiglottic cartilage and involving right carotid",
            "Total Larynx and Bilateral Selective Neck Dissections II-IV",
            "Total laryngectomy, left hemithyroidectomy & bilateral neck dissection II-V, with a small amount of left level I included",
            "Left Supraglottis",
            "Resection, Supra-cricoid partial laryngectomy",
        ]:
            is_larynx_neck = x in [
                "Larynx,Eroding throughEpiglottic cartilage and involving right carotid",
                "Total Larynx and Bilateral Selective Neck Dissections II-IV",
                "Total laryngectomy, left hemithyroidectomy & bilateral neck dissection II-V, with a small amount of left level I included",
            ]
            updates["Anatomical Location"] = (
                "Larynx - Neck" if is_larynx_neck else "Larynx"
            )

            if x == "Larynx/vocal cord":
                updates["Anatomical Sub-Location"] = "Vocal Cord"
            elif (
                x
                == "Larynx,Eroding throughEpiglottic cartilage and involving right carotid"
            ):
                updates["Anatomical Sub-Location"] = (
                    "Larynx (Eroding Through Epiglottic Cartilage And Involving Right Carotid)"
                )
            elif x == "Total Larynx and Bilateral Selective Neck Dissections II-IV":
                updates["Anatomical Sub-Location"] = (
                    "Larynx (Total) And Neck (Both Sides - Levels 2-4)"
                )
            elif (
                x
                == "Total laryngectomy, left hemithyroidectomy & bilateral neck dissection II-V, with a small amount of left level I included"
            ):
                updates["Anatomical Sub-Location"] = (
                    "Larynx (Total), Thyroid (Left Lobe), And Neck (Both Sides - Lymph Nodes Level 2-5 And Part of Level 1)"
                )
            elif x == "Left Supraglottis":
                updates["Anatomical Sub-Location"] = "Supraglottis (Left)"
            elif x == "Resection, Supra-cricoid partial laryngectomy":
                updates["Anatomical Sub-Location"] = "Larynx (Part)"
            elif x == "LARYNX CARCINOMA":
                updates["Anatomical Sub-Location"] = "NR"  # Keep existing value if 'NR'

    # Other Specific Matches
    if x in [
        "Not stated",
        "GENERAL",
        "Immunotherapy",
        "Nivo 2021",
        "NIVO 2020",
        "PEMBRO 2020",
        "Pembro/Nivo 2017",
        "Pembro/Nivo 2018",
        "Pembro/Nivo 2019",
        "Pembro 2021",
        "HNSCC",
        "SCLC",
        "No information",
        "poor quality",
        "poor qualiaty",
        "WT",
        "Immunotherapy",
        "only remained tumor tissue",
        "according to TNM5, pT2, N0, Mx, R0,according to TNM7, pT4, N0, Mx. (due to periductal spread)",
        "According to TNM 6: pT3 N1 Mx V1 R1, According to TNM 7 the staging would be: pT2b N1 Mx V1 R1",
        "V1,Pne1,R0",
        "R0",
        "V1, G1, R0",
        "R1",
        "V1,R0",
        "R1/R2",
        "V1",
        "V0",
        "V0,G2,R1",
        "C34 Malignant neoplasm of bronchus and lung",
        "Primary tumor",
        "Resection tuberosity (main specimen is LH19-50983)",
        "pT2 N1a Mx",
        "pT1c N2a Mx",
        2,
        3,
        "3",
        "primary chemo change ",
        "frontal",
        "frontotemporal",
        "parietal",
        "4",
        "-",
        "Not specified",
        "Unknown",
        "\u3000",
    ]:
        updates["Others"] = "NR"
    elif x in ["Apparent surfaceExpression of IgMK; Weak CD5 staining"]:
        updates["Molecular Information"] = (
            "Apparent Surface Expression Of IgML; Weak CD5 Staining"
        )
        updates["Others"] = "NR"

    elif x in ["Side: Left", "Side: Right", "RIGHT SIDED", "LEFT SIDED"]:
        updates["Anatomical Location"] = "Colon"
        updates["Anatomical Sub-Location"] = (
            "Colon (Right)" if x in ["Side: Right", "RIGHT SIDED"] else "Colon (Left)"
        )
        updates["Others"] = "NR"
    elif x in [
        "Axillary lymph node",
        "Axillary lymph nodes",
        "Inguinal lymph node",
        "Left Inguinal Lymph Node",
        "Left Neck Lymph Node",
        "Lymph Node Pri.",
        "Lymph node",
        "Lymph node - Neck",
        "Lymph node -Epipharynx",
        "Mediastinal lymph nodes",
        "One lymph node",
        "One lymph node ",
        "Para-oesophageal lymph nodes ",
        "Paraoesophageal lymph node",
        "Perigastric lymph node",
        "Retroperitoneal lymph nodes",
    ]:
        updates["Anatomical Location"] = (
            updates["Anatomical Location"]
            if updates["Anatomical Location"] != "NR"
            else "Lymph Nodes"
        )
        if updates["Anatomical Sub-Location"] == "NR":
            sub_map = {
                "Axillary": "Lymph Nodes (Axillary)",
                "Inguinal": "Lymph Nodes (Inguinal)",
                "Left Inguinal": "Lymph Nodes (Left Inguinal)",
                "Left Neck": "Lymph Nodes (Left Neck)",
                "Neck": "Lymph Nodes (Neck)",
                "Epipharynx": "Lymph Nodes (Epipharynx)",
                "Mediastinal": "Lymph Nodes (Mediastinal)",
                "Para-oesophageal": "Lymph Nodes (Para-Oesophageal)",
                "Perigastric": "Lymph Nodes (Perigastric)",
                "Retroperitoneal": "Lymph Nodes (Retroperitoneal)",
            }
            updates["Anatomical Sub-Location"] = sub_map.get(
                x.split("lymph node")[0]
                .strip()
                .replace("-", "")
                .title()
                .replace(" ", ""),
                "NR",
            )
        updates["Others"] = "NR"
    elif x in [
        "HGNEC",
        "skin",
        "GIEndo",
        "Nefrectomy",
        "normal adrenal gland",
        "Panhysterectomy I type",
        "Just above oesophagus distally",
        "Greater curve",
        "Tissue around GOJ ",
        "Pre-treatment",
        "2",
    ]:
        updates["Others"] = "NR"
    elif x == "mucinous (8cm)":
        updates["Histological Subtype"] = "Mucinous Carcinoma"
        updates["Tumor Size"] = "8 cm"
        updates["Others"] = "NR"
    elif x == "serous":
        updates["Histological Subtype"] = "Serous Carcinoma"
        updates["Others"] = "NR"
    elif x == "vulva":
        updates["Anatomical Location"] = "Vulva"
        updates["Others"] = "NR"
    elif x == "vagina":
        updates["Anatomical Location"] = "Vagina"
        updates["Others"] = "NR"
    elif x == "Part of matched sets":
        updates["Paired Sample"] = "P"
        updates["Others"] = "NR"
    elif x == "L OV":
        updates["Anatomical Location"] = "Ovary"
        updates["Anatomical Sub-Location"] = "Ovary (Left)"
        updates["Others"] = "NR"
    elif x == "R OV":
        updates["Anatomical Location"] = "Ovary"
        updates["Anatomical Sub-Location"] = "Ovary (Right)"
        updates["Others"] = "NR"
    elif x == "R OV+TUBAL":
        updates["Anatomical Location"] = "Ovary/Fallopian Tube"
        updates["Anatomical Sub-Location"] = "Ovary (Right) And Fallopian Tube"
        updates["Others"] = "NR"
    elif x == "L OV + tube":
        updates["Anatomical Location"] = "Ovary/Fallopian Tube"
        updates["Anatomical Sub-Location"] = "Ovary (Left) And Fallopian Tube"
        updates["Others"] = "NR"
    elif x in ["Perigastric nodes"]:
        updates["Anatomical Sub-Location"] = "Lymph Nodes (Perigastric)"
        updates["Others"] = "NR"
    elif x == "Lymph node ":
        updates["Anatomical Sub-Location"] = "Lymph Nodes"
        updates["Others"] = "NR"
    elif (
        x
        == "CD23 +, CD5 +, CD20 +, Bcl-2 +, CD43 +, Ki-67 15% +;\nZap70 -, CD3 -, Bcl-6 -, CD10 -, Cyclin D1 -"
    ):
        updates["Molecular Information"] = (
            "CD23+ CD5+ CD20+ Bcl-2+ CD43+ Ki-67 15%+ ; Zap70- CD3- Bcl-6- CD10- Cyclin D1-"
        )
        updates["Others"] = "NR"
    elif x == "Skin - Gluteal":
        updates["Anatomical Location"] = "Skin"
        updates["Anatomical Sub-Location"] = "Skin (Gluteal)"
        updates["Others"] = "NR"
    elif x == "Left Inguinal Lymph Node\xa0":
        updates["Anatomical Location"] = "Lymph Nodes"
        updates["Anatomical Sub-Location"] = "Lymph Nodes (Left Inguinal)"
        updates["Others"] = "NR"
    elif x == "Stomach":
        updates["Anatomical Location"] = "Stomach"
        updates["Others"] = "NR"
    elif x == "Testis":
        updates["Anatomical Location"] = "Testis"
        updates["Others"] = "NR"
    elif x == "Left Submandibular Gland\xa0":
        updates["Anatomical Location"] = "Salivary Gland"
        updates["Anatomical Sub-Location"] = "Salivary Gland (Left Submandibular)"
        updates["Others"] = "NR"
    elif x == "Left Supraclavicular":
        updates["Anatomical Location"] = "Lymph Nodes"
        updates["Anatomical Sub-Location"] = "Lymph Nodes (Left Supraclavicular)"
        updates["Others"] = "NR"
    elif x == "Left Neck Node\xa0":
        updates["Anatomical Location"] = "Lymph Nodes"
        updates["Anatomical Sub-Location"] = "Lymph Nodes (Left Neck)"
        updates["Others"] = "NR"
    elif x == "Right Partoid\xa0":
        updates["Anatomical Location"] = "Salivary Gland"
        updates["Anatomical Sub-Location"] = "Salivary Gland (Right Parotid)"
        updates["Others"] = "NR"
    elif x == "Left Neck Lymph Node\xa0":
        updates["Anatomical Location"] = "Lymph Nodes"
        updates["Anatomical Sub-Location"] = "Lymph Nodes (Left Neck)"
        updates["Others"] = "NR"
    elif x == "B cell, alive: 17":
        updates["Histological Diagnosis (Original)"] = "B-Cell Lymphoma"
        updates["FollowUp_Result"] = "alive"
        updates["Others"] = "NR"
    elif x == "Spleen":
        updates["Anatomical Location"] = "Spleen"
        updates["Others"] = "NR"
    elif x == "Right Level I Nodes\xa0":
        updates["Anatomical Location"] = "Lymph Nodes"
        updates["Anatomical Sub-Location"] = "Lymph Nodes (Right Level I)"
        updates["Others"] = "NR"
    elif x == "Year of collection: 2019":
        updates["Year_Collection"] = "2019"
        updates["Others"] = "NR"
    elif x in ["B cell , alive: 74", "B cell, death: 7"]:
        updates["Histological Diagnosis (Original)"] = "B-Cell Lymphoma"
        updates["FollowUp_Result"] = "alive" if "alive" in x else "death"
        updates["Others"] = "NR"
    elif x == "Strongly positive for CD20; Negative for CD5 and CD23":
        updates["Molecular Information"] = "CD20+(Strong) CD5- CD23-"
        updates["Others"] = "NR"
    elif x == "Positive for CD20, CD5 and Cyclin D1":
        updates["Molecular Information"] = "CD20+ CD5+ Cyclin D1+"
        updates["Others"] = "NR"
    elif (
        x
        == "Not stated. Previous biopsy 3 months prior was positive for CD20, CD5, CD79a and Cyclin D1"
    ):
        updates["Molecular Information"] = "CD20+ CD5+ CD79a+ Cyclin D1+"
        updates["Others"] = "NR"
    elif (
        x
        == "The tumourExtends upto the capsule of the liver but not through it. The surgical resection margin is uninvolved."
    ):
        updates["Others"] = "NR"
    elif x == "3.8 cm, alive: 46":
        updates["Tumor Size"] = "3.8 cm"
        updates["FollowUp_Result"] = "alive"
        updates["Others"] = "NR"
    elif x == "Status [KRAS]:WT /MUT [KRAS]:WT":
        updates["Molecular Information"] = "KRAS:WT"
        updates["Others"] = "NR"
    elif x in ["B cell , alive: 74B cell, death: 7"]:
        updates["Histological Diagnosis (Original)"] = "B-Cell Lymphoma"
        updates["FollowUp_Result"] = "alive" if "alive" in x else "death"
        updates["Others"] = "NR"
    elif x in ["Right hemilectomy", "Left hemilectomy"]:
        updates["Anatomical Sub-Location"] = (
            "Colon (Right)" if x == "Right hemilectomy" else "Colon (Left)"
        )
        updates["Anatomical Location"] = "Colon"
        updates["Others"] = "NR"
    elif x == "Anterior rectum resection":
        updates["Anatomical Location"] = "Rectum"
        updates["Anatomical Sub-Location"] = "Rectum (Anterior Resection)"
        updates["Others"] = "NR"
    elif x == "Sigmoidectomy":
        updates["Anatomical Location"] = "Colon"
        updates["Anatomical Sub-Location"] = "Colon (Sigmoid)"
        updates["Others"] = "NR"
    elif x in [
        "Resection - left selective neck dissection levels I-IV",
        "Resection - left selective neck dissection levels I-III",
    ]:
        updates["Anatomical Location"] = "Lymph Nodes"
        updates["Anatomical Sub-Location"] = (
            "Lymph Nodes (Left Neck - Levels I-IV)"
            if "I-IV" in x
            else "Lymph Nodes (Left Neck - Levels I-III)"
        )
        updates["Others"] = "NR"
    elif (
        x
        == "Resection, left subtotal glossectomy with adjacent left floor of mouth, anterior rim mandibulectomy including alveolar mucosa, and left selective neck dissection, levels I-IV"
    ):
        updates["Anatomical Location"] = "Oral Cavity"
        updates["Anatomical Sub-Location"] = (
            "Tongue (Significant Part Of Left), Mouth (Left Floor), Mandible (Anterior Rim And Alveolar Mucosa), And Lymph Nodes (Left Neck - Levels I-IV)"
        )
        updates["Others"] = "NR"
    elif (
        x
        == "Resection, left hemi-glossectomy, left floor of mouth,    left lingual rim mandibulectomy, with attached, left selective neck dissection, levels I-IV"
    ):
        updates["Anatomical Location"] = "Oral Cavity"
        updates["Anatomical Sub-Location"] = (
            "Tongue (Left Half), Mouth (Left Floor), Mandible (Left Lingual Rim), And Lymph Nodes (Left Neck - Levels I-IV)"
        )
        updates["Others"] = "NR"
    elif x == "Left oropharynx ":
        updates["Anatomical Location"] = "Oropharynx"
        updates["Anatomical Sub-Location"] = "Oropharynx (Left)"
        updates["Others"] = "NR"
    elif x == "1.8 cm, alive: 31":
        updates["Tumor Size"] = "1.8 cm"
        updates["FollowUp_Result"] = "alive"
        updates["Others"] = "NR"
    elif x == "2.5 cm, death: 59":
        updates["Tumor Size"] = "2.5 cm"
        updates["FollowUp_Result"] = "death"
        updates["Others"] = "NR"
    elif x == "Brain Tumor, 1p36+, 19q13-, MGMT-, ":
        updates["Molecular Information"] = "1p36+, 19q13-, MGMT-"
        updates["Others"] = "NR"
    elif x == "Left oropharynx (tonsil)":
        updates["Anatomical Location"] = "Tonsil"
        updates["Anatomical Sub-Location"] = "Tonsil (Left)"
        updates["Others"] = "NR"
    elif x in ["B cell , alive: 74B cell, death: 7"]:
        updates["Histological Diagnosis (Original)"] = "B-Cell Lymphoma"
        updates["FollowUp_Result"] = "alive" if "alive" in x else "death"
        updates["Others"] = "NR"
    elif x == "N-myc not amplified":
        updates["Molecular Information"] = "N-MYC:Not Amplified"
        updates["Others"] = "NR"
    elif x == "N-myc amplified":
        updates["Molecular Information"] = "N-MYC:Amplified"
        updates["Others"] = "NR"
    elif x == "Gleason 10":
        updates["Grade"] = "Gleason 10"
        updates["Others"] = "NR"
    elif x == "ALK+, 3+":
        updates["Molecular Information"] = "ALK+ 3+"
        updates["Others"] = "NR"
    elif x in ["Upper Lobectomy", "Upper lobectomy"]:
        updates["Anatomical Sub-Location"] = "Lung (Upper Lobe)"
        updates["Anatomical Location"] = "Lung"
        updates["Others"] = "NR"
    elif x == "S2 right lung resection":
        updates["Anatomical Sub-Location"] = "Lung (Right Upper Lobe - Segment 2)"
        updates["Anatomical Location"] = "Lung"
        updates["Others"] = "NR"
    elif x in ["2018", "2017"]:
        updates["Year_Collection"] = x
        updates["Others"] = "NR"
    elif x in [
        "Upper right lobectomy",
        "Upper left lobectomy",
        "Right upper lobectomy",
    ]:
        updates["Anatomical Sub-Location"] = (
            "Lung (Right Upper Lobe)"
            if "right" in x.lower()
            else "Lung (Left Upper Lobe)"
        )
        updates["Anatomical Location"] = "Lung"
        updates["Others"] = "NR"
    elif x == "Resection":
        updates["Anatomical Sub-Location"] = (
            row["Anatomical Sub-Location"] + " (Resection)"
            if row["Anatomical Sub-Location"] != "NR"
            else "Resection"
        )
        updates["Others"] = "NR"
    elif x == "Pre treatment":
        updates["Histological Subtype"] = (
            updates["Histological Subtype"] + " (Pre-Treatment)"
        )
        updates["Others"] = "NR"
    elif x in ["Primary", "primary"]:
        updates["Specimen Type"] = "P"
        updates["Others"] = "NR"
    elif x == "biopsy":
        updates["Anatomical Sub-Location"] = (
            row["Anatomical Sub-Location"] + " (Biopsy)"
            if row["Anatomical Sub-Location"] != "NR"
            else "Biopsy"
        )
        updates["Others"] = "NR"
    elif x == "Pneumectomy":
        updates["Anatomical Location"] = "Lung"
        updates["Others"] = "NR"
    elif x in [
        "CHOP; prednisolone (prior to and after chemo); GCSF ",
        "mHSPC",
        "DexamethasoneRituximab; R-Bendamustine; PrednisolonePrednisolone",
        "CHOP",
        "Rituximab; R-CHOP; Prednisolone; Dexamethasone post-chemo",
        "Chloroambucil and dexamethasone",
        "Cyclophosphamide and dexamethasone;\nVinblastine; Prednisolone",
        "IVE; BEAM; GCSF; methotrexate; autologous stem cell transplant ",
        "Rituximab; R-Bendamustine; Prednisolone",
        "Dexamethasone",
        "Prednisolone",
        "TKI",
    ]:
        combined_inventory["Treatmnent_Information"] = x
        updates["Others"] = "NR"
    elif x == "peritoneum":
        updates["Anatomical Location"] = "Peritoneum"
        updates["Others"] = "NR"
    elif x == "pleura":
        updates["Anatomical Location"] = "Pleura"
        updates["Others"] = "NR"
    elif x == "GBM":
        updates["Histological Diagnosis (Original)"] = (
            "Glioblastoma"
            if row["Histological Diagnosis (Original)"] == "NR"
            else row["Histological Diagnosis (Original)"]
        )
        updates["Others"] = "NR"
    elif x in ["mCRPC ", "mCRPC"]:
        updates["Histological Subtype"] = (
            row["Histological Diagnosis (Original)"] + " (mCRPC)"
        )
        updates["Others"] = "NR"

    # Set to 'NR' if it's one of the processed-but-unclassified Head & Neck terms
    if updates["Others"] != x and x in HN_others_col:
        updates["Others"] = "NR"

    # Return the Series of updated values.
    # We explicitly set 'Others' to 'NR' if it was processed, otherwise keep original 'x'
    updates["Others"] = (
        "NR"
        if updates["Others"] != x
        or x in ["Not stated", "GENERAL", "HNSCC", "Immunotherapy"]
        else x
    )

    return pd.Series(updates)


# Fill the NA values in the Others column with 'NR'
combined_inventory.fillna({"Others": "NR"}, inplace=True)

# Apply the function to the 'Others' column (using the merged function)
cols_to_update = [
    "Molecular Information",
    "Grade",
    "Paired Sample ID",
    "Paired Sample",
    "Anatomical Location",
    "Anatomical Sub-Location",
    "Specimen Type",
    "Histological Diagnosis (Original)",
    "Histological Subtype",
    "Pairing Site / Primary Tumour Site",
    "Others",
    "Tissue Source",
    "Tumor Size",
    "Age",
    "FollowUp_Result",
    "IHC_Molecular_Data",
    "Chemo_Targeted_Tx",
    "Year_Collection",
    "T",
    "N",
    "M",
]

# When the Sample Type has a value, move that to the Specimen Type column
for idx, row in combined_inventory.iterrows():
    sample_type = row["Sample Type"]
    specimen_type = row["Specimen Type"]
    if sample_type == "Mets":
        combined_inventory.at[idx, "Specimen Type"] = "M"
    elif sample_type == "Primary":
        combined_inventory.at[idx, "Specimen Type"] = "P"
    elif sample_type == "Normal":
        combined_inventory.at[idx, "Specimen Type"] = "N"

# Apply the single merged function
updated_cols_df = combined_inventory.apply(process_inventory_row, axis=1)

# Assign back the updated columns
combined_inventory[cols_to_update] = updated_cols_df[cols_to_update]

# Remove unnecessary columns
combined_inventory = combined_inventory.drop(
    columns=[
        "Unnamed: 66",
        "　",
        "Sample Type",
        "RNA_Scope",
        "Specimen_Histological_Diagnosis",
        "Sub_Type",
        "Project",
        "Qty_Verified",
        "Unusable",
        "Current_Inventory",
        "Shipped",
        "Discarded",
        "Outgoing_TRK_from_SBC",
        "Manifest",
        "Remark",
        "TA_Number1",
        "Core_ID_1",
        "TA_Number2",
        "Core_ID_2",
        "TA_Number3",
        "Core_ID_3",
        "TA_Number4",
        "Core ID_4",
        "TA_Number5",
        "Core_ID_5",
        "TA_Number6",
        "Core_ID_6",
        "TA_Number7",
        "Core_ID_7",
        "TA_Number8",
        "Core_ID_8",
        "TA_Number9",
        "Core_ID_9",
        "TA_Number10",
        "Core_ID_10",
        "TA_Number11",
        "Core_ID_11",
        "TA_Number12",
        "Core_ID_12",
        "TA_Number13",
        "Core_ID_13",
    ]
)

# Rename the columns
column_name_map = {
    "Source data field": "Source_Data_Field",
    "Tissue Source": "Tissue_Source",
    "Anatomical Location": "Anatomical_Location",
    "Anatomical Sub-Location": "Anatomical_Sub_Location",
    "Specimen Type": "Specimen_Type",
    "Histological Diagnosis (Original)": "Histological_Diagnosis_Original",
    "Histological Subtype": "Histological_Subtype",
    "Pairing Site / Primary Tumour Site": "Pairing_Site_Primary_Tumour_Site",
    "Molecular Information": "Molecular_Information",
    "Paired Sample ID": "Paired_Sample_ID",
    "Paired Sample": "Paired_Sample",
    "Tumor Size": "Tumor_Size",
}
combined_inventory = combined_inventory.rename(columns=column_name_map)

# When the Mutation_Data has a value, move that to the Molecular_Information Column
for idx, row in combined_inventory.iterrows():
    mutation_data = row["Mutation_Data"]
    molecular_info = row["Molecular_Information"]
    if mutation_data not in ["　", "", "NC", "Not specifed", "Not Specified", "NR"]:
        if molecular_info == "NR":
            combined_inventory.at[idx, "Molecular_Information"] = mutation_data
            combined_inventory.at[idx, "Mutation_Data"] = "NR"

# Clean the IHC_Molecular_Data column
for idx, row in combined_inventory.iterrows():
    ihc_data = row["IHC_Molecular_Data"]
    if ihc_data in [
        "7710 3117 5594",
        "　",
        "",
        "Not stated",
        "NR",
        "Normal (non-malignant)",
    ]:
        combined_inventory.at[idx, "IHC_Molecular_Data"] = "NR"
    elif (
        ihc_data
        == "Immunohistochemistry.   shows positivity for CK7 and focal positive staining for TTF-1 and CK20.  There is immunonegativity for CDX-2."
    ):
        combined_inventory.at[idx, "IHC_Molecular_Data"] = (
            "CK7+ Focal CK20+ TTF-1+ CDX-2-"
        )
    elif (
        ihc_data
        == "Immunohistochemistry has been performed on the mediastinal tissue (Specimen B) and the tumour shows diffuse positivity for CK7 and focal positivity for CK20 and CDX2 together with positivity for TTF-1."
    ):
        combined_inventory.at[idx, "IHC_Molecular_Data"] = "CK7+ CK20+ CDX2+ TTF-1+"


# re-order the columns
cols_order = [
    "Donor_ID",
    "Sample_ID",
    "Original_ID",
    "Block_Number",
    "Block_Dimensions",
    "Source_File",
    "Source_Data_Field",
    "Condition_Status",
    "Origin",
    "Ethnicity",
    "Gender",
    "Age",
    "T",
    "N",
    "M",
    "Stage",
    "Grade",
    "Date_Diagnosis",
    "Year_Diagnosis",
    "Year_Collection",
    "Month_Diagnosis",
    "Tissue_Source",
    "Anatomical_Location",
    "Anatomical_Sub_Location",
    "Specimen_Type",
    "Histological_Diagnosis_Original",
    "Histological_Subtype",
    "Paired_Sample",
    "Pairing_Site_Primary_Tumour_Site",
    "Paired_Sample_ID",
    "Tumor_Size",
    "Molecular_Information",
    "Mutation_Data",
    "IHC_Molecular_Data",
    "Chemo_Targeted_Tx",
    "IO_Other_Targeted_Tx",
    "FollowUp_Result",
    "FollowUp_Months",
    "Others",
]
combined_inventory = combined_inventory[cols_order]

# remove spaces from the ID columns
cols_to_trim = ["Donor_ID", "Sample_ID", "Original_ID"]
for col in cols_to_trim:
    combined_inventory[col] = combined_inventory[col].str.replace(
        r"\s+", "", regex=True
    )

# In[]: Export the File
combined_inventory.to_excel(
    "C:/Users/medaamm/Box/Data file (AA)/Other Files/combined_inventory.xlsx",
    index=False,
)
combined_inventory_unique_combos.to_excel(
    "C:/Users/medaamm/Box/Data file (AA)/Other Files/combined_inventory_unique_combos.xlsx",
    index=False,
)

# %%
