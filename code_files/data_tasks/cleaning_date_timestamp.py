import pandas as pd

# Load your file
df = pd.read_excel(
    "C:/Users/medaamm/Box/Data file (AA)/Miscellaneous Requests/TA4525 & TA4526.xlsx",
    sheet_name="Sheet1",
    header=0,
)

# Identify all date columns containing timestamp patterns
date_cols = [
    "Date_Of_Surgery",
    "Specimen_1_Collection_Date",
    "Date_Of_Histological_Diagnosis",
    "Other_Relapse_1",
    "Other_Relapse_2",
    "Other_Relapse_3",
    "Radiation_1_Start_Date",
    "Radiation_1_End_Date",
    "Radiation_2_Start_Date",
    "Adjuvant_ChemoTx_1_Start_Date",
    "Adjuvant_ChemoTx_1_End_Date",
    "Metastatic_Chemo_Targeted_Tx_1_Start_Date",
    "Metastatic_Chemo_Targeted_Tx_1_End_Date",
    "Metastatic_Chemo_Targeted_Tx_2_Start_Date",
    "Date_Last_Seen",
    "Death_Date",
]

# Strip 'T00:00:00Z' and convert to pure YYYY-MM-DD
for col in date_cols:
    if col in df.columns:
        df[col] = (
            df[col]
            .astype(str)
            .apply(
                lambda x: (
                    x.split("T")[0]
                    if "T" in str(x) and x != "nan"
                    else ("" if x == "nan" else x)
                )
            )
        )

# Save to a new Excel file
# df.to_excel('Book1_Cleaned.xlsx', index=False)
