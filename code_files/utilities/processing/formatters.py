# Import the required libraries
import re

import numpy as np
import pandas as pd


def process_date_columns(main_workbook, dataset_type):
    for col in main_workbook.columns:
        col_name = str(col).strip()
        col_name_lower = col_name.lower()

        # Skip the processing for any TTE columns
        if "tte_" in col_name_lower:
            # Let process_dec_columns or standard numeric conversion handle this column instead
            continue

        # Process the date columns in the unanonymised data
        if "unanonymised" in dataset_type and (
            "date" in col_name_lower or "dob" in col_name_lower
        ):
            # Clean string placeholders to NaN
            main_workbook[col] = (
                main_workbook[col]
                .replace(["N/A", "NR", "", " ", "N/A (biopsied)", "N/R"], np.nan)
                .infer_objects(copy=False)
            )

            # Catch all variations of numeric 0 / 0.0 entries causing the bigint datetime2 clash
            if main_workbook[col].dtype in ["float64", "int64", "object"]:
                main_workbook[col] = (
                    main_workbook[col]
                    .replace({0: np.nan, "0": np.nan, "0.0": np.nan})
                    .infer_objects(copy=False)
                )

            # Coerce everything to datetime64[ns] (corrupt/unparseable elements turn into NaT)
            main_workbook[col] = pd.to_datetime(main_workbook[col], errors="coerce")

            # Convert pandas NaT/NaN objects into Python None so SQLAlchemy passes clean SQL NULLs
            main_workbook[col] = main_workbook[col].where(
                main_workbook[col].notnull(), None
            )
    return main_workbook


def process_dec_columns(main_workbook):
    # identify the TTE columns while excluding the date columns
    all_dec_columns = [
        col
        for col in main_workbook.columns
        if ("TTE_" in str(col) or "Leibovich" in str(col))
        and "date" not in str(col).lower()
        and "seen" not in str(col).lower()
    ]

    # Keep the columns that exist in the current working sheet
    existing_colums = [col for col in all_dec_columns if col in main_workbook.columns]

    # Common placeholders to clean to NaN
    na_values = [
        "/",
        " ",
        "Uncertain",
        "NR",
        "N/a",
        "Indeterminate",
        "NR ",
        "Unknown",
        "X",
        "x",
        "N/R",
        "Not specified",
        "N/A (biopsied)",
        "?",
        "N/A (see comments)",
        "No (see comments from clinical team)",
    ]

    # cast true numeric timeline metrics without introducing raw 0s
    for col in existing_colums:
        main_workbook[col] = (
            main_workbook[col].replace(na_values, np.nan).infer_objects(copy=False)
        )
        main_workbook[col] = pd.to_numeric(main_workbook[col], errors="coerce")

        # Convert pandas NaN values to standard Python None (maps to database NULL)
        main_workbook[col] = main_workbook[col].where(
            main_workbook[col].notnull(), None
        )

    return main_workbook


# ----------------------------------------------------------------------------------------------------------------------------------------

# In[24]: Function to process the BIT type columns

def process_bit_columns(main_workbook):
    # all the possible bit columns
    all_bit_columns = [
        "AnyImmunotherapy",
        "SpecIsRecurrence",
        "AnyNeoadjuvantTherapy",
        "DeathTumourRelated",
        "Event_RecurrenceProgression",
        "Event_CancerDeath",
        "Event_DeathAllCauses",
        "Event_RecurrenceCancerDeath",
        "ResectionMarginsClear",
        "Castrate_Resistant_Prostate_Cancer",
        "Spec_IO_Naive",
        "MatchedMet1_IO_Naive",
        "MatchedMet2_IO_Naive",
        "MatchedMet3_IO_Naive",
        "AllBlocksReturnedFromTriStar",
        "Recall",
        "SentToTristar",
        "ResectionMarginsClear_HepaticRecurrences",
        "Castrate_Resistant_Disease_At_Time_Sample_Surgery",
        "Met_Disease_at_DX_distMets_nonRegionalLN",
        "Met_Disease_At_Time_Sample_Surgery_distMets_nonRegionalLN",
    ]

    # keep the bit columns that exist in the dataset being worked with
    existing_columns = [
        col for col in all_bit_columns if col in main_workbook.columns
    ]

    if not existing_columns:
        return main_workbook

    # Define regex patterns (case-insensitive)
    # Catches all YES patterns
    yes_pattern = r"^\s*(yes|y|r0|1)\b"

    # Catches all NO patterns
    no_pattern = r"^\s*(no|n|r1|r2|0)\b"

    for col in existing_columns:
        # Convert series to string type and strip surrounding whitespace
        s = main_workbook[col].astype(str).str.strip()

        # Identify matches using case-insensitive regex
        is_yes = s.str.contains(yes_pattern, case=False, na=False, regex=True)
        is_no = s.str.contains(no_pattern, case=False, na=False, regex=True)

        # Vectorized assignment: maps Yes -> 1, No -> 0, everything else -> NaN
        main_workbook[col] = np.select([is_yes, is_no], [1, 0], default=np.nan)

    return main_workbook
# ----------------------------------------------------------------------------------------------------------------------------------------


def process_mutations(val, pattern, mapping):
    if pd.isna(val) or not isinstance(val, str):
        return val

    match = re.search(r"\((.*?)\)", val)
    if match:
        extracted_text = match.group(1)
        return pattern.sub(lambda m: mapping[m.group(0)], extracted_text)
    return val
