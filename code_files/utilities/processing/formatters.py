# Import the required libraries
import re
import numpy as np
import openpyxl
import pandas as pd

from itertools import combinations
from code_files.utilities.config.mappings import set_data_types
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

# Import the required functions
from code_files.utilities.io.json_manager import load_decisions, save_decisions

def process_date_columns(main_workbook, dataset_type):
    """
    This function processes the date columns in the unanonymised data before uploading the data to the 
    database without causing errors.
    """
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

# ----------------------------------------------------------------------------------------------------------------------------------------

def process_dec_columns(main_workbook, dataset_type):
    """
    This function processes the decimal/float columns in the data to replace values which will violate the data type restrictions
    """
    # get the data types
    data_types = set_data_types(main_workbook, dataset_type)

    # filter for float types
    all_dec_int_columns = [key for key, value in data_types.items() if 'float' in str(value).lower() or 'int' in str(value).lower()]
    
    # Keep the columns that exist in the current working sheet
    existing_colums = [col for col in all_dec_int_columns if col in main_workbook.columns]

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
        "N/A"
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

def process_bit_columns(main_workbook, dataset_type):
    """
    This function processes the BIT columns in the data and replaces the inaccurate Yes/No/NA values to prevent any violations
    to the data type restrictions.
    """
    # get the data types
    data_types = set_data_types(main_workbook, dataset_type)

    # get the bit type columns
    all_bit_columns = [key for key, value in data_types.items() if 'bit' in str(value).lower()]

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
        "Metastatic_Disease_At_DX_DistantMets_NonRegionalLNs"
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

# ----------------------------------------------------------------------------------------------------------------------------------------


def combine_and_drop(df, cols):
    """
    This function combines columns using backfill. The function keeps the name of the first column in the list
    and drops the remaining ones.
    """
    # backfill the values in the first column of the list cols[0] using the others and then drop the extra columns
    df[cols[0]] = df[cols].bfill(axis=1).iloc[:, 0]
    return df.drop(columns=cols[1:])

# ----------------------------------------------------------------------------------------------------------------------------------------

def resolve_conflict(df, primary, secondary):
    """
    This function resolves conflicts between two columns whereby they have conflicting and non-NA values.
    """
    # Create a mask (True/False) for each row to determine whether the values in the two columns are both Non-NA and different.
    mask = (
        df[primary].notna() &
        df[secondary].notna() &
        (df[primary] != df[secondary])
    )
    # In the rows where there is a conflict, set the values of the secondary column equal to the primary column
    df.loc[mask, secondary] = df.loc[mask, primary]
    return df

# ----------------------------------------------------------------------------------------------------------------------------------------

def _get_merge_choice(decisions, decision_key, col1, col2, prompt_label, allow_skip=True):
    """
    This function checks whether a decision key is inside the JSON file. 
    If it is, then the choice is retrieved and used to combine columns. Otherwise, the function creates the prompt to retrieve
    the choice from the user.
    """
    # If the pair of columns (decision_key) is in the JSON file, get the choice i.e. the value associated with that key
    if decision_key in decisions:
        choice = decisions[decision_key]
        # print a message to show that the choice was loaded and return that choice
        print(f"⚡ Loaded saved decision for {decision_key}: '{choice}'")
        return choice

    
    # If the pair of columns (decision_key) is not in the JSON file
    
    # Define the valid options
    valid_options = {'1', '2', 'skip'} if allow_skip else {'1', '2'}
    
    
    # Create the prompt for the "input" from the user
    prompt = f"{prompt_label}: Keep 1 ({col1}), 2 ({col2})" + (", or 'skip': " if allow_skip else ": ")

    
    # while True creates an infinite loop that keeps running until the input from the user is amongst the valid options
    while True:
        raw = input(prompt).strip().lower()
        # if the user enters a valid option, set the choice variable
        if raw in valid_options:
            choice = 'skip' if raw == 'skip' else (col1 if raw == '1' else col2)
            # add the new key, value pair to the dictionary and return the choice
            decisions[decision_key] = choice
            return choice
        # if the input from the user is invalid, print a message accordingly
        print(f"Invalid input '{raw}'. Choose from {valid_options}.")

# ----------------------------------------------------------------------------------------------------------------------------------------

def find_combinable_columns(df, cols, path):
    """
    This function checks for combinable columns by determining when they overlap and its safe to merge them directly. When there
    is a conflict between the columns, the function asks the user to decide whether to combine them based on one column's values
    or to skip the combination altogether. The function uses choices present in the JSON file when available and creates a user
    prompt otherwise.
    """

    # load the decisions dictionary from the JSON file
    decisions = load_decisions(path)

    # remove aany duplicate names in cols
    cols = list(dict.fromkeys(cols))

    # Loop through the combinations (2 per combination) of columns from the cols list
    for col1, col2 in combinations(cols, 2):
        # if any of col1 or col2 is not in the dataframe, skip the loop iteration
        if col1 not in df.columns or col2 not in df.columns:
            continue


        # Otherwise, create the decision_key variable
        decision_key = f"{col1}|{col2}"


        # Boolean variable both_filled for when the values in col1 and col2 are both non-NA
        both_filled = df[col1].notna() & df[col2].notna()


        # The number of rows where both col1 and col2 values are both non-NA
        n_both = both_filled.sum()


        # Case 1: Disjoint (never populated in the same row)
        # When the number of rows where both are non-NA is 0
        if n_both == 0:
            
            # get the number of rows where each of col1 and col2 is non-NA
            n1, n2 = df[col1].notna().sum(), df[col2].notna().sum()
            print(f"✅ Disjoint: '{col1}' ({n1} non-NA values) and '{col2}' ({n2} non-NA values) never overlap.")

            # print the first 10 non-NA values of each column when they overlap
            print(f"First 10 values of {col1}: {df.loc[df[col1].notna()][col1].head(10)}")
            print(f"First 10 values of {col2}: {df.loc[df[col2].notna()][col2].head(10)}")
            
            # get the primary column either from the JSON file (if it exists) or through a user prompt
            primary_col = _get_merge_choice(decisions, decision_key, col1, col2, "SAFE TO MERGE", allow_skip=True)

        # Case 2: Overlapping rows exist
        # When there are cases where both col1 and col2 have non-NA values
        else:
            # get the value from col1
            s1 = df.loc[both_filled, col1].astype(str).str.strip().str.lower()

            # get the value from col2
            s2 = df.loc[both_filled, col2].astype(str).str.strip().str.lower()
            
            # get the number of rows where col1 and col2 have the same value
            matches = (s1 == s2).sum()

            # subtract the number of matches from the number of rows where both columns have values to get the number of conflicts
            conflicts = n_both - matches

            # if col1 and col2 have the same values whenever both columns are non-NA, then they are safe to merge
            if conflicts == 0:
                print(f"✅ Identical: '{col1}' and '{col2}' agree on all {n_both} overlapping rows.")
                # get the primary column either from the JSON file or through user input
                primary_col = _get_merge_choice(decisions, decision_key, col1, col2, "SAFE TO MERGE", allow_skip=True)
            else:
                # otherwise, there is a conflict between the columns
                print(f"❌ {conflicts} conflicts between '{col1}' and '{col2}' (agree on {matches} rows):")
                # show the first 10 values of each column
                print(df.loc[both_filled & (s1 != s2), [col1, col2]].drop_duplicates().head(10))
                # get the primary column either from the JSON file or through the user input
                primary_col = _get_merge_choice(decisions, decision_key, col1, col2, "CONFLICT", allow_skip=True)


        # save the decisions made to the JSON file
        save_decisions(decisions, path)


        # if the user entered 'skip' in the prompt, the combination step should be skipped and the loop moves to the next iteration
        if primary_col == 'skip':
            print(f"⏭️ Skipped {col1} vs {col2}\n" + "*" * 80)
            continue

        # If a value was entered by the user, get the secondary_col based on the primary_col needed by the user
        secondary_col = col2 if primary_col == col1 else col1

        # resolve the conflict between the columns
        df = resolve_conflict(df, primary=primary_col, secondary=secondary_col)
        df = combine_and_drop(df, [primary_col, secondary_col])
        print(f"✅ Merged '{secondary_col}' into '{primary_col}'\n" + "*" * 80)

    return df


def correct_misplaced_values(df, col, strs_to_check):

    # Load the decisions file
    path = project_root / "json_files" / "decisions.json"
    decisions = load_decisions(path = path)


    print("=" * 100)
    print(f"The column '{col}' has the below unique values:")
    print(df[col].dropna().unique())
    print("=" * 100)

    pattern = "|".join(strs_to_check)
    mask = df[col].astype(str).str.contains(pattern, case=False, na=False)

    matched_vals = df.loc[mask, col].unique()
    unmatched_vals = df.loc[~mask, col].unique()
    print(f"Found {len(matched_vals)} matching values in '{col}':\n{matched_vals}")
    print("=" * 100)
    print(f"Found {len(unmatched_vals)} unmatching values in '{col}':\n{unmatched_vals}")
    print("=" * 100)

    if col in decisions.keys():
        replacement_column = decisions.get(col)
    
    else:
        replacement_column = input("Enter the name of the column to move the values to: ").strip()

        # Add the choice to the decisions json file
        decisions[f"{col}"] = replacement_column
        save_decisions(decisions, path)

    # 1. Store the exact row mask before modifying df
    swap_mask = (~mask) & df[col].notna() & (~df[col].astype(str).str.contains('NR', na=False))

    print("Before replacement:")
    print(df.loc[swap_mask, [col, replacement_column]])

    # 2. Single .loc assignment with .to_numpy() swaps both columns atomically in-place
    df.loc[swap_mask, [col, replacement_column]] = df.loc[swap_mask, [replacement_column, col]].to_numpy()

    print("=" * 100)
    print("After replacement:")
    print(df.loc[swap_mask, [col, replacement_column]])

    return df


# ----------------------------------------------------------------------------------------------------------------------------------------

def process_percentages(file_path, sheet_name, header_row):
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb[sheet_name]
    data = []
    for row in ws.iter_rows(values_only = False):
        row_values = []
        for cell in row:
            val = cell.value
            # Check if Excel formatted this specific cell as a percentage
            if isinstance(val, (int, float)) and cell.number_format and '%' in cell.number_format:
                # Convert any decimals back to a percentage by multiplying by a 100
                val = f"{val * 100:g}%" # :g strips unneccessary trailing zeroes
            row_values.append(val)
        data.append(row_values)

    df = pd.DataFrame(data[1:], columns=data[0])
    return df



