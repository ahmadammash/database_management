# Import the required libraries
import json
import re
from pathlib import Path

# Import the required functions
from code_files.automation.update_cohort_mapping import clean_logic
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

# Define the path to the file path to cohort mapping json file
filepath_cohort_mapping_path = (
    project_root / "json_files" / "filepath_cohortname_mapping.json"
)

# Define the path to the cohort names mapping json file
cohort_names_mapping_path = project_root / "json_files" / "cohort_names_mapping.json"

# Define the file path to the Box files
box_followup = Path(
    "C:/Users/AmmashA.IA/Box/Follow up data (Leeds)/Updated follow up data sheets"
)
box_prelim = Path(
    "C:/Users/AmmashA.IA/Box/Preliminary data files (Leeds)/Updated prelim data sheets"
)
all_box_files = list(box_followup.glob("*.xlsx")) + list(box_prelim.glob("*.xlsx"))


# Define a function to update the filepath to cohort name mappings
def update_filepath_cohortname_mappings():
    """
    This function retrieves all the file names from the box files, cleans them, and then maps them to the cohort name
    by extracting the cohort name from them. The function accounts for specific file names to ensure that there are
    no wrongly mapped file names.
    """
    # Load the cohort names mapping json file
    with open(cohort_names_mapping_path, "r") as f:
        cohort_names_mapping = json.load(f)

    # Sort cohorts by length descending so longer/more specific names match first
    sorted_cohorts = sorted(set(cohort_names_mapping.values()), key=len, reverse=True)

    # create the empty filepath to cohort name mapping dictionary
    filepath_cohortname_dict = {}

    # Loop through all the box files
    for file in all_box_files:
        # create the pattern for the file name (prelim and follow up)
        pattern = r"(L00.*?)_(Follow|Prelim)"
        # Search for the pattern in each file name within the loop
        match = re.search(pattern, str(file))
        # Skip the file if it does not match the pattern
        if not match:
            print(f"ℹ️ Skipping the file {file}. Match not found")
            continue
        # If a match is found, extract the first group which is the file name
        file_name = match.group(1)
        # Clean it by removing extra spaces or double underscores
        file_name = re.sub(r"_\s+", "_", file_name)
        file_name = file_name.replace("__", "_")

        # Isolate segment after bracketed ID
        raw_segment = file_name.split(")_")[-1]

        # Use clean_logic directly to guarantee identical string cleaning
        normalized_temp = clean_logic(raw_segment)

        # Additional plural handling for cholangiocarcinoma
        normalized_temp = normalized_temp.replace(
            "cholangiocarcinomas", "cholangiocarcinoma"
        )

        # Additional spelling correction for sarcoma
        normalized_temp = normalized_temp.replace("sarcorma", "sarcoma")

        # replace the "ileumcolon" with "colonileum"
        normalized_temp = normalized_temp.replace("ileumcolon", "colonileum")

        # replace "brain1xxxs1" with brain1s1xxxx in normalized_temp - e.g., brain1breasts1 with brain1s1breast
        normalized_temp = re.sub(r"brain([0-9])([a-z]+)s([0-9])", r"brain1s\1\2", normalized_temp)


        # Define the matched boolean variable as false by default - this is to map the file name to the cohort name
        matched = False

        # Loop through the sorted cohort names retrieved from the cohort mapping dict values
        for cohort in sorted_cohorts:
            # clean the cohort name to ensure there are no errors
            normalized_cohort = clean_logic(cohort).replace(
                "cholangiocarcinomas", "cholangiocarcinoma"
            )
            # define the pattern for the SCLC cohorts
            pattern = r"^sclc(\d.*)"
            # search for the SCLC cohorts pattern in the current cohort name
            match = re.search(pattern, normalized_temp)
            # if there is a match and if the current cohort is SCLC, map the file name to ths 1st group in the pattern
            # which is the number of the cohort and the shipment number
            if match:
                filepath_cohortname_dict[file_name] = f"sclc{match.group(1)}"
                # Set the matched variable to True and exit the loop
                matched = True
                break
            # More conditions for other cohorts to ensure that the file name is mapped to the correct cohort name
            elif (
                "hnio1" in normalized_temp
                and "pembro" not in normalized_temp
                and "nivo" not in normalized_temp
            ):
                filepath_cohortname_dict[file_name] = "hnio1"
                matched = True
                break
            elif "crc4s10" in normalized_temp:
                filepath_cohortname_dict[file_name] = "crc4s10"
                matched = True
                break
            elif "crc4s11" in normalized_temp:
                filepath_cohortname_dict[file_name] = "crc4s11"
                matched = True
                break
            elif 'crc4s1' in normalized_temp and 'crc4s10' not in normalized_temp:
                filepath_cohortname_dict[file_name] = 'crc4s1'
                matched = True
                break
            elif (
                "malignant" in normalized_temp and "nonmalignant" not in normalized_temp
            ):
                filepath_cohortname_dict[file_name] = (
                    "biopsiesandresectionsbiliarytractmalignant"
                )
                matched = True
                break
            elif "rcc2" in normalized_temp and "5kio" not in normalized_temp:
                filepath_cohortname_dict[file_name] = "rcc2"
                matched = True
                break
            elif "nsclc3adc" in normalized_temp and "met" not in normalized_temp:
                filepath_cohortname_dict[file_name] = "nsclc3adc"
                matched = True
                break
            elif "nsclc3scc" in normalized_temp and "met" not in normalized_temp:
                filepath_cohortname_dict[file_name] = "nsclc3scc"
                matched = True
                break
            elif (
                normalized_cohort in normalized_temp
                or normalized_temp in normalized_cohort
            ):
                filepath_cohortname_dict[file_name] = cohort
                matched = True
                break
        # In the case where no match was found, print out a statement
        if not matched:
            print(
                f"⚠️ No matching cohort found for: {file_name} (temp: {normalized_temp})"
            )
    # Dump the changes into the filepath to cohort name mapping file
    with open(filepath_cohort_mapping_path, "w") as f:
        json.dump(filepath_cohortname_dict, f, indent=4)
