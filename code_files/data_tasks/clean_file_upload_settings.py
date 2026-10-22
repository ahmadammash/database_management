import json
import re

import pandas as pd

# Import the required libraries
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

# define the paths to the json files
file_path = project_root / "json_files" / "file_upload_settings.json"
filepath_cohortname_mapping_path = (
    project_root / "json_files" / "filepath_cohortname_mapping.json"
)
cohort_mapping_path = project_root / "json_files" / "cohort_names_mapping.json"

with open(file_path, "r") as f:
    file_upload_settings = json.load(f)

with open(filepath_cohortname_mapping_path, "r") as f:
    filepath_cohortname_mapping = json.load(f)

with open(cohort_mapping_path, "r") as f:
    cohort_mapping = json.load(f)

cohorts = []
for cohort, config in file_upload_settings.items():
    cohorts.append(cohort)
cohorts = sorted(cohorts)


all_tables_query = """
SELECT table_schema, table_name
FROM information_schema.tables
WHERE table_schema NOT IN ('information_schema', 'pg_catalog', 'guest')
                 AND table_schema NOT LIKE '%sys%';
"""

all_tables = pd.read_sql(all_tables_query, cnxn)

for tbl, schema in zip(all_tables["table_name"], all_tables["table_schema"]):
    if tbl not in cohorts:
        print(tbl)

# strip out the part before the table type suffix and replace it with the cohort name from the mapping and replace
# the sql table name in the config to then run the batch re-upload
# work on a copy of the dictionary
# for table, config in list(file_upload_settings.items()):
#     pattern_1 = r'(.*)(_FollowUp_Anon|_Supplementary_Anon|_Prelim_Anon)'
#     pattern_2 = r'.*(/|\\)(L0.*)_(Prelim.*|Follow.*|Supplementary.*)'

#     match_1 = re.search(pattern_1, table)
#     match_2 = re.search(pattern_2, config['local_file_path'])

#     if match_1 and match_2:
#         original_cohort = (
#             match_2.group(2)
#             .replace('__', '_')
#             .replace('_ ', '_')
#         )

#         mapped_cohort = filepath_cohortname_mapping.get(original_cohort)

#         if mapped_cohort:
#             new_key = mapped_cohort + match_1.group(2)

#             file_upload_settings[new_key] = file_upload_settings.pop(table)
for table, config in list(file_upload_settings.items()):
    pattern_1 = r"(.*)(_FollowUp_Unanon|_Supplementary_Unanon|_Prelim_Unanon)"
    pattern_2 = r".*(\\|/)(.*?)(Leeds [Ff]ollow up.*|[Ff]ollow up.*)"
    match_1 = re.search(pattern_1, table)
    match_2 = re.search(pattern_2, config["local_file_path"])
    if match_1 and not match_2:
        print(config["local_file_path"])
    if match_1 and match_2:
        original_cohort = match_2.group(2).strip()
        if "Normal" in original_cohort and "tissue" not in original_cohort:
            original_cohort = original_cohort.replace("Normal", "Normal tissue")
        mapped_cohort = cohort_mapping.get(original_cohort)
        if mapped_cohort is None:
            print(f"Mapping not found for {table} with {original_cohort}")
        else:
            file_upload_settings[table]["sql_table_name"] = (
                mapped_cohort + match_1.group(2)
            )
            new_key = mapped_cohort + match_1.group(2)

            file_upload_settings[new_key] = file_upload_settings.pop(table)


# Save entire updated config
with open(file_path, "w") as f:
    json.dump(file_upload_settings, f, indent=4)

# fetch_all_tables = """
# SELECT table_schema, table_name
# FROM information_schema.tables
# WHERE table_schema NOT IN ('information_schema', 'pg_catalog', 'guest')
#                  AND table_schema NOT LIKE '%sys%'
#                  AND table_schema NOT LIKE 'imp_documents'
# """
# all_tables = pd.read_sql(fetch_all_tables, cnxn)

# for idx, row in all_tables.iterrows():
#     schema = row['table_schema']
#     table = row['table_name']
#     del_query = f"""DROP TABLE [{schema}].[{table}]"""
#     cursor.execute(del_query)
# cnxn.commit()
