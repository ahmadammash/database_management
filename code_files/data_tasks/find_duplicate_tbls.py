# Import the required libraries
import json

# Import required functions
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

# load the cohort mapping from the json file
cohort_names_mapping_path = project_root / "json_files" / "cohort_names_mapping.json"
with open(cohort_names_mapping_path, "r") as f:
    cohort_mapping = json.load(f)


# Find any duplicate table names regardless of the schema name and return them as a list
def find_duplicate_tbls():
    """
    This function queries the database to identify any tables whose names are duplicated - even if present in different
    schemas.
    """
    # Query to get all table names from the database
    query = "SELECT TABLE_SCHEMA, TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE';"

    # Execute the query and fetch the results
    result = cursor.execute(query)
    tables = result.fetchall()

    # Create a dictionary to count occurrences of each table name
    table_count = {}
    for schema, table in tables:
        if table in table_count:
            table_count[table] += 1
        else:
            table_count[table] = 1

    # Find duplicates (table names with count greater than 1)
    duplicates = [table for table, count in table_count.items() if count > 1]

    return duplicates


# delete all tables from user generated schemas
def delete_user_generated_schemas():
    """
    This function deletes all tables from any user-generated schema.
    """
    # Query to get all user-generated schemas (schemas that are not system schemas)
    query = """
    SELECT schema_name
    FROM information_schema.schemata
    WHERE schema_name NOT IN ('dbo', 'sys', 'imp_documents', 'transferred_files', 'INFORMATION_SCHEMA', 'guest', 'db_owner', 'db_accessadmin', 'db_securityadmin', 'db_ddladmin', 'db_backupoperator', 'db_datareader', 'db_datawriter', 'db_denydatareader', 'db_denydatawriter');
    """

    # Execute the query and fetch the results
    result = cursor.execute(query)
    user_schemas = result.fetchall()

    # Loop through each user-generated schema and drop all tables within it
    for (schema,) in user_schemas:
        # Query to get all tables in the current schema
        table_query = f"SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = '{schema}' AND TABLE_TYPE = 'BASE TABLE';"
        table_result = cursor.execute(table_query)
        tables_in_schema = table_result.fetchall()

        # Drop each table in the current schema
        for (table,) in tables_in_schema:
            drop_query = f"DROP TABLE [{schema}].[{table}];"
            try:
                cursor.execute(drop_query)
                print(f"Dropped table: {schema}.{table}")
            except Exception as e:
                print(f"Failed to drop table: {schema}.{table}. Error: {e}")

    # Commit the changes to the database
    cnxn.commit()
