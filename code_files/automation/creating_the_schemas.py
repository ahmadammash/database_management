# Import the required libraries
import os

# Import the required functions
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

# Create a connection cursor
cursor = cnxn.cursor()

# Get the list of schemas to create
txt_path = os.path.join(current_dir, "..", "..", "schemas_list.txt")
with open(txt_path, "r") as f:
    schemas = [line.strip() for line in f.read().splitlines() if line.strip]

# Iterate over the list of schemas and create each one
for schema in schemas:
    cursor.execute(f"""
    IF NOT EXISTS (SELECT * FROM sys.schemas WHERE name = '{schema}')
    BEGIN
        EXEC('CREATE SCHEMA {schema}')
    END
    """)

# Commit the changes
cnxn.commit()

# Close the cursor and connection
cursor.close()
cnxn.close()
