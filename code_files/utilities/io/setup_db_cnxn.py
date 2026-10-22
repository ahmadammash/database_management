# Import required libraries
from pathlib import Path

import pyodbc

# Import required functions (Works globally after pip install -e .)
from code_files.utilities.io.database import create_engine, create_gcloud_engine


def setup_cnxn():
    # Anchor project_root directly to this file's position on disk
    current_file = Path(__file__).resolve()
    
    # Walks up parent directories from this file until it finds 'database_management' or 'pyproject.toml'
    project_root = next(
        (p for p in current_file.parents if (p / "pyproject.toml").exists() or p.name == "database_management"),
        current_file.parents[3]
    )

    cnxn = None
    cursor = None
    engine = None

    # Determine connection type based on path or environment
    if "AmmashA.IA" in str(current_file):
        SERVER_NAME = "172.16.44.34"
        DATABASE_NAME = "WTBB_L4"
        CONNECTION_STRING = (
            f"DRIVER={{ODBC Driver 17 for SQL Server}};"
            f"SERVER={SERVER_NAME};"
            f"DATABASE={DATABASE_NAME};"
            f"Trusted_Connection=yes;"
        )
        try:
            cnxn = pyodbc.connect(CONNECTION_STRING)
            cursor = cnxn.cursor()
            
            # Create SQLAlchemy connection engine
            connection_string = (
                f"mssql+pyodbc://@{SERVER_NAME}/{DATABASE_NAME}?"
                f"driver=ODBC+Driver+17+for+SQL+Server&trusted_connection=yes"
            )
            engine = create_engine(connection_string)

        except pyodbc.Error as ex:
            print(f"Connection failed: {ex}")

    elif "medaamm" in str(current_file):
        env_path = project_root / "code_files" / "utilities" / "io" / ".env"
        engine = create_gcloud_engine(env_path)

    return project_root, cnxn, cursor, engine, current_file