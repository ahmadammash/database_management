import json
import os
import urllib.parse

import google.auth
from dotenv import load_dotenv
from googleapiclient import discovery
from sqlalchemy import create_engine, text


def save_all_settings(updated_config, settings_file_path):
    try:
        with open(settings_file_path, "w") as f:
            json.dump(updated_config, f, indent=4)
    except Exception as e:
        import sys

        sys.stderr.write(
            f"\nCRITICAL ERROR: Failed to save the updated configuration to disk: {e}\n"
        )
        sys.exit(1)


def create_sql_engine():
    SERVER_NAME = "172.16.44.34"
    DATABASE_NAME = "WTBB_L4"
    connection_string = f"mssql+pyodbc://@{SERVER_NAME}/{DATABASE_NAME}?driver=ODBC+Driver+17+for+SQL+Server&trusted_connection=yes"
    engine = create_engine(connection_string)
    return engine


def create_gcloud_engine(path):

    load_dotenv(path)

    safe_password = urllib.parse.quote_plus(os.getenv("gcloud_password"))

    # Use the safe_password in the string
    connection_string = f"postgresql+psycopg2://{os.getenv('gcloud_username')}:{safe_password}@{os.getenv('gcloud_host')}:{os.getenv('gcloud_port')}/{os.getenv('gcloud_db')}"

    engine = create_engine(connection_string)

    # Test the connection
    try:
        with engine.connect() as conn:
            # text() is required in SQLAlchemy 2.0+
            result = conn.execute(text("SELECT current_user;"))
            print(f"Successfully connected as: {result.fetchone()[0]}")
    except Exception as e:
        print(f"Connection failed: {e}")
    return engine


def manage_sql_instance(action):

    # load the credentials from the .env file
    load_dotenv()

    # get the project id and instance name from the .env file
    project_id = os.getenv("gcloud_project")
    instance_name = os.getenv("gcloud_db")

    # get default credentials automatically
    credentials, project = google.auth.default()

    # Build the service using the modern credentials
    service = discovery.build("sqladmin", "v1beta4", credentials=credentials)

    policy = "ALWAYS" if action == "start" else "NEVER"
    db_instance_body = {"settings": {"activationPolicy": policy}}

    request = service.instances().patch(
        project=project_id, instance=instance_name, body=db_instance_body
    )
    response = request.execute()
    print(f"✅ Cloud SQL instance {instance_name} set to {action}.")
