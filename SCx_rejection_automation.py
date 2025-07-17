import os
from dotenv import load_dotenv
import numpy as np
import pandas as pd
import urllib.parse
from sqlalchemy import create_engine
import pyodbc
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File

# Load environment variables
load_dotenv(override=True)

# Set up Connection String
server = "S-HQ-DATAFARM"
database = "DataAnalytics"
username = "insight"
pwd = os.getenv('DMPWD')
driver = "SQL Server Native Client 11.0"

connection_string_sql = 'mssql+pyodbc://insight:%s@S-HQ-DATAFARM/DataAnalytics?driver=SQL Server Native Client 11.0' % urllib.parse.quote_plus(f"{pwd}")
engine_msql = create_engine(connection_string_sql)
conn_msql = engine_msql.raw_connection()
cursor_msql = conn_msql.cursor()

# SharePoint site and folder details
sharepoint_site = "https://armholdco.sharepoint.com/sites/DataInsightsTeam"
folder_url = "/sites/DataInsightsTeam/Shared Documents/CX Automation/FailedSubRedDump"

# Authentication details
username = os.getenv('SUNME')
password = os.getenv('SPWD')

# Authenticate and create a client context
try:
    credentials = UserCredential(username, password)
    ctx = ClientContext(sharepoint_site).with_credentials(credentials)

    # Get the folder
    folder = ctx.web.get_folder_by_server_relative_url(folder_url)
    files = folder.files
    ctx.load(files)
    ctx.execute_query()

    # Find the newest file
    newest_file = None
    for file in files:
        if newest_file is None or file.time_last_modified > newest_file.time_last_modified:
            newest_file = file

    if newest_file:
        # Download the newest file
        response = File.open_binary(ctx, newest_file.serverRelativeUrl)

        # Save the file locally
        local_filename = newest_file.name
        with open(local_filename, "wb") as local_file:
            local_file.write(response.content)

        # Read the Excel file into a DataFrame
        df = pd.read_excel(local_filename, engine='openpyxl')

        # Process the DataFrame
        df['Client Id'] = df['Client Id'].astype('string')
        df['Client Id'] = df['Client Id'].str.replace('.0', '', regex=False)
        df_id = df['Client Id']
        df = df[df['Client Id'].str.match(r'^\d')]
        client_ids = tuple(df['Client Id'].astype(str))

        # Format the client_ids tuple into a string suitable for SQL IN clause
        formatted_client_ids = ', '.join(f"'{id}'" for id in client_ids)

        # Corrected SQL query
        QUERY = f"""
        SELECT accountnumber, name, armone_accountname, emailaddress1, telephone1
        FROM new_crm_account NOLOCK
        WHERE accountnumber IN ({formatted_client_ids})
        """

        df1 = pd.read_sql(QUERY, conn_msql)
        merged_df = pd.merge(df, df1, left_on='Client Id', right_on='accountnumber', how='inner')

        # Save the merged DataFrame to an Excel file
        lfilename = 'update_' + local_filename 
        merged_df.to_excel(lfilename, index=False)

        # SharePoint site and folder details for upload
        target_folder_url = "/sites/DataInsightsTeam/Shared Documents/CX Automation/FailedSubRedOutput"

        # Upload the file to SharePoint
        with open(lfilename, "rb") as file_content:
            target_folder = ctx.web.get_folder_by_server_relative_url(target_folder_url)
            target_file = target_folder.upload_file(lfilename, file_content.read()).execute_query()

        print(f"File uploaded to {target_file.serverRelativeUrl}")
    else:
        print("No files found in the specified folder.")
except ImportError as e:
    print("Required module not found:", e)
except Exception as e:
    print("An error occurred:", e)
