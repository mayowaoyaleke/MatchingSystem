import time
import os
import logging
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
import papermill as pm
from dotenv import load_dotenv

# Load environment variables
load_dotenv(override=True)

# Configure logging
logging.basicConfig(filename='monitor_folder.log', level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Replace with your SharePoint site URL and folder details
SHAREPOINT_URL = "https://armholdco.sharepoint.com/sites/DataInsightsTeam"
FOLDER_URL = "/sites/DataInsightsTeam/Shared Documents/CX Automation/FailedSubRedDump"  # Path to the SharePoint folder
USERNAME = os.getenv('SUNME')
PASSWORD = os.getenv('SPWD')
SAVED_FILE_LIST = "saved_file_list.txt"  # A file to store the previously checked files

def get_sharepoint_files():
    try:
        # Authenticate using username and password
        ctx = ClientContext(SHAREPOINT_URL).with_credentials(UserCredential(USERNAME, PASSWORD))
        folder = ctx.web.get_folder_by_server_relative_url(FOLDER_URL)
        files = folder.files
        ctx.load(files)
        ctx.execute_query()

        # Extract file names
        file_names = [file.properties["Name"] for file in files]
        logging.info(f"Retrieved {len(file_names)} files from SharePoint.")
        return file_names
    except Exception as e:
        logging.error(f"Error retrieving files from SharePoint: {e}")
        return []

def read_saved_file_list():
    if os.path.exists(SAVED_FILE_LIST):
        with open(SAVED_FILE_LIST, "r") as f:
            return set(f.read().splitlines())
    return set()

def save_file_list(file_list):
    with open(SAVED_FILE_LIST, "w") as f:
        f.write("\n".join(file_list))

def run_notebook(notebook_path, output_path):
    try:
        logging.info(f"Running notebook: {notebook_path}")
        pm.execute_notebook(notebook_path, output_path)
        logging.info(f"Notebook executed successfully: {output_path}")
    except Exception as e:
        logging.error(f"Error running notebook: {e}")

def monitor_folder():
    previously_checked_files = read_saved_file_list()
    logging.info("Monitoring SharePoint folder for new files...")

    while True:
        current_files = set(get_sharepoint_files())

        # Find new files that are not in previously checked files
        new_files = current_files - previously_checked_files

        if new_files:
            logging.info(f"New files detected: {new_files}")
            previously_checked_files.update(new_files)
            save_file_list(previously_checked_files)

            # Run the notebook whenever a new file is detected
            run_notebook("SCx_rejection_automation.ipynb", "output_notebook.ipynb")
        
        time.sleep(30)  # Wait for a minute before checking again

if __name__ == "__main__":
    monitor_folder()
