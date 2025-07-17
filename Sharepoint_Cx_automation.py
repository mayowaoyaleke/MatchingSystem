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

def get_sharepoint_files():
    """Get all files from SharePoint folder with their creation timestamps"""
    try:
        # Authenticate using username and password
        ctx = ClientContext(SHAREPOINT_URL).with_credentials(UserCredential(USERNAME, PASSWORD))
        folder = ctx.web.get_folder_by_server_relative_url(FOLDER_URL)
        files = folder.files
        ctx.load(files)
        ctx.execute_query()
        
        # Extract file names and creation times
        file_info = []
        for file in files:
            file_info.append({
                'name': file.properties["Name"],
                'created': file.properties["TimeCreated"],
                'modified': file.properties["TimeLastModified"]
            })
        
        logging.info(f"Retrieved {len(file_info)} files from SharePoint.")
        return file_info
        
    except Exception as e:
        logging.error(f"Error retrieving files from SharePoint: {e}")
        return []

def check_for_recent_files(minutes_threshold=5):
    """Check for files created or modified within the last X minutes"""
    try:
        from datetime import datetime, timezone, timedelta
        
        current_files = get_sharepoint_files()
        if not current_files:
            return False
        
        # Calculate threshold time
        threshold_time = datetime.now(timezone.utc) - timedelta(minutes=minutes_threshold)
        
        recent_files = []
        for file_info in current_files:
            # Parse SharePoint datetime (ISO format)
            created_time = datetime.fromisoformat(file_info['created'].replace('Z', '+00:00'))
            modified_time = datetime.fromisoformat(file_info['modified'].replace('Z', '+00:00'))
            
            # Check if file was created or modified recently
            if created_time > threshold_time or modified_time > threshold_time:
                recent_files.append(file_info['name'])
        
        if recent_files:
            logging.info(f"Recent files detected (within {minutes_threshold} minutes): {recent_files}")
            return True
        
        return False
        
    except Exception as e:
        logging.error(f"Error checking for recent files: {e}")
        return False

def check_for_any_files():
    """Simple check - if any files exist in folder, run notebook"""
    try:
        current_files = get_sharepoint_files()
        if current_files:
            logging.info(f"Files detected in folder: {len(current_files)} files")
            return True
        return False
        
    except Exception as e:
        logging.error(f"Error checking for files: {e}")
        return False

def run_notebook(notebook_path, output_path):
    """Execute the notebook using papermill"""
    try:
        logging.info(f"Running notebook: {notebook_path}")
        pm.execute_notebook(notebook_path, output_path)
        logging.info(f"Notebook executed successfully: {output_path}")
        return True
    except Exception as e:
        logging.error(f"Error running notebook: {e}")
        return False

def monitor_folder_recent_files(check_interval=30, file_age_threshold=5):
    """
    Monitor folder for recently created/modified files
    
    Args:
        check_interval (int): Seconds between checks
        file_age_threshold (int): Minutes threshold for considering files as "recent"
    """
    logging.info(f"Monitoring SharePoint folder for files created/modified within {file_age_threshold} minutes...")
    
    while True:
        try:
            if check_for_recent_files(minutes_threshold=file_age_threshold):
                # Run the notebook when recent files are detected
                success = run_notebook("SCx_rejection_automation.ipynb", "output_notebook.ipynb")
                if success:
                    logging.info("Notebook execution completed successfully")
                else:
                    logging.error("Notebook execution failed")
            
            time.sleep(check_interval)
            
        except KeyboardInterrupt:
            logging.info("Monitoring stopped by user")
            break
        except Exception as e:
            logging.error(f"Unexpected error in monitoring loop: {e}")
            time.sleep(check_interval)

def monitor_folder_any_files(check_interval=30, cooldown_period=300):
    """
    Monitor folder and run notebook if any files exist
    
    Args:
        check_interval (int): Seconds between checks
        cooldown_period (int): Seconds to wait after processing before checking again
    """
    logging.info("Monitoring SharePoint folder for any files...")
    last_execution_time = 0
    
    while True:
        try:
            current_time = time.time()
            
            # Check if enough time has passed since last execution
            if current_time - last_execution_time > cooldown_period:
                if check_for_any_files():
                    logging.info("Files detected - running notebook...")
                    
                    # Run the notebook when files are detected
                    success = run_notebook("SCx_rejection_automation.ipynb", "output_notebook.ipynb")
                    
                    if success:
                        logging.info("Notebook execution completed successfully")
                        last_execution_time = current_time
                        logging.info(f"Entering cooldown period for {cooldown_period} seconds...")
                    else:
                        logging.error("Notebook execution failed - will retry on next check")
                else:
                    logging.info("No files detected in SharePoint folder")
            else:
                remaining_cooldown = cooldown_period - (current_time - last_execution_time)
                logging.info(f"Still in cooldown period - {remaining_cooldown:.0f} seconds remaining")
            
            time.sleep(check_interval)
            
        except KeyboardInterrupt:
            logging.info("Monitoring stopped by user")
            break
        except Exception as e:
            logging.error(f"Unexpected error in monitoring loop: {e}")
            time.sleep(check_interval)

def run_once_if_files_exist():
    """Run notebook once if files exist, then exit"""
    logging.info("Checking SharePoint folder for files (one-time check)...")
    
    if check_for_any_files():
        success = run_notebook("SCx_rejection_automation.ipynb", "output_notebook.ipynb")
        if success:
            logging.info("Notebook execution completed successfully")
        else:
            logging.error("Notebook execution failed")
    else:
        logging.info("No files found in SharePoint folder")

if __name__ == "__main__":
    # Choose one of the monitoring strategies:
    
    # Option 1: Monitor for recently created/modified files (within last 5 minutes)
    # monitor_folder_recent_files(check_interval=30, file_age_threshold=5)
    
    # Option 2: Monitor for any files and run notebook if files exist
    monitor_folder_any_files(check_interval=30, cooldown_period=300)
    
    # Option 3: Run once if files exist, then exit
    # run_once_if_files_exist()
    