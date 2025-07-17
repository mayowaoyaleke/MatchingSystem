# Matching System

This project automates the process of retrieving files from a SharePoint folder, processing the data, and uploading the results back to SharePoint. It also integrates with a SQL Server database for data enrichment.

## Features

- **SharePoint Integration**: Downloads files from a specified SharePoint folder and uploads processed results back to SharePoint.
- **SQL Server Integration**: Queries a SQL Server database to enrich data.
- **Data Processing**: Processes Excel files using pandas for data cleaning and transformation.

## Prerequisites

- Python 3.9 or higher
- Required Python libraries (install via `requirements.txt`):
  - `pandas`
  - `numpy`
  - `sqlalchemy`
  - `pyodbc`
  - `python-dotenv`
  - `office365-rest-python-client`
  - `openpyxl`
- Access to the SharePoint site and SQL Server database.

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/your-repo/matching-system.git
   cd matching-system

## Create .env
SUNME=your_sharepoint_username
SPWD=your_sharepoint_password
DMPWD=your_sql_server_password

## Run Code
python SCx_rejection_automation.py

Monitoring SharePoint Folder
The script automatically monitors the specified SharePoint folder for new files, processes them, and uploads the results.

## File Structure
SCx_rejection_automation.py: Main script for processing files.
Logging
Logs are saved to monitor_folder.log for debugging and tracking the script's execution.
