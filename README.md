
# Harbour Auditor

## Description

This script automatically compares access logs (formatted as an HTML file) with an Excel format (*.xlsx*) workbook tracker stored on the UCL SharePoint. The script extracts user details from the HTML logs via scraping (using BeautifulSoup) and also downloads the tracker content from the workbook using the Office365-REST-Python-Client package. Because the tracker contains sensitive data, the script uses the io.BytesIO module. This is an in-memory binary buffer and ensures that rather than being written to a temporary file, the parsed and classified aut outputs are written to CSV, but the tracker content is not.

Matching is performed based on user email. It is therefore assumed that the lead applicant's email in the tracker will match their email in the access logs.  

The script four CSV files:
1. approved_users.csv -- Users who appear in both the tracker workbook and the HTML access logs. These are approved users for whom no action is required.
2. left_ucl.csv -- Users who appear in the tracker but who do not appear in the HTML logs. These users should be assumed to have left UCL and may need their access permissions changed.
3. expired_projects.csv -- Users who appear in the HTML access logs and who appear in the tracker but where the recorded project expiry date has passed. These projects should be assumed to have concluded and the users may need their access permissions changed.
4. ineligible_users.csv -- Users who appear in the HTML access logs and who do not appear in the tracker. These users may be assumed to be ineligible and their access permissions should be updated.

In addition, a fifth file named unparsed_dates.csv is generated containing user details for those users for whom a project end date is present, the project end date cannot be converted to a datetime object and the project end date is not "Open" or "open". This file will only be generated where relevant and will not be output if no matching records are found.

Files are written to ./Outputs/audit_tables. These folders will be created if they do not already exist.

## Getting Started

### Dependencies
- Python 3.12+ (Please note: this application is currently incompatible with Python 3.15 due to a segfault bug and should only be run from a stable Python release)
- Python packages: please see `requirements.txt`
- A `.env` file with the required SharePoint variables -- see Setup below.
- An Entra ID app registration with delegated SharePoint and Graph read access to the tracker site.
- A manually-saved copy of the All of Us HTML access report (currently not fetched automatically)

### Setup

1. Download or otherwise clone the repository from GitHub.
2. Install the required dependencies from requirements.txt: `pip install -r requirements.txt`
3. Initialise the environments file .env. This should contain four variables: 
- UCL_DOMAIN_NAME -- the UCL domain (tenant)
- SHAREPOINT_SITE_URL -- the SharePoint URL where the tracker is stored. This URL must point to the staff folder
- SHAREPOINT_CLIENT_ID -- an Entra ID registered application (client) ID. This ID needs SharePoint API permissions, specifically the AllSites.Read permission.
- TRACKER_FILE_URL -- the path for the tracker file which can be copied from SharePoint. This path should begin /sites/

### How to Run

1. Open a terminal.
2. Navigate to the `Harbour Auditor` folder.
3. Run `Python main.py`

## Authors

## Version History

## License