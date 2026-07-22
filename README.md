
# Harbour Auditor

## Description

This script automatically compares access logs (formatted as an HTML file) with an Excel format (*.xlsx*) workbook tracker stored on the UCL SharePoint. The script extracts user details from the HTML logs via scraping (using BeautifulSoup) and also downloads the tracker content from the workbook using the Office365-REST-Python-Client package. Because the tracker contains sensitive data, the script uses the io.BytesIO module. This is an in-memory binary buffer and ensures that rather than being written to a temporary file, the parsed and classified aut outputs are written to CSV, but the tracker content is not.

The script four CSV files:
1. approved_users.csv -- Users who appear in both the tracker workbook and the HTML access logs. These are approved users for whom no action is required.
2. left_ucl.csv -- Users who appear in the tracker but who do not appear in the HTML logs. These users should be assumed to have left UCL and may need their access permissions changed.
3. expired_projects.csv -- Users who appear in the HTML access logs and who appear in the tracker but where the recorded project expiry date has passed. These projects should be assumed to have concluded and the users may need their access permissions changed.
4. ineligible_users.csv -- Users who appear in the HTML access logs and who do not appear in the tracker. These users may be assumed to be ineligible and their access permissions should be updated.

In addition, a fifth file named unparsed_dates.csv is generated containing user details for those users for whom a project end date is present, the project end date cannot be converted to a datetime object and the project end date is not "Open" or "open". This file will only be generated where relevant and will not be output if no matching records are found.

Files are written to ./Outputs/audit_tables. These folders will be created if they do not already exist.


## Getting Started

### Dependencies

### Installation

### How to Run



## Authors

## Version History

## License