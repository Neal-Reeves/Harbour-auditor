
# Harbour Auditor

## Description

This script automatically compares access logs (formatted as an HTML file) with an Excel format (*.xlsx*) workbook tracker stored on the UCL SharePoint. The script extracts user details from the HTML logs via scraping (using BeautifulSoup) and also downloads the tracker content from the workbook using the Office365-REST-Python-Client package. Because the tracker contains sensitive data, the script uses the io.BytesIO module. This is an in-memory binary buffer and ensures that rather than being written to a temporary file, the parsed and classified aut outputs are written to CSV, but the tracker content is not.

The script four CSV files:
1. approved_users.csv -- 
2. left_ucl.csv --
3. expired_projects.csv --
4. ineligible_users.csv -- 

In addition, a fifth file named unparsed_dates.csv is generated containing user details for those users for whom a project end date is present, the project end date cannot be converted to a datetime object and the project end date is not "Open" or "open". This file will only be generated where relevant and will not be output if no matching records are found.

Files are written to ./Outputs/audit_tables. These folders will be created if they do not already exist.


## Getting Started

### Dependencies

### Installation

### How to Run



## Authors

## Version History

## License