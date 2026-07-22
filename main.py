import os
import io
import argparse
from dotenv import load_dotenv
import pandas as pd
from bs4 import BeautifulSoup
from dataclasses import dataclass
from typing import Optional
from datetime import date
from office365.sharepoint.client_context import ClientContext

load_dotenv()

#Column Details
NAME_COLUMN = 1
EMAIL_COLUMN = 4
PROJECT_END_COLUMN = 11
SUPERVISOR_NAME_COLUMN = 3
SUPERVISOR_EMAIL_COLUMN = None

SOURCE_URL = "../Test.html"
TRACKER_URL = ""
OUTPUT_DIR = "/outputs/audit_tables"

DOMAIN_NAME = os.environ.get("UCL_DOMAIN_NAME")
SITE_URL = os.environ.get("SHAREPOINT_SITE_URL")
CLIENT_ID = os.environ.get("SHAREPOINT_CLIENT_ID")
FILE_URL = os.environ.get("TRACKER_FILE_URL")

@dataclass
class PortalUser:
    name: str
    email: str
    active_workspaces: int = 0
    controlled_tier_access: bool = False

@dataclass
class TrackerUser:
    name: str
    email: str
    supervisor_name: str
    supervisor_email: Optional[str] = None
    project_end_date: Optional[date] = None


@dataclass
class AuditOutput:
    name: str
    email: str
    status: str = "Active" #Options: Active, Left UCL, Project Expired, Ineligible
    supervisor_name: Optional[str] = None
    supervisor_email: Optional[str] = None

def open_tracker():
    print("Initialising connection to SharePoint")

    client = ClientContext(SITE_URL).with_interactive(
        tenant= DOMAIN_NAME,
        client_id = CLIENT_ID
    )
    print("Streaming tracker data to memory...")

    file_content = io.BytesIO()
    client.web.get_file_by_server_relative_url(FILE_URL).download(file_content).execute_query()
    file_content.seek(0)

    return pd.read_excel(file_content)

def all_of_us_parser(html_file_path):
    #Parses All of Us HTML report to extract active users
    with open(html_file_path, "r", encoding="utf-8") as logfile:
        soup = BeautifulSoup(logfile.read(), "html.parser")
    
    users = []

    #Obtain all table rows
    rows = soup.find_all("tr")

    #Iterate to find rows with user email
    for row in rows:
        cells = row.find_all("td")
        if len(cells) >= 4:
            user_obj = PortalUser(
                name = cells[0].get_text(strip=True).strip("'\""),
                email = cells[1].get_text(strip=True).strip("'\""),
                active_workspaces=int(cells[2].get_text(strip=True)) if cells[2].get_text(strip=True).isdigit() else 0,
                controlled_tier_access=cells[3].get_text(strip=True) == "X"
            )
        
            users.append(user_obj)

    return pd.DataFrame(users)

def extract_tracker_users(tracker):
    tracker_users = []
    for row in tracker.iterrows():
        row_obj = TrackerUser(
            name = row.iloc[NAME_COLUMN],
            email = row.iloc[EMAIL_COLUMN],
            project_end_date = row.iloc[PROJECT_END_COLUMN],
            supervisor_name = row.iloc[SUPERVISOR_NAME_COLUMN],
            supervisor_email = row.iloc[SUPERVISOR_EMAIL_COLUMN]
        )
        tracker_users.append(row_obj)
    
    return pd.DataFrame(tracker_users)

def compare_user_lists(source_table, tracker_table):
    source_table["email"] = source_table["email"].str.strip().lower()
    tracker_table["email"] = tracker_table["email"].str.strip().lower()

    audit_frame = pd.merge(
        source_table,
        tracker_table,
        on="email",
        how="outer",
        indicator=True
    )

    approved = audit_frame[audit_frame["_merge"] == "both"]
    no_longer_at_ucl = audit_frame[audit_frame["_merge"] == "right_only"]
    
    today = pd.to_datetime("today")
    audit_frame["project_end_date"] = pd.to_datetime(audit_frame["project_end_date"])

    access_not_needed = audit_frame[
        (audit_frame["_merge"] == "both") &
        (audit_frame["project_end_date"] < today)
    ]

    return approved, no_longer_at_ucl, access_not_needed


def run_audit(html_file):

    all_of_us_users = all_of_us_parser(html_file)
    tracker_users = open_tracker()

    approved, no_longer_at_ucl, access_not_needed = compare_user_lists(all_of_us_users, tracker_users)

    audit_reports = {
        "approved_users.csv": approved,
        "left_UCL.csv": no_longer_at_ucl,
        "expired_projects.csv": access_not_needed
    }

    for filename, dataframe in audi_reports.items():
        export_df = df.drop(columns=["_merge"])
        file_path = os.path.join(OUTPUT_DIR, filename)
        export_df.to_csv(file_path, index=False)
        print(f"successfully generated report: {file_path}")



def parse_args():
    parser = argparse.ArgumentParser(
        description = "Compare the All of Us Researcher Workbench access report against the UCL tracker."
    )

    return parser.parse_args()

if __name__ == "__main__":
    args = parse_args()
    run_audit(SOURCE_URL)