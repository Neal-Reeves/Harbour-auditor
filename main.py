import os
import io
import requests
import argparse
from dotenv import load_dotenv
import pandas as pd
from bs4 import BeautifulSoup
from dataclasses import dataclass, asdict
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
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs", "audit_tables")

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
    project_end_date: Optional[date] = None


@dataclass
class AuditOutput:
    name: str
    email: str
    status: str = "Active" #Options: Active, Left UCL, Project Expired, Ineligible
    supervisor_name: Optional[str] = None

def open_tracker():
    print("Initialising connection to SharePoint")
    try:

        client = ClientContext(SITE_URL).with_interactive(
            tenant= DOMAIN_NAME,
            client_id = CLIENT_ID
        )
        print("Streaming tracker data to memory...")

        file_content = io.BytesIO()
        client.web.get_file_by_server_relative_url(FILE_URL).download(file_content).execute_query()
        file_content.seek(0)

        return pd.read_excel(file_content)
    except Exception as e:
        raise RunTimeError(
            f"Failed to fetch tracker from SharePoint: {SITE_URL}."
            f"Check .env values and ensure completion of sign-in prompt. Error details: {e}."
        )

def all_of_us_parser(html_file_path):
    #Parses All of Us HTML report to extract active users
    response = requests.get(html_file_path)
    response .raise_for_status()
    soup = BeautifulSoup(response.text, "html.parser")
    
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
    for index, row in tracker.iterrows():
        row_obj = TrackerUser(
            name = row.iloc[NAME_COLUMN],
            email = row.iloc[EMAIL_COLUMN],
            project_end_date = row.iloc[PROJECT_END_COLUMN],
            supervisor_name = row.iloc[SUPERVISOR_NAME_COLUMN],
        )
        tracker_users.append(row_obj)
    
    return pd.DataFrame(tracker_users)

def compare_user_lists(source_table, tracker_table):
    source_table["email"] = source_table["email"].str.strip().str.lower()
    tracker_table["email"] = tracker_table["email"].str.strip().str.lower()

    audit_frame = pd.merge(
        source_table,
        tracker_table,
        on="email",
        how="outer",
        indicator=True
    )

    audit_frame["project_end_date_raw"] = audit_frame["project_end_date"]
    audit_frame["project_end_date"] = pd.to_datetime(audit_frame["project_end_date"], errors="coerce")

    unparsed = audit_frame[
        audit_frame["project_end_date"].isna() &
        audit_frame["project_end_date"].notna() &
        (audit_frame["project_end_date_raw"].astype(str).str.lower() != "open")
    ]
    return audit_frame, unparsed

def generate_audit_outputs(audit_df):
    today = pd.to_datetime("today")
    outputs = []

    for index, row in audit_df.iterrows():
        name = row.get("name_x") if pd.notna(row.get("name_x")) else row.get("name_y")

        if row["_merge"] == "left_only":
            status = "Ineligible"
        elif row["_merge"] == "right_only":
            status = "Left UCL"
        elif pd.notna(row["project_end_date"]) and row["project_end_date"] < today:
            status = "Project Expired"
        else:
            status = "Active"
        
        outputs.append(AuditOutput(
            name = name,
            email = row["email"],
            status = status,
            supervisor_name = row.get("supervisor_name")
        ))
        
        outputs_df = pd.DataFrame([asdict(o) for o in outputs])
        return outputs_df


def run_audit(html_file):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    all_of_us_users = all_of_us_parser(html_file)
    tracker = open_tracker()
    tracker_users = extract_tracker_users(tracker)
    audit_frame, unparsed = compare_user_lists(all_of_us_users, tracker_users)
    output_df = generate_audit_outputs((audit_frame))

    status_for_filenames = {
        "Active": "approved_users.csv",
        "Left UCL": "left_ucl.csv",
        "Project Expired": "expired_projects.csv",
        "Ineligible": "ineligible_users.csv"
    }

    for status, filename in status_for_filenames.items():
        subset = output_df[output_df["status"] == status]
        file_path = os.path.join(OUTPUT_DIR, filename)
        subset.to_csv(file_path, index=False)
        print(f"Successfully generated report: {file_path} ({len(subset)} row(s))")

    if not unparsed.empty:
        unparsed_path = os.path.join(OUTPUT_DIR, "unparsed_dates.csv")
        unparsed.to_csv(unparsed_path, index=False)
        print(f"Warning: {len(unparsed)} row(s) have an unparseable project end date. Check manually.")

def parse_args():
    parser = argparse.ArgumentParser(
        description = "Compare the All of Us Researcher Workbench access report against the UCL tracker."
    )
    parser.add_argument(
        "log_url",
        default = None,
        help = "URL for All of Us access logs."
    )

    parser.add_argument(
        "--dataset-name",
        default="All of Us",
        help = "Name of dataset for comparison."
    )

    return parser.parse_args()

if __name__ == "__main__":
    args = parse_args()
    run_audit(args.log_url)