import os
import io
import requests
import argparse
import time
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, asdict
from datetime import date
from typing import Optional

from bs4 import BeautifulSoup
from dotenv import load_dotenv
from office365.sharepoint.client_context import ClientContext
import pandas as pd

load_dotenv()

#Column Details
NAME_COLUMN = 1
EMAIL_COLUMN = 4
PROJECT_END_COLUMN = 11
SUPERVISOR_NAME_COLUMN = 3
SUPERVISOR_EMAIL_COLUMN = None

#Determine file path for output CSVs. Default is to write to ./outputs/audit_tables.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs", "audit_tables")

DOMAIN_NAME = os.environ.get("UCL_DOMAIN_NAME")
SITE_URL = os.environ.get("SHAREPOINT_SITE_URL")
CLIENT_ID = os.environ.get("APP_CLIENT_ID")
FILE_URL = os.environ.get("TRACKER_FILE_URL")
PERSON_API_URL = os.environ.get("PERSON_API_URL")
PERSON_API_CLIENT_SECRET = os.environ.get("PERSON_API_CLIENT_SECRET")
PERSON_API_TARGET_CLIENT_ID = os.environ.get("PERSON_API_TARGET_CLIENT_ID")

_token_cache = {}
session = requests.Session()

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
    status: str = "Active" #Options: Active, Left UCL, Project Expired, Ineligible, Flag
    supervisor_name: Optional[str] = None

def open_tracker() -> pd.ExcelFile:
    #Attempts to stream tracker Workbook into memory
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
        raise RuntimeError(
            f"Failed to fetch tracker from SharePoint: {SITE_URL}."
            f"Check .env values and ensure completion of sign-in prompt. Error details: {e}."
        )

def all_of_us_parser(html_file_path: str) -> pd.DataFrame:
    #Parses All of Us HTML report to extract active users
    response = requests.get(html_file_path)
    response .raise_for_status()

    soup = BeautifulSoup(response.text, "lxml")

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

def extract_tracker_users(tracker) -> pd.DataFrame:
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

def compare_user_lists(source_table, tracker_table) -> pd.DataFrame:
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
    #Convert end dates to datetime type. Invalid values (e.g., strings) will be null.
    audit_frame["project_end_date"] = pd.to_datetime(audit_frame["project_end_date"], errors="coerce")

    return audit_frame


def generate_audit_outputs(audit_df, employment_map) -> pd.DataFrame:
    today = pd.to_datetime("today")
    outputs = []

    for index, row in audit_df.iterrows():
        name = row.get("name_x") if pd.notna(row.get("name_x")) else row.get("name_y")

        email = row.get("email")

        raw_date = row.get("project_end_date_raw")

        has_unparsed_date = (
            pd.isna(row.get("project_end_date")) and 
            pd.notna(raw_date) and
            str(raw_date).strip().lower() != "open"
        )

        has_missing_email = pd.isna(email)

        if has_missing_email or has_unparsed_date: 
            status = "Flag"
        else:
            is_active = employment_map.get(email, False)
            merge_status = row.get("_merge")

            if merge_status == "left_only":
                status = "Ineligible"
            elif not is_active:
                status = "Left UCL"
            elif pd.notna(row.get("project_end_date")) and row["project_end_date"] < today:
                status = "Project Expired"
            else:
                status = "Approved"

        outputs.append(AuditOutput(
            name=name,
            email=email,
            status=status,
            supervisor_name=row.get("supervisor_name")
        ))

    return pd.DataFrame([asdict(o) for o in outputs])

def get_api_token():
    cached = _token_cache.get("token")
    if cached and cached["expires_at"] > time.time() + 30:
        return cached["access_token"]

    token_url = f"https://login.microsoftonline.com/{DOMAIN_NAME}/oauth2/v2.0/token"

    data = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": PERSON_API_CLIENT_SECRET,
        "scope": f"{PERSON_API_TARGET_CLIENT_ID}/.default"
    }

    response = requests.post(token_url, data=data)
    response.raise_for_status()
    payload = response.json()

    _token_cache["token"] = {
        "access_token": payload["access_token"],
        "expires_at": time.time() + payload["expires_in"]
    }
    return payload["access_token"]

def request_affiliation(session, token, user_email):
    request_url = PERSON_API_URL + f"person?email={user_email}"
    response = session.get(request_url, headers={"Authorization": f"Bearer {token}"})

    response.raise_for_status()
    
    return response.json()

def still_at_ucl(user_json: dict) -> bool:
    person_collection = (user_json or {}).get("person_collection", [])

    associations = []
    for property in person_collection:
        associations.extend(property.get("association") or [])

    if not associations: 
        return False

    return any(assoc.get("currency") != 3 for assoc in associations)

def fetch_employment_status(email: str, token) -> str | bool:
    if not email or pd.isna(email):
        return (email, False)
    
    try:
        data = request_affiliation(session, token, email)
        return (email, still_at_ucl(data))
    except Exception:
        return (email, False)

def batch_check_ucl_status(emails: list) -> dict:
    token = get_api_token()
    # Adjust max_workers (10-20 is usually a safe sweet spot without hitting rate limits)
    with ThreadPoolExecutor(max_workers=10) as executor:
        futures = [executor.submit(fetch_employment_status, email, token) for email in emails]
        results = [f.result() for f in futures]

    return dict(results)

def run_audit(html_file: str):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    all_of_us_users = all_of_us_parser(html_file)
    tracker = open_tracker()
    tracker_users = extract_tracker_users(tracker)
    audit_frame = compare_user_lists(all_of_us_users, tracker_users)

    unique_emails = audit_frame["email"].dropna().unique().tolist()

    employment_map = batch_check_ucl_status(unique_emails)

    output_df = generate_audit_outputs(audit_frame, employment_map)

    status_for_filenames = {
        "Approved": "approved_users.csv",
        "Left UCL": "left_ucl.csv",
        "Project Expired": "expired_projects.csv",
        "Ineligible": "ineligible_users.csv",
        "Flag": "flagged_users.csv"
    }

    for status, filename in status_for_filenames.items():
        subset = output_df[output_df["status"] == status]
        file_path = os.path.join(OUTPUT_DIR, filename)
        subset.to_csv(file_path, index=False)
        print(f"Successfully generated report: {file_path} ({len(subset)} row(s))")

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