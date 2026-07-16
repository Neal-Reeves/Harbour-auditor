import pandas as pd
from bs4 import BeautifulSoup
from dataclasses import dataclass
from typing import Optional
from datetime import date

#Column Details
NAME_COLUMN = 1
EMAIL_COLUMN = 4
PROJECT_END_COLUMN = 11
SUPERVISOR_NAME_COLUMN = 3
SUPERVISOR_EMAIL_COLUMN = None

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
    project_end_date: Optional[date] = None
    supervisor_name: str
    supervisor_email: str

@dataclass
class AuditOutput:
    name: str
    email: str
    status: str = "Active" #Options: Active, Left UCL, Project Expired, Ineligible
    supervisor_name: Optional[str] = None
    supervisor_email: Optional[str] = None

def open_tracker(tracker_url):
    return pd.read_excel(tracker_url)

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
            name = row[NAME_COLUMN],
            email = row[EMAIL_COLUMN],
            project_end_date = row[PROJECT_END_COLUMN],
            supervisor_name = row[SUPERVISOR_NAME_COLUMN],
            supervisor_email = row[SUPERVISOR_EMAIL_COLUMN]
        )
        tracker_users.append(row_obj)
    
    return pd.DataFrame(tracker_users)

