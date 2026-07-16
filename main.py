import pandas as pd
from bs4 import BeautifulSoup
from dataclasses import dataclass
from typing import Optional
from datetime import date

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
            name = cells[0].get_text(strip=True).strip("'\"")
            email = cells[1].get_text(strip=True).strip("'\"")
        
        users.append({
            "name": name,
            "email": email.lower()
        })

    return pd.DataFrame(users)