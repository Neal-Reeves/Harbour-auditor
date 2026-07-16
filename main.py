import pandas as pd
from bs4 import BeautifulSoup

def open_tracker(tracker_url):
    return pd.read_excel(tracker_url)

def all_of_us_parser(html_file_path):
    #Parses All of Us HTML report to extract active users
    with open(html_file_path, "r", encoding="utf-8") as logfile:
        soup = BeautifulSoup(f.read(), "html.parser")
    
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