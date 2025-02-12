from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd
import os
import psutil
import json
import platform
from html import unescape
import sys

# === CONSTANTS ===
# PWA_INSTANCE_URL = "https://mrvengenhariasa.sharepoint.com/sites/PlanejamentoDIPRD/"
PWA_INSTANCE_URL = input("Enter the PWA instance URL: ").strip().rstrip("/") + "/"  # Ensure URL ends with a trailing slash
FILE_NAME = "pwa_users_groups"
OUTPUT_PATH = f"~/Downloads/{FILE_NAME}.xlsx"

# URLs
LOGIN_URL = "https://login.microsoftonline.com/"
GROUPS_PAGE = f"{PWA_INSTANCE_URL}_layouts/15/PWA/Admin/ManageGroups.aspx"
GROUP_EDIT_PAGE = f"{PWA_INSTANCE_URL}_layouts/15/PWA/Admin/AddModifyGroup.aspx?groupUid="

# Element Identifiers
GROUP_CONTAINER_ID = "GridDataRow"
USER_CONTAINER_ID = "ctl00_ctl00_PlaceHolderMain_PWA_PlaceHolderMain_idFormSectionUsers_ctl02_idSwpUsers_BetaList_Container"

sys.stdout.reconfigure(encoding='utf-8')

def close_chrome():
    for process in psutil.process_iter(attrs=['pid', 'name']):
        if "chrome" in process.info['name'].lower():
            try:
                os.kill(process.info['pid'], 9)
            except psutil.AccessDenied:
                print(f"Access denied to process {process.info['pid']}. Skipping.")
            except Exception as e:
                print(f"Could not close process {process.info['pid']}: {e}")
    print("Attempted to close all running Chrome instances.")

close_chrome()

def get_chrome_profiles():
    system = platform.system()
    if system == "Windows":
        chrome_path = os.path.expanduser("~\\AppData\\Local\\Google\\Chrome\\User Data")
        local_state_path = os.path.join(chrome_path, "Local State")
    else:
        return None, None
    
    if os.path.exists(local_state_path):
        with open(local_state_path, "r", encoding="utf-8") as f:
            data = json.load(f)
            profiles = data.get("profile", {}).get("info_cache", {})
            profile_mapping = {v["name"]: k for k, v in sorted(profiles.items(), key=lambda x: x[1]["name"])}
    else:
        return None, None
    
    return chrome_path, profile_mapping

chrome_path, profile_mapping = get_chrome_profiles()
use_guest = input("Use guest profile? (options: yes; no [select a profile]): ").strip().lower()

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")

if use_guest == "yes":
    options.add_argument("--guest")
else:
    if not chrome_path or not profile_mapping:
        raise Exception("Could not detect Chrome profiles. Ensure Chrome is installed and used before running this script.")
    
    print("Available Chrome Profiles:")
    for index, (profile_name, folder) in enumerate(profile_mapping.items(), 1):
        print(f"[{index}] {profile_name}")
    
    selected_index = int(input("Enter the number corresponding to the profile you want to use: ")) - 1
    profile_name, profile_folder = list(profile_mapping.items())[selected_index]
    profile_path = os.path.join(chrome_path, profile_folder)
    
    options.add_argument(f"--user-data-dir={chrome_path}")
    options.add_argument(f"--profile-directory={profile_folder}")

driver = webdriver.Chrome(options=options)

# === LOGIN ===
driver.get(LOGIN_URL)
print("Please complete the login process in the browser window...")
WebDriverWait(driver, 120).until(EC.url_changes(LOGIN_URL))

# === NAVIGATE TO GROUP MANAGEMENT ===
driver.get(GROUPS_PAGE)
print("Navigated to Gerenciar Grupos. Waiting for page to fully load...")

try:
    element_present = EC.presence_of_element_located((By.ID, GROUP_CONTAINER_ID))
    WebDriverWait(driver, 20).until(element_present)
    print("Page is ready.")
except:
    print("Timed out waiting for page to load.")
    driver.quit()
    exit()

full_html_text = driver.page_source
soup = BeautifulSoup(full_html_text, "html.parser")

# === EXTRACT GROUP INFORMATION ===
groups = []
grid_rows = soup.find_all("tr", id=GROUP_CONTAINER_ID)
for row in grid_rows:
    columns = row.find_all("td")
    if len(columns) >= 5:
        group_name = columns[1].get_text(strip=True)
        group_description = columns[2].get_text(strip=True)
        ad_group = columns[3].get_text(strip=True)
        last_synced = columns[4].get_text(strip=True)
        group_uid = row.get('rowid')
        if group_uid:
            full_group_url = f"{GROUP_EDIT_PAGE}{group_uid}"
            groups.append({
                "Group UID": group_uid,
                "Group Name": group_name,
                "Group Description": group_description,
                "AD Group": ad_group,
                "Last Synchronized": last_synced,
                "URL": full_group_url
            })

if not groups:
    print("Error: No groups found.")
    driver.quit()
    exit()

print(f"Extracted {len(groups)} groups.")

# === EXTRACT USER INFORMATION FROM EACH GROUP ===
users = []
for group in groups:
    driver.get(group["URL"])
    print(f"Accessing group page: {group['Group Name']}")
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, USER_CONTAINER_ID))
        )
        group_html = driver.page_source
        group_soup = BeautifulSoup(group_html, "html.parser")
        user_container = group_soup.find("td", id=USER_CONTAINER_ID)
        if user_container:
            for option in user_container.find_all("option", value=True):
                user_uid = option["value"].strip()
                user_name = option.text.strip()
                users.append({
                    "Group UID": group["Group UID"],
                    "Group Name": group["Group Name"],
                    "User UID": user_uid,
                    "User Name": user_name
                })
    except:
        print(f"Failed to extract users for group: {group['Group Name']}")

df_users = pd.DataFrame(users)
df_users.to_excel(OUTPUT_PATH, index=False)

print(f"Data extraction complete! Saved to {OUTPUT_PATH}")
driver.quit()

os.system("pause")  # Windows-only command
