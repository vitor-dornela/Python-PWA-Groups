import os
import re
import sys
import json
import psutil
import platform
import logging
from html import unescape

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

# User Constants
FILE_NAME = "pwa_data"
OUTPUT_DIRECTORY = "~/Downloads"

# Program Constants
LOGIN_URL = "https://login.microsoftonline.com/"
GROUP_CONTAINER_ID = "GridDataRow"
USER_CONTAINER_ID = "ctl00_ctl00_PlaceHolderMain_PWA_PlaceHolderMain_idFormSectionUsers_ctl02_idSwpUsers_BetaList_Container"
CATEGORY_CONTAINER_ID = "ctl00_ctl00_PlaceHolderMain_PWA_PlaceHolderMain_idFormSectionCategories_ctl02_idSwpCategories_BetaList_Container"
CHROME_TIMEOUT = 10

sys.stdout.reconfigure(encoding='utf-8')


def validate_pwa_url(url: str) -> bool:
    pattern = r"^https://[a-zA-Z0-9-]+\.sharepoint\.com/sites/[a-zA-Z0-9_-]+/$"
    return re.match(pattern, url)


def close_chrome():
    for process in psutil.process_iter(attrs=["pid", "name"]):
        if "chrome" in process.info["name"].lower():
            try:
                os.kill(process.info["pid"], 9)
            except psutil.AccessDenied:
                logging.warning("Access denied to process %s. Skipping.", process.info["pid"])
            except Exception as e:
                logging.error("Could not close process %s: %s", process.info["pid"], e)
    logging.info("Attempted to close all running Chrome instances.")


def get_output_file() -> str:
    output_path = os.path.expanduser(OUTPUT_DIRECTORY)
    file_counter = 1
    output_file = os.path.join(output_path, f"{FILE_NAME}.xlsx")
    while os.path.exists(output_file):
        output_file = os.path.join(output_path, f"{FILE_NAME}_{file_counter}.xlsx")
        file_counter += 1
    return output_file


def get_pwa_instance_url() -> str:
    while True:
        url = input("Enter the PWA instance URL: ").strip().rstrip("/") + "/"
        if validate_pwa_url(url):
            return url
        print("Invalid URL format. The URL should follow this pattern: https://TENANT_NAME.sharepoint.com/sites/PWA_SITE/")


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
            # Sorted by profile name
            profile_mapping = {v["name"]: k for k, v in sorted(profiles.items(), key=lambda x: x[1]["name"])}
    else:
        return None, None

    return chrome_path, profile_mapping


def select_chrome_profile(chrome_path, profile_mapping, options):
    use_guest = input("Use guest profile? (options: yes; no [select a profile]): ").strip().lower()
    if use_guest == "yes":
        options.add_argument("--guest")
    else:
        if not chrome_path or not profile_mapping:
            raise Exception("Could not detect Chrome profiles. Ensure Chrome is installed and used before running this script.")

        logging.info("Available Chrome Profiles:")
        for index, (profile_name, folder) in enumerate(profile_mapping.items(), 1):
            logging.info("[%s] %s", index, profile_name)

        selected_index = int(input("Enter the number corresponding to the profile you want to use: ")) - 1
        profile_name, profile_folder = list(profile_mapping.items())[selected_index]
        options.add_argument(f"--user-data-dir={chrome_path}")
        options.add_argument(f"--profile-directory={profile_folder}")
    return options


def wait_for_element(driver, by, identifier, timeout=CHROME_TIMEOUT):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, identifier)))


def extract_groups(soup, group_edit_page: str) -> list:
    groups = []
    grid_rows = soup.find_all("tr", id=GROUP_CONTAINER_ID)
    for row in grid_rows:
        columns = row.find_all("td")
        if len(columns) >= 5:
            group_name = columns[1].get_text(strip=True)
            group_description = columns[2].get_text(strip=True)
            ad_group = columns[3].get_text(strip=True)
            last_synced = columns[4].get_text(strip=True)
            group_uid = row.get("rowid")
            if group_uid:
                full_group_url = f"{group_edit_page}{group_uid}"
                groups.append({
                    "Group UID": group_uid,
                    "Group Name": group_name,
                    "Group Description": group_description,
                    "AD Group": ad_group,
                    "Last Synchronized": last_synced,
                    "URL": full_group_url
                })
    return groups


def extract_details_from_group(driver, group: dict, users: list, categories: list):
    driver.get(group["URL"])
    logging.info("Accessing group page: %s", group["Group Name"])
    try:
        wait_for_element(driver, By.ID, USER_CONTAINER_ID, timeout=10)
        group_soup = BeautifulSoup(driver.page_source, "html.parser")
        user_container = group_soup.find("td", id=USER_CONTAINER_ID)
        category_container = group_soup.find("td", id=CATEGORY_CONTAINER_ID)

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

        if category_container:
            for option in category_container.find_all("option", value=True):
                category_uid = option["value"].strip()
                category_name = option.text.strip()
                categories.append({
                    "Category UID": category_uid,
                    "Category Name": category_name,
                    "Group UID": group["Group UID"],
                    "Group Name": group["Group Name"]
                })
    except Exception as e:
        logging.error("Failed to extract details for group: %s (%s)", group["Group Name"], e)


def save_to_excel(groups: list, users: list, categories: list, output_file: str):
    df_groups = pd.DataFrame(groups)
    df_users = pd.DataFrame(users)
    df_categories = pd.DataFrame(categories).sort_values(by="Category Name")

    with pd.ExcelWriter(output_file) as writer:        
        df_users.to_excel(writer, sheet_name="Users", index=False)
        df_groups.to_excel(writer, sheet_name="Groups", index=False)
        df_categories.to_excel(writer, sheet_name="Categories", index=False)
    logging.info("Data extraction complete! Saved to %s", output_file)


def main():
    # Close running Chrome instances
    close_chrome()

    # Prepare URLs and output file
    pwa_instance_url = get_pwa_instance_url()
    group_edit_page = f"{pwa_instance_url}_layouts/15/PWA/Admin/AddModifyGroup.aspx?groupUid="
    groups_page = f"{pwa_instance_url}_layouts/15/PWA/Admin/ManageGroups.aspx"
    output_file = get_output_file()

    # Setup Selenium options
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    chrome_path, profile_mapping = get_chrome_profiles()
    options = select_chrome_profile(chrome_path, profile_mapping, options)

    driver = webdriver.Chrome(options=options)
    try:
        # Login
        driver.get(LOGIN_URL)
        logging.info("Please complete the login process in the browser window...")
        WebDriverWait(driver, 180).until(EC.url_changes(LOGIN_URL))

        # Navigate to Groups page
        driver.get(groups_page)
        logging.info("Navigated to group management page. Waiting for the page to load...")
        try:
            wait_for_element(driver, By.ID, GROUP_CONTAINER_ID, timeout=20)
            logging.info("Page is ready.")
        except Exception:
            logging.error("Timed out waiting for page to load.")
            return

        soup = BeautifulSoup(driver.page_source, "html.parser")
        groups = extract_groups(soup, group_edit_page)
        users = []
        categories = []

        for group in groups:
            extract_details_from_group(driver, group, users, categories)

        save_to_excel(groups, users, categories, output_file)
    finally:
        driver.quit()
        if platform.system() == "Windows":
            os.system("pause")


if __name__ == "__main__":
    main()
