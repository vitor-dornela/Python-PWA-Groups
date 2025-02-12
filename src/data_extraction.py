import logging
from bs4 import BeautifulSoup
from .config import GROUP_CONTAINER_ID, USER_CONTAINER_ID, CATEGORY_CONTAINER_ID

def extract_groups(soup: BeautifulSoup, group_edit_page: str) -> list:
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
                groups.append({
                    "Group UID": group_uid,
                    "Group Name": group_name,
                    "Group Description": group_description,
                    "AD Group": ad_group,
                    "Last Synchronized": last_synced
                })
    return groups

def extract_details_from_group(driver, group: dict, users: list, categories: list, wait_for_element, group_edit_page: str):
    # Construct the URL dynamically using group UID
    group_url = f"{group_edit_page}{group['Group UID']}"
    driver.get(group_url)
    logging.info("Acessando a página do grupo: %s", group["Group Name"])
    try:
        wait_for_element(driver, "id", USER_CONTAINER_ID, timeout=10)
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
        logging.error("Falha ao extrair detalhes para o grupo: %s (%s)", group["Group Name"], e)
