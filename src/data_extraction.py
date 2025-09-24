import logging
import time
from bs4 import BeautifulSoup
from .config import GROUP_CONTAINER_ID, USER_CONTAINER_ID, CATEGORY_CONTAINER_ID, USERS_GRID_ID

def extract_groups(soup: BeautifulSoup, group_edit_page: str) -> list:
    groups = []
    grid_rows = soup.find_all("tr", id=GROUP_CONTAINER_ID)
    logging.info(f"Encontradas {len(grid_rows)} linhas de grupos na página")
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
                logging.info(f"Grupo extraído: '{group_name}'")
    logging.info(f"Total de grupos extraídos: {len(groups)}")
    return groups

def extract_details_from_group(driver, group: dict, users: list, categories: list, wait_for_element, group_edit_page: str):
    # Construct the URL dynamically using group UID
    group_url = f"{group_edit_page}{group['Group UID']}"
    driver.get(group_url)
    try:
        wait_for_element(driver, "id", USER_CONTAINER_ID, timeout=10)
        group_soup = BeautifulSoup(driver.page_source, "html.parser")
        user_container = group_soup.find("td", id=USER_CONTAINER_ID)
        category_container = group_soup.find("td", id=CATEGORY_CONTAINER_ID)

        user_count = 0
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
                user_count += 1

        category_count = 0
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
                category_count += 1
        
        logging.info("Grupo '%s' processado: %d usuários, %d categorias", group["Group Name"], user_count, category_count)
    except Exception as e:
        logging.error("Falha ao extrair detalhes para o grupo: %s (%s)", group["Group Name"], e)


def extract_details_from_users(driver, manage_users_url):
    """
    Extract users from PWA using BeautifulSoup with simple pagination support.
    """
    try:
        logging.info("Navegando para a página de usuários...")
        
        # Navigate to ManageUsers page
        driver.get(manage_users_url)
        time.sleep(3)  # Wait for page to load
        
        all_users = []
        page_number = 1
        
        while True:
            logging.info(f"Processando página {page_number}...")
            
            # Parse current page with BeautifulSoup
            soup = BeautifulSoup(driver.page_source, "html.parser")
            page_users = extract_users(soup)
            
            if page_users:
                all_users.extend(page_users)
                logging.info(f"Página {page_number}: {len(page_users)} usuários extraídos")
            else:
                logging.warning(f"⚠️ Página {page_number}: nenhum usuário encontrado")
            
            # Try to go to next page using simple BeautifulSoup + driver click
            if _go_to_next_page_simple(driver):
                page_number += 1
                time.sleep(3)  # Wait for new page to load
            else:
                break
        
        if all_users:
            logging.info(f"Extração finalizada: {len(all_users)} usuários de {page_number} processados")
            return all_users
        else:
            logging.warning("⚠️ Nenhum usuário encontrado em todas as páginas")
            return []
            
    except Exception as e:
        logging.error(f"❌ Erro na extração de usuários: {e}")
        return []


def extract_users(soup: BeautifulSoup) -> list:
    """Extract users from ManageUsers page using BeautifulSoup (similar to extract_groups)."""
    users = []
    
    # Find the users grid table
    users_table = soup.find("table", id=USERS_GRID_ID)
    if not users_table:
        logging.error(f"❌ Tabela de usuários não encontrada. ID procurado: {USERS_GRID_ID}")
        return []
    
    # Find all data rows (skip header row)
    grid_rows = users_table.find_all("tr", id="GridDataRow")
    logging.info(f"Encontradas {len(grid_rows)} linhas de usuários na página")
    
    for row in grid_rows:
        columns = row.find_all("td")
        if len(columns) >= 6:  # Need at least 6 columns (checkbox + 5 data columns)
            # Extract user ID from row attribute
            user_id = row.get("rowid")
            
            # Extract data from columns (skip checkbox column at index 0)
            user_name = columns[1].get_text(strip=True) if len(columns) > 1 else ""
            email_address = columns[2].get_text(strip=True) if len(columns) > 2 else ""
            logon_account = columns[3].get_text(strip=True) if len(columns) > 3 else ""
            state = columns[4].get_text(strip=True) if len(columns) > 4 else ""
            rbs = columns[5].get_text(strip=True) if len(columns) > 5 else ""
            
            if user_name or email_address:  # Only add if we have essential data
                users.append({
                    "USER UID": user_id,
                    "User Name": user_name,
                    "Email Address": email_address,
                    "User Logon Account": logon_account,
                    "State": state,
                    "RBS": rbs
                })
    
    logging.info(f"Total de usuários extraídos: {len(users)}")
    return users


def _go_to_next_page_simple(driver):
    """Simple pagination: find and click next page link if available."""
    try:
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        
        # Find the "Próxima" link directly using Selenium (avoid BeautifulSoup for clicking)
        try:
            # Look for the next page link with "Próxima" text
            next_link = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'XmlGridPrevNextLink') and contains(text(), 'Próxima')]"))
            )
            
            # Get link info for logging
            link_text = next_link.text.strip()
            link_href = next_link.get_attribute('href')
            logging.info(f"🔗 Link encontrado: '{link_text}' -> {link_href}")
            
            # Click the link directly instead of executing JavaScript
            next_link.click()
            logging.info(f"▶️ Navegando para próxima página via clique direto")
            return True
            
        except Exception as find_error:
            logging.info("🔚 Nenhum link 'Próxima' encontrado - última página alcançada")
            logging.debug(f"Detalhes: {find_error}")
            return False
        
    except Exception as e:
        logging.error(f"❌ Erro ao navegar para próxima página: {e}")
        return False
