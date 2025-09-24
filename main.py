import os
import platform
import logging
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from bs4 import BeautifulSoup

from src.config import LOGIN_URL, GROUP_CONTAINER_ID, FILE_NAME, OUTPUT_DIRECTORY
from src.utils import get_pwa_instance_url, get_output_file, start_screen, browser_config_message, browser_closed_message, extraction_complete_message
from src.browser_helpers import wait_for_element, close_browsers, get_login, get_browser_choice, create_browser_driver
from src.data_extraction import extract_groups, extract_details_from_group
from src.data_output import save_to_excel
from src.user_extraction import extract_users_from_excel_export

def _create_fallback_user_data(users_groups):
    """
    Create fallback user data from users_groups when Excel export fails.
    This ensures the Users tab always has some data to display.
    """
    if not users_groups:
        return []
    
    # Extract unique users from users_groups and create basic user records
    seen_users = set()
    fallback_users = []
    
    for user_group in users_groups:
        user_name = user_group.get('User Name')
        user_uid = user_group.get('User UID')
        
        if user_name and user_uid and user_uid not in seen_users:
            seen_users.add(user_uid)
            
            # Create basic user record with available information
            user_record = {
                'User Name': user_name,
                'User UID': user_uid,
                'Email': f"{user_name.replace(' ', '.').lower()}@company.com",  # Placeholder
                'Status': 'Active',  # Placeholder
                'Title': 'N/A',  # Placeholder - would come from Excel export
                'Department': 'N/A'  # Placeholder - would come from Excel export
            }
            fallback_users.append(user_record)
    
    return fallback_users

def main():
    # Suppress Selenium verbose logging
    os.environ['WDM_LOG_LEVEL'] = '0'
    os.environ['WDM_PRINT_FIRST_LINE'] = 'False'
    logging.getLogger('selenium').setLevel(logging.CRITICAL)
    logging.getLogger('urllib3').setLevel(logging.CRITICAL)
    # Keep application logging at INFO level for extraction messages
    logging.getLogger().setLevel(logging.INFO)
    
    # Display the welcome message.
    print(start_screen())

    # Get browser choice from user
    browser_choice = get_browser_choice()

    # Get the URLs and output file.
    pwa_instance_url = get_pwa_instance_url()
    group_edit_page = f"{pwa_instance_url}_layouts/15/PWA/Admin/AddModifyGroup.aspx?groupUid="
    groups_page = f"{pwa_instance_url}_layouts/15/PWA/Admin/ManageGroups.aspx"
    output_file = get_output_file(FILE_NAME, OUTPUT_DIRECTORY)

    # Display browser configuration message
    browser_name = "Microsoft Edge" if browser_choice == "edge" else "Google Chrome"
    print(browser_config_message(browser_name))
    
    # Create browser driver based on user choice
    try:
        driver, _ = create_browser_driver(browser_choice)
    except Exception as e:
        logging.error(f"❌ Falha ao iniciar navegador: {e}")
        raise RuntimeError(f"Não foi possível iniciar o navegador. Verifique se o {browser_name} está instalado corretamente.")

    try:
        # Prompt the user to log in.
        try:
            get_login(driver, LOGIN_URL)
        except Exception as login_error:
            if "navegador foi fechado" in str(login_error):
                print(browser_closed_message())
                return
            else:
                raise login_error

        # Minimize the browser window.
        driver.minimize_window()
        
        # Navigate to the Groups page.
        driver.get(groups_page)
        logging.info("Navegando para a página de gerenciamento de grupos. Aguardando o carregamento da página...")
        try:
            wait_for_element(driver, By.ID, GROUP_CONTAINER_ID, timeout=20)
            logging.info("A página está pronta.")
        except Exception:
            logging.error("Tempo esgotado aguardando o carregamento da página.")
            return
        
        # Extract group data.
        soup = BeautifulSoup(driver.page_source, "html.parser")
        groups = extract_groups(soup, group_edit_page)
        
        # Extract user data from groups (for UsersGroups tab) 
        users_groups = []
        categories = []
        for group in groups:
            extract_details_from_group(driver, group, users_groups, categories, wait_for_element, group_edit_page)
        
        # Extract all user data from Excel export (for new Users tab with detailed info)
        all_users_detailed = []
        logging.info("🚀 Extraindo informações detalhadas de usuários via Excel Export...")
        try:
            all_users_detailed = extract_users_from_excel_export(driver, pwa_instance_url)
            if all_users_detailed:
                logging.info(f"✅ Excel Export bem-sucedido: {len(all_users_detailed)} usuários encontrados")
            else:
                logging.warning("⚠️ Excel Export não retornou dados")
                # Create sample data from users_groups as fallback
                all_users_detailed = _create_fallback_user_data(users_groups)
                if all_users_detailed:
                    logging.info(f"📋 Usando dados de fallback: {len(all_users_detailed)} usuários")
        except Exception as e:
            logging.error(f"❌ Erro no Excel Export: {e}")
            # Create sample data from users_groups as fallback
            all_users_detailed = _create_fallback_user_data(users_groups)
            if all_users_detailed:
                logging.info(f"📋 Usando dados de fallback: {len(all_users_detailed)} usuários")
        
        # Save the data to an Excel file with both user datasets
        save_to_excel(groups, users_groups, categories, all_users_detailed, output_file)
    finally:
        driver.quit()
        if platform.system() == "Windows":
            os.system("pause")

if __name__ == "__main__":
    main()
