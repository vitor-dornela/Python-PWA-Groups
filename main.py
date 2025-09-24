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

def get_extraction_method_choice():
    """Ask user to choose extraction method for user data."""
    print("\n📊 Selecione o método de extração de usuários:")
    print("1. Extração tradicional por página (método atual)")
    print("2. Export to Excel (novo método - mais rápido) 🚀")
    
    while True:
        choice = input("\nDigite 1 ou 2 (padrão: Excel Export): ").strip()
        
        if choice == "" or choice == "2":
            return "excel_export"
        elif choice == "1":
            return "page_scraping"
        else:
            print("❌ Opção inválida. Digite 1 ou 2.")

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

    # Get extraction method choice from user  
    extraction_method = get_extraction_method_choice()

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
        
        # Extract user data based on chosen method
        users = []
        categories = []
        
        if extraction_method == "excel_export":
            logging.info("🚀 Usando método Excel Export para extração de usuários...")
            try:
                # Try Excel export method first
                all_users = extract_users_from_excel_export(driver, pwa_instance_url)
                
                if all_users:
                    logging.info(f"✅ Excel Export bem-sucedido: {len(all_users)} usuários encontrados")
                    users = all_users
                else:
                    logging.warning("⚠️ Excel Export falhou, usando método tradicional como fallback...")
                    extraction_method = "page_scraping"
                    
            except Exception as e:
                logging.error(f"❌ Erro no Excel Export: {e}")
                logging.info("📄 Fazendo fallback para método tradicional...")
                extraction_method = "page_scraping"
        
        if extraction_method == "page_scraping":
            logging.info("📄 Usando método tradicional para extração de usuários...")
            # Traditional method: extract users from each group page
            for group in groups:
                extract_details_from_group(driver, group, users, categories, wait_for_element, group_edit_page)
        else:
            # For Excel export, we still need to extract categories from group pages
            logging.info("📋 Extraindo categorias dos grupos...")
            for group in groups:
                temp_users = []
                extract_details_from_group(driver, group, temp_users, categories, wait_for_element, group_edit_page)
        
        # Save the data to an Excel file.
        save_to_excel(groups, users, categories, output_file)
    finally:
        driver.quit()
        if platform.system() == "Windows":
            os.system("pause")

if __name__ == "__main__":
    main()
