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
from src.utils import get_pwa_instance_url, get_output_file, start_screen, chrome_config_message, browser_closed_message, extraction_complete_message
from src.chrome_helpers import wait_for_element, close_chrome, get_login
from src.data_extraction import extract_groups, extract_details_from_group
from src.data_output import save_to_excel

def main():
    # Suppress Selenium and Chrome verbose logging
    os.environ['WDM_LOG_LEVEL'] = '0'
    os.environ['WDM_PRINT_FIRST_LINE'] = 'False'
    logging.getLogger('selenium').setLevel(logging.CRITICAL)
    logging.getLogger('urllib3').setLevel(logging.CRITICAL)
    logging.getLogger().setLevel(logging.WARNING)  # Set root logger to WARNING level
    
    # Display the welcome message.
    print(start_screen())

    # Get the URLs and output file.
    pwa_instance_url = get_pwa_instance_url()
    group_edit_page = f"{pwa_instance_url}_layouts/15/PWA/Admin/AddModifyGroup.aspx?groupUid="
    groups_page = f"{pwa_instance_url}_layouts/15/PWA/Admin/ManageGroups.aspx"
    output_file = get_output_file(FILE_NAME, OUTPUT_DIRECTORY)

    # Close any running Chrome instances.
    close_chrome()

    # Display Chrome configuration message
    print(chrome_config_message())
    
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-gpu")
    options.add_argument("--guest")
    options.add_argument("--disable-background-timer-throttling")
    options.add_argument("--disable-backgrounding-occluded-windows")
    options.add_argument("--disable-renderer-backgrounding")
    options.add_argument("--disable-background-networking")
    options.add_argument("--disable-hang-monitor")
    
    # Suppress verbose Chrome logging and error messages
    options.add_argument("--log-level=3")  # Suppress INFO, WARNING, and ERROR
    options.add_argument("--silent")
    options.add_argument("--disable-logging")
    options.add_argument("--disable-default-apps")
    options.add_argument("--disable-background-mode")
    options.add_argument("--disable-background-networking")
    options.add_argument("--disable-sync")
    options.add_argument("--disable-translate")
    options.add_argument("--hide-scrollbars")
    options.add_argument("--mute-audio")
    options.add_experimental_option("useAutomationExtension", False)
    options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_experimental_option('useAutomationExtension', False)

    # Create the Chrome driver (single attempt, no retries needed)
    try:
        # Suppress Chrome's stderr output temporarily
        import sys
        from contextlib import redirect_stderr
        import io
        
        with redirect_stderr(io.StringIO()):
            driver = webdriver.Chrome(options=options)
        logging.info("✅ Chrome iniciado com sucesso no modo convidado")
    except Exception as e:
        logging.error(f"❌ Falha ao iniciar Chrome: {e}")
        raise Exception("Não foi possível iniciar o Chrome. Verifique se o Chrome está instalado corretamente.")

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
        users = []
        categories = []
        for group in groups:
            extract_details_from_group(driver, group, users, categories, wait_for_element, group_edit_page)
        
        # Save the data to an Excel file.
        save_to_excel(groups, users, categories, output_file)
    finally:
        driver.quit()
        if platform.system() == "Windows":
            os.system("pause")

if __name__ == "__main__":
    main()
