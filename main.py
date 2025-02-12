import os
import platform
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from bs4 import BeautifulSoup

from src.config import LOGIN_URL, GROUP_CONTAINER_ID, FILE_NAME, OUTPUT_DIRECTORY
from src.utils import get_pwa_instance_url, get_output_file
from src.chrome_helpers import get_chrome_profiles, select_chrome_profile, wait_for_element, close_chrome
from src.data_extraction import extract_groups, extract_details_from_group
from src.data_output import save_to_excel

def main():
    # Get the URLs and output file.
    pwa_instance_url = get_pwa_instance_url()
    group_edit_page = f"{pwa_instance_url}_layouts/15/PWA/Admin/AddModifyGroup.aspx?groupUid="
    groups_page = f"{pwa_instance_url}_layouts/15/PWA/Admin/ManageGroups.aspx"
    output_file = get_output_file(FILE_NAME, OUTPUT_DIRECTORY)

    # Close any running Chrome instances.
    close_chrome()

    # Set up Selenium Chrome options.
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    chrome_path, profile_mapping = get_chrome_profiles()
    options = select_chrome_profile(chrome_path, profile_mapping, options)

    driver = webdriver.Chrome(options=options)
    try:
        # Prompt the user to log in.
        driver.get(LOGIN_URL)
        logging.info("Por favor, complete o processo de login na janela do navegador...")
        WebDriverWait(driver, 180).until(EC.url_changes(LOGIN_URL))

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
