import logging
import psutil
import time
import subprocess
import os
import glob
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from .config import BROWSER_TIMEOUT


def wait_for_element(driver, by, identifier, timeout=BROWSER_TIMEOUT):
    """Wait for an element to be present on the page."""
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, identifier)))


def close_browsers(force_close=False):
    """Closes running browser processes only if explicitly requested."""
    if not force_close:
        logging.info("Mantendo navegadores existentes abertos...")
        return
        
    # Only close all browsers if explicitly requested
    browsers = ["chrome.exe", "msedge.exe", "MicrosoftEdge.exe"]
    logging.info("Fechando todas as instâncias dos navegadores...")
    
    for browser in browsers:
        # First, try to close browser gracefully
        try:
            subprocess.run(["taskkill", "/f", "/im", browser], 
                          capture_output=True, text=True, timeout=10)
        except (subprocess.TimeoutExpired, FileNotFoundError):
            pass
    
    # Wait a moment for processes to close
    time.sleep(2)
    
    # Use psutil for more thorough cleanup
    for process in psutil.process_iter(attrs=["pid", "name", "cmdline"]):
        process_name = process.info["name"].lower()
        if any(browser_name in process_name for browser_name in ["chrome", "msedge", "edge"]):
            try:
                # Use psutil to terminate the process more gracefully
                proc = psutil.Process(process.info["pid"])
                proc.terminate()
                # Wait for the process to terminate
                proc.wait(timeout=5)
            except psutil.NoSuchProcess:
                # Process already terminated
                pass
            except psutil.AccessDenied:
                logging.warning("Acesso negado ao processo %s. Ignorando.", process.info["pid"])
            except psutil.TimeoutExpired:
                # Force kill if terminate doesn't work
                try:
                    proc.kill()
                    proc.wait(timeout=3)
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
            except Exception as e:
                logging.error("Não foi possível fechar o processo %s: %s", process.info["pid"], e)
    
    # Give extra time to ensure all processes are fully closed and file locks are released
    time.sleep(3)


def close_chrome():
    """Legacy function - redirects to close_browsers() for backward compatibility."""
    close_browsers()


def get_browser_choice():
    """Ask user to choose between Chrome and Edge browsers."""
    print("🌐 Escolha o navegador:")
    print("1. Microsoft Edge (padrão)")
    print("2. Google Chrome")
    
    while True:
        choice = input("Navegador escolhido: ").strip()
        
        if choice == "" or choice == "1":
            return "edge"
        elif choice == "2":
            return "chrome"
        else:
            print("❌ Opção inválida. Digite 1 ou 2.")


def create_chrome_driver():
    """Create and configure Chrome WebDriver."""
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
    
    # Download preferences
    download_prefs = {
        "profile.default_content_settings.popups": 0,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", download_prefs)
    options.add_experimental_option("useAutomationExtension", False)
    options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_experimental_option('useAutomationExtension', False)

    # Create the Chrome driver
    import sys
    from contextlib import redirect_stderr
    import io
    
    with redirect_stderr(io.StringIO()):
        driver = webdriver.Chrome(options=options)
    
    return driver


def create_edge_driver():
    """Create and configure Edge WebDriver."""
    options = webdriver.EdgeOptions()
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
    
    # Suppress verbose Edge logging and error messages
    options.add_argument("--log-level=3")
    options.add_argument("--silent")
    options.add_argument("--disable-logging")
    options.add_argument("--disable-default-apps")
    options.add_argument("--disable-background-mode")
    options.add_argument("--disable-background-networking")
    options.add_argument("--disable-sync")
    options.add_argument("--disable-translate")
    options.add_argument("--hide-scrollbars")
    options.add_argument("--mute-audio")
    
    # Download preferences for Edge
    download_prefs = {
        "profile.default_content_settings.popups": 0,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", download_prefs)
    options.add_experimental_option("useAutomationExtension", False)
    options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])

    # Create the Edge driver
    import sys
    from contextlib import redirect_stderr
    import io
    
    with redirect_stderr(io.StringIO()):
        driver = webdriver.Edge(options=options)
    
    return driver


def create_browser_driver(browser_choice="chrome"):
    """Create WebDriver based on browser choice."""
    try:
        if browser_choice == "edge":
            driver = create_edge_driver()
            logging.info("✅ Microsoft Edge iniciado com sucesso no modo convidado")
            return driver, "Edge"
        else:  # Default to Chrome
            driver = create_chrome_driver()
            logging.info("✅ Chrome iniciado com sucesso no modo convidado")
            return driver, "Chrome"
            
    except Exception as e:
        browser_name = "Edge" if browser_choice == "edge" else "Chrome"
        logging.error(f"❌ Falha ao iniciar {browser_name}: {e}")
        
        if browser_choice == "edge":
            raise Exception("Não foi possível iniciar o Microsoft Edge. Verifique se o Edge está instalado corretamente.")
        else:
            raise Exception("Não foi possível iniciar o Chrome. Verifique se o Chrome está instalado corretamente.")


def get_login(driver, login_url):     
    """Handle user login and wait for completion."""
    driver.get(login_url)
    logging.info("Por favor, complete o processo de login na janela do navegador...")
    logging.info("IMPORTANTE: Não feche o navegador! Aguarde até ser redirecionado após o login.")

    try:        
        def check_login_completion(d):
            try:
                # Check if browser is still alive
                current_url = d.current_url
                if current_url is None:
                    return False
                    
                # Check if we're no longer on the Microsoft login page
                return "login.microsoftonline.com" not in current_url
                
            except Exception:
                # If we can't get the current URL, the browser might be closed
                # Don't log the full error details, just raise a clean exception
                raise Exception("O navegador foi fechado durante o processo de login.")
        
        WebDriverWait(driver, 600).until(check_login_completion)
        
    except TimeoutException:
        logging.error("Autenticação não concluída dentro do tempo limite de 600 segundos.")
        raise Exception("Timeout: Microsoft authentication not completed.")
    except Exception as e:
        if "navegador foi fechado" in str(e):
            raise e
        else:
            logging.error(f"Erro durante o processo de login: {e}")
            raise Exception("Erro durante o processo de login. Verifique se o navegador não foi fechado.")

    logging.info("Autenticação concluída.")


def export_users_to_excel(driver, manage_users_url, download_dir=None):
    """
    Navigate to ManageUsers page and download Excel export.
    Uses a dedicated temporary directory to ensure we get the right file.
    Returns the path to the downloaded file or None if failed.
    """
    import tempfile
    import shutil
    
    # Create a unique temporary directory for this download
    temp_download_dir = tempfile.mkdtemp(prefix="pwa_users_export_")
    
    try:
        logging.info("Navegando para a página de gerenciamento de usuários...")
        driver.get(manage_users_url)
        
        # Wait for page to load
        wait_for_element(driver, By.TAG_NAME, "body", timeout=20)
        time.sleep(3)  # Extra wait for page to fully render
        
        # Configure browser to download to our temp directory
        _configure_browser_downloads(driver, temp_download_dir)
        
        # Look for Export to Excel button - try extensive selectors
        export_button = None
        possible_selectors = [
            # Standard button/input patterns
            "//input[@value='Export to Excel']",
            "//button[contains(text(), 'Export to Excel')]",
            "//a[contains(text(), 'Export to Excel')]",
            "//input[contains(@id, 'Export')]",
            "//input[contains(@class, 'export')]",
            "//*[contains(text(), 'Export to Excel')]",
            "//span[contains(text(), 'Export to Excel')]/parent::*",
            "//*[@title='Export to Excel']",
            
            # SharePoint specific patterns
            "//input[contains(@id, 'ExportToExcel')]", 
            "//input[contains(@name, 'Export')]",
            "//a[contains(@href, 'Export')]",
            "//button[contains(@class, 'export')]",
            "//*[contains(@onclick, 'Export')]",
            
            # Ribbon/menu patterns
            "//span[text()='Export to Excel']",
            "//div[contains(@title, 'Export')]",
            "//*[@aria-label='Export to Excel']",
            
            # Case variations
            "//input[@value='Export To Excel']",
            "//*[contains(text(), 'Export To Excel')]",
            "//*[contains(text(), 'EXPORT TO EXCEL')]",
            "//*[contains(text(), 'export to excel')]",
            
            # Partial text matches
            "//*[contains(text(), 'Export')]//parent::*[contains(text(), 'Excel')]",
            "//*[contains(text(), 'Excel')]//parent::*[contains(text(), 'Export')]"
        ]
        
        logging.info("Procurando botão 'Export to Excel' com múltiplos seletores...")
        
        for i, selector in enumerate(possible_selectors):
            try:
                export_button = driver.find_element(By.XPATH, selector)
                logging.info(f"✅ Botão 'Export to Excel' encontrado usando seletor #{i+1}: {selector}")
                break
            except NoSuchElementException:
                continue
        
        if not export_button:
            logging.error("❌ Botão 'Export to Excel' não encontrado com nenhum seletor")
            
            # Enhanced debugging information
            try:
                page_text = driver.page_source.lower()
                
                # Check for various export-related terms
                export_terms = ['export', 'excel', 'download', 'save', 'download']
                found_terms = [term for term in export_terms if term in page_text]
                
                if found_terms:
                    logging.info(f"🔍 Termos relacionados encontrados na página: {found_terms}")
                
                # Look for any buttons or inputs on the page
                buttons = driver.find_elements(By.TAG_NAME, "button")
                inputs = driver.find_elements(By.TAG_NAME, "input")
                links = driver.find_elements(By.TAG_NAME, "a")
                
                logging.info(f"🔍 Elementos encontrados na página: {len(buttons)} botões, {len(inputs)} inputs, {len(links)} links")
                
                # Log some button/input text for debugging
                for i, button in enumerate(buttons[:5]):  # First 5 buttons
                    try:
                        text = button.text or button.get_attribute('value') or button.get_attribute('title')
                        if text:
                            logging.info(f"   Botão {i+1}: '{text}'")
                    except:
                        pass
                        
                for i, inp in enumerate(inputs[:5]):  # First 5 inputs
                    try:
                        value = inp.get_attribute('value') or inp.get_attribute('title')
                        if value:
                            logging.info(f"   Input {i+1}: '{value}'")
                    except:
                        pass
                
            except Exception as debug_error:
                logging.warning(f"Erro ao coletar informações de debug: {debug_error}")
            
            return None
        
        logging.info("Clicando no botão 'Export to Excel'...")
        
        # Record timestamp before clicking
        click_time = time.time()
        export_button.click()
        
        # Wait for download to complete in temp directory
        max_wait_time = 45  # seconds - increased for larger files
        start_time = time.time()
        
        downloaded_file = None
        while time.time() - start_time < max_wait_time:
            # Look for any Excel files in temp directory
            excel_files = glob.glob(os.path.join(temp_download_dir, "*.xlsx"))
            
            if excel_files:
                # Find the file created after our click
                for file_path in excel_files:
                    file_creation_time = os.path.getctime(file_path)
                    if file_creation_time >= click_time:
                        # Verify file is completely downloaded (not partial)
                        if _is_download_complete(file_path):
                            downloaded_file = file_path
                            break
                
                if downloaded_file:
                    break
            
            time.sleep(1)
        
        if not downloaded_file:
            logging.error("Timeout esperando o download do arquivo Excel")
            return None
        
        # Move file to final destination with timestamp to avoid conflicts
        if download_dir is None:
            download_dir = os.path.expanduser("~/Downloads")
        
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        final_filename = f"pwa_users_export_{timestamp}.xlsx"
        final_path = os.path.join(download_dir, final_filename)
        
        shutil.move(downloaded_file, final_path)
        logging.info(f"Download completo e movido para: {final_path}")
        
        return final_path
        
    except Exception as e:
        logging.error(f"Erro ao exportar usuários para Excel: {e}")
        return None
        
    finally:
        # Clean up temp directory
        try:
            shutil.rmtree(temp_download_dir, ignore_errors=True)
        except:
            pass


def _configure_browser_downloads(driver, download_dir):
    """Configure browser to download files to specific directory."""
    try:
        # This works for Chrome - need to check browser type for Edge
        driver.execute_cdp_cmd('Page.setDownloadBehavior', {
            'behavior': 'allow',
            'downloadPath': download_dir
        })
        logging.info(f"Configurado diretório de download: {download_dir}")
    except Exception as e:
        logging.warning(f"Não foi possível configurar diretório de download: {e}")


def _is_download_complete(file_path, min_size_bytes=1024):
    """
    Check if download is complete by verifying:
    1. File exists and has minimum size
    2. File is not being written to (stable size over time)
    """
    try:
        if not os.path.exists(file_path):
            return False
        
        # Check minimum size
        initial_size = os.path.getsize(file_path)
        if initial_size < min_size_bytes:
            return False
        
        # Wait and check if size is stable (not growing)
        time.sleep(1)
        final_size = os.path.getsize(file_path)
        
        return initial_size == final_size
        
    except Exception:
        return False
