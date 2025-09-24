import logging
import psutil
import time
import subprocess
import os
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


def go_to_next_page(driver):
    """Simple pagination: find and click next page link if available."""
    try:
        # Find the "Próxima" link directly using Selenium (avoid BeautifulSoup for clicking)
        try:
            # Look for the next page link with "Próxima" (Portuguese) or "Next" (English) text
            next_link = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'XmlGridPrevNextLink') and (contains(text(), 'Próxima') or contains(text(), 'Next'))]"))
            )
            
            # Click the link directly instead of executing JavaScript
            next_link.click()
            return True
            
        except Exception as find_error:
            logging.info("🔚 Última página alcançada")
            logging.debug(f"Detalhes: {find_error}")
            return False
        
    except Exception as e:
        logging.error(f"❌ Erro ao navegar para próxima página: {e}")
        return False
