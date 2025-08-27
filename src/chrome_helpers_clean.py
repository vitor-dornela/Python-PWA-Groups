import logging
import psutil
import time
import subprocess
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from .config import CHROME_TIMEOUT


def wait_for_element(driver, by, identifier, timeout=CHROME_TIMEOUT):
    """Wait for an element to be present on the page."""
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, identifier)))


def close_chrome():
    """Closes any running Chrome processes aggressively."""
    # First, try to close Chrome gracefully
    try:
        subprocess.run(["taskkill", "/f", "/im", "chrome.exe"], 
                      capture_output=True, text=True, timeout=10)
    except (subprocess.TimeoutExpired, FileNotFoundError):
        pass
    
    # Wait a moment for processes to close
    time.sleep(2)
    
    # Use psutil for more thorough cleanup
    for process in psutil.process_iter(attrs=["pid", "name", "cmdline"]):
        if "chrome" in process.info["name"].lower():
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
    logging.info("Tentativa de fechar todas as instâncias do Chrome em execução.")


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
                
            except Exception as e:
                # If we can't get the current URL, the browser might be closed
                logging.error(f"Erro ao verificar o estado do navegador: {e}")
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
