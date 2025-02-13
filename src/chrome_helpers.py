import os
import platform
import json
import logging
import psutil
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from .config import CHROME_TIMEOUT

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

import logging

def select_chrome_profile(chrome_path, profile_mapping, options):
    try:
        use_guest = input("Usar perfil existente? (opções: sim [s]; nao [n] -> Iniciar como convidado): ").strip().lower()
    except Exception as e:
        logging.error("Erro ao ler a entrada do usuário: %s", e)
        raise

    if use_guest in ["nao", "n"]:
        options.add_argument("--guest")
    elif use_guest in ["sim", "s"]:
        if not chrome_path or not profile_mapping:
            raise Exception("Não foi possível detectar os perfis do Chrome. Certifique-se de que o Chrome está instalado e foi utilizado antes de executar este script.")

        logging.info("Perfis do Chrome disponíveis:")
        for index, (profile_name, folder) in enumerate(profile_mapping.items(), 1):
            logging.info("[%s] %s", index, profile_name)

        try:
            selected_index_str = input("Digite o número correspondente ao perfil que deseja usar: ")
            selected_index = int(selected_index_str) - 1
            profile_list = list(profile_mapping.items())
            if selected_index < 0 or selected_index >= len(profile_list):
                raise ValueError("Índice selecionado fora do intervalo.")
            profile_name, profile_folder = profile_list[selected_index]
        except ValueError as ve:
            logging.error("Entrada inválida: %s", ve)
            raise
        except Exception as e:
            logging.error("Erro ao selecionar o perfil: %s", e)
            raise

        options.add_argument(f"--user-data-dir={chrome_path}")
        options.add_argument(f"--profile-directory={profile_folder}")
    else:
        raise Exception("Opção inválida. Por favor, responda com 'sim' ou 'nao' (ou 's' ou 'n').")

    return options


def wait_for_element(driver, by, identifier, timeout=CHROME_TIMEOUT):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, identifier)))

def close_chrome():
    """Closes any running Chrome processes."""
    for process in psutil.process_iter(attrs=["pid", "name"]):
        if "chrome" in process.info["name"].lower():
            try:
                os.kill(process.info["pid"], 9)
            except psutil.AccessDenied:
                logging.warning("Acesso negado ao processo %s. Ignorando.", process.info["pid"])
            except Exception as e:
                logging.error("Não foi possível fechar o processo %s: %s", process.info["pid"], e)
    logging.info("Tentativa de fechar todas as instâncias do Chrome em execução.")



def get_login(driver, login_url):     
    driver.get(login_url)
    logging.info("Por favor, complete o processo de login na janela do navegador...")

    try:        
        WebDriverWait(driver, 600).until(lambda d: "login.microsoftonline.com" not in d.current_url)
    except TimeoutException:
        logging.error("Autenticação não concluída dentro do tempo limite de 600 segundos.")
        raise Exception("Timeout: Microsoft authentication not completed.")

    logging.info("Autenticação concluída.")
