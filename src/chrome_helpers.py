import os
import platform
import json
import logging
import psutil
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
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

def select_chrome_profile(chrome_path, profile_mapping, options):
    use_guest = input("Usar perfil de convidado? (opções: sim; não [selecione um perfil]): ").strip().lower()
    if use_guest == "sim":
        options.add_argument("--guest")
    else:
        if not chrome_path or not profile_mapping:
            raise Exception("Não foi possível detectar os perfis do Chrome. Certifique-se de que o Chrome está instalado e foi utilizado antes de executar este script.")

        logging.info("Perfis do Chrome disponíveis:")
        for index, (profile_name, folder) in enumerate(profile_mapping.items(), 1):
            logging.info("[%s] %s", index, profile_name)

        selected_index = int(input("Digite o número correspondente ao perfil que deseja usar: ")) - 1
        profile_name, profile_folder = list(profile_mapping.items())[selected_index]
        options.add_argument(f"--user-data-dir={chrome_path}")
        options.add_argument(f"--profile-directory={profile_folder}")
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
