import logging
import psutil
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from .config import CHROME_TIMEOUT

def check_profile_availability(profile_path):
    """Check if a Chrome profile is available for use."""
    try:
        # Check for common lock files
        lock_files = ["SingletonLock", "SingletonSocket", "lockfile"]
        for lock_file in lock_files:
            lock_path = os.path.join(profile_path, lock_file)
            if os.path.exists(lock_path):
                try:
                    # Try to read the lock file
                    with open(lock_path, 'r') as f:
                        pass
                except (PermissionError, OSError):
                    return False
        
        # Check if we can write to the profile directory
        test_file = os.path.join(profile_path, "test_write_access")
        try:
            with open(test_file, 'w') as f:
                f.write("test")
            os.remove(test_file)
            return True
        except (PermissionError, OSError):
            return False
            
    except Exception:
        return False

def is_profile_in_use(profile_path):
    """Check if a Chrome profile is currently being used by another process."""
    try:
        # Check for lock files that Chrome creates when using a profile
        lock_files = ["SingletonLock", "SingletonSocket", "lockfile"]
        for lock_file in lock_files:
            lock_path = os.path.join(profile_path, lock_file)
            if os.path.exists(lock_path):
                try:
                    # Try to open the lock file exclusively
                    with open(lock_path, 'a') as f:
                        pass
                except (PermissionError, OSError):
                    return True
        return False
    except Exception:
        return True  # Assume it's in use if we can't check

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
        print("\n" + "="*80)
        print("               ESCOLHA DO MODO DE AUTENTICAÇÃO")
        print("="*80)
        print("⚠️  AVISO: Chrome 139 possui restrições de segurança que impedem")
        print("   o uso de perfis existentes com automação (Chrome 136+)")
        print("   Mais informações: https://developer.chrome.com/blog/remote-debugging-port")
        print("")
        print("OPÇÃO 1: Tentar Perfil Existente (Experimental)")
        print("  ⚠️  Limitação: Pode não funcionar devido às restrições do Chrome 139")
        print("  💡 Quando usar: Apenas para teste, geralmente não funciona")
        print("")
        print("OPÇÃO 2: Modo Convidado (Recomendado) ⭐")
        print("  ✅ Vantagem: Sempre funciona, sem conflitos")
        print("  ✅ Vantagem: Compatível com todas as versões do Chrome")
        print("  ⚠️  Limitação: Precisa inserir credenciais manualmente")
        print("  💡 Quando usar: Para operação confiável (RECOMENDADO)")
        print("="*80)
        
        use_guest = input("Escolha o modo (opções: [1] experimental / [2] recomendado): ").strip()
    except Exception as e:
        logging.error("Erro ao ler a entrada do usuário: %s", e)
        raise

    if use_guest in ["2", "recomendado", "guest", "convidado"]:
        options.add_argument("--guest")
        print("✅ Usando modo convidado - máxima compatibilidade garantida!")
        logging.info("Usando modo convidado para máxima compatibilidade")
        
    elif use_guest in ["1", "experimental", "perfil", "profile"]:
        print("⚠️  Modo experimental selecionado - pode não funcionar no Chrome 139+")
        
        if not chrome_path or not profile_mapping:
            print("❌ Não foi possível detectar os perfis do Chrome.")
            print("🔄 Mudando automaticamente para modo convidado...")
            options.add_argument("--guest")
            return options

        print("\nPerfis do Chrome disponíveis:")
        for index, (profile_name, folder) in enumerate(profile_mapping.items(), 1):
            print(f"  [{index}] {profile_name}")

        try:
            selected_index_str = input("\nDigite o número do perfil que deseja usar (ou 0 para cancelar): ")
            if selected_index_str.strip() == "0":
                print("🔄 Mudando para modo convidado...")
                options.add_argument("--guest")
                return options
                
            selected_index = int(selected_index_str) - 1
            profile_list = list(profile_mapping.items())
            if selected_index < 0 or selected_index >= len(profile_list):
                raise ValueError("Índice selecionado fora do intervalo.")
            profile_name, profile_folder = profile_list[selected_index]
        except ValueError as ve:
            print(f"❌ Entrada inválida: {ve}")
            print("🔄 Mudando para modo convidado...")
            options.add_argument("--guest")
            return options
        except Exception as e:
            logging.error("Erro ao selecionar o perfil: %s", e)
            print("🔄 Mudando para modo convidado...")
            options.add_argument("--guest")
            return options

        # Check if profile directory exists
        original_profile_path = os.path.join(chrome_path, profile_folder)
        if not os.path.exists(original_profile_path):
            print(f"❌ Perfil {profile_name} não encontrado.")
            print("🔄 Mudando automaticamente para modo convidado...")
            options.add_argument("--guest")
            return options
            
        # Attempt to use the profile (likely to fail in Chrome 139+)
        try:
            options.add_argument(f"--user-data-dir={chrome_path}")
            options.add_argument(f"--profile-directory={profile_folder}")
            
            # Minimal arguments for Chrome 139+ compatibility
            options.add_argument("--disable-background-timer-throttling")
            options.add_argument("--disable-backgrounding-occluded-windows")
            options.add_argument("--disable-renderer-backgrounding")
            options.add_argument("--disable-background-networking")
            options.add_argument("--disable-hang-monitor")
            options.add_argument("--disable-web-security")  # May help with restrictions
            options.add_argument("--disable-features=VizDisplayCompositor")
            
            print(f"⚠️  Tentativa experimental de usar perfil {profile_name}")
            print("💡 Se falhar, o sistema mudará automaticamente para modo convidado")
                
        except Exception as profile_error:
            logging.warning(f"Erro ao configurar perfil {profile_name}: {profile_error}")
            print("🔄 Mudando automaticamente para modo convidado...")
            options.add_argument("--guest")
    else:
        print("❌ Opção inválida. Usando modo convidado por segurança.")
        options.add_argument("--guest")

    return options


def wait_for_element(driver, by, identifier, timeout=CHROME_TIMEOUT):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, identifier)))

def close_chrome():
    """Closes any running Chrome processes aggressively."""
    import time
    import subprocess
    
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
