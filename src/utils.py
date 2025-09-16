import os
import re
import textwrap

def validate_pwa_url(url: str) -> bool:
    # Accept URLs with the basic SharePoint site structure, anything after the site name is acceptable
    pattern = r"^https://[a-zA-Z0-9-]+\.sharepoint\.com/sites/[a-zA-Z0-9_-]+(/.*)?$"
    return re.match(pattern, url)

def get_output_file(file_name: str, output_directory: str) -> str:
    output_path = os.path.expanduser(output_directory)
    file_counter = 1
    output_file = os.path.join(output_path, f"{file_name}.xlsx")
    while os.path.exists(output_file):
        output_file = os.path.join(output_path, f"{file_name}_{file_counter}.xlsx")
        file_counter += 1
    return output_file

def invalid_url_message():
    """Return invalid URL format message."""
    message = textwrap.dedent("""
        Formato de URL inválido. A URL deve seguir este padrão:
          - https://TENANT_NAME.sharepoint.com/sites/PWA_SITE/
          - https://TENANT_NAME.sharepoint.com/sites/PWA_SITE/default.aspx
          - https://TENANT_NAME.sharepoint.com/sites/PWA_SITE/qualquer/caminho/
    """)
    return message


def get_pwa_instance_url() -> str:
    while True:
        url = input("Digite a URL da instância do PWA: \n").strip()
        
        if validate_pwa_url(url):
            # Extract the base URL (everything up to and including the site name)
            # Pattern: https://tenant.sharepoint.com/sites/SITE_NAME/
            match = re.match(r"(https://[a-zA-Z0-9-]+\.sharepoint\.com/sites/[a-zA-Z0-9_-]+)", url)
            if match:
                base_url = match.group(1) + "/"
                return base_url
        
        print(invalid_url_message())

def start_screen():
    welcome_message = textwrap.dedent(
    """
        ==========================================================================
                        Bem-vindo ao extrator de dados do PWA
        ==========================================================================
        
        Requisitos:
          - Navegador Chrome ou Microsoft Edge instalado.            
          - Possuir em mãos o link da instância do PWA 
            (Ex.: https://<TENANT_NAME>.sharepoint.com/sites/<PWA_SITE>/

        ---------------------------------------------------------------------------
        ℹ️  INFORMAÇÃO: O script abrirá uma nova janela do navegador selecionado.
        ---------------------------------------------------------------------------

        Saída:
          - Um arquivo Excel será gerado na pasta Downloads com o nome 'pwa_data'.
          - No arquivo há 3 páginas:
              - Users: Lista de usuários associados a cada grupo
              - Groups: Lista de grupos e suas informações
              - Categories: Lista de categorias associadas a cada grupo
        ---------------------------------------------------------------------------
    """)
    return welcome_message


def browser_config_message(browser_name="Chrome"):
    """Return browser configuration information message."""
    message = textwrap.dedent(f"""
        
        ---------------------------------------------------------------------------
        🚀                     INICIANDO {browser_name.upper()}                     🚀
        ---------------------------------------------------------------------------
        ℹ️   Abrindo nova janela em modo convidado. 
        
        Você precisará inserir suas credenciais manualmente.
        ---------------------------------------------------------------------------
    """)
    return message


def chrome_config_message():
    """Return Chrome configuration information message - legacy function."""
    return browser_config_message("Chrome")


def browser_closed_message():
    """Return message when browser is closed during login."""
    message = textwrap.dedent("""
        ===========================================================================
        ⚠️          O navegador foi fechado durante o processo de login.        ⚠️
        ===========================================================================
        Execute o script novamente e aguarde o redirecionamento após o login.
        ===========================================================================
    """)
    return message


def extraction_complete_message(output_file):
    """Return completion message with file location."""
    return f"INFO: Extração de dados concluída! Salvo em {output_file}"