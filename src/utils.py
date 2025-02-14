import os
import re
import textwrap

def validate_pwa_url(url: str) -> bool:
    pattern = r"^https://[a-zA-Z0-9-]+\.sharepoint\.com/sites/[a-zA-Z0-9_-]+/$"
    return re.match(pattern, url)

def get_output_file(file_name: str, output_directory: str) -> str:
    output_path = os.path.expanduser(output_directory)
    file_counter = 1
    output_file = os.path.join(output_path, f"{file_name}.xlsx")
    while os.path.exists(output_file):
        output_file = os.path.join(output_path, f"{file_name}_{file_counter}.xlsx")
        file_counter += 1
    return output_file

def get_pwa_instance_url() -> str:
    while True:
        url = input("Digite a URL da instância do PWA: ").strip().rstrip("/") + "/"
        if validate_pwa_url(url):
            return url
        print("Formato de URL inválido. A URL deve seguir este padrão: https://TENANT_NAME.sharepoint.com/sites/PWA_SITE/")

def start_screen():
    welcome_message = textwrap.dedent(
    """
        ==========================================================================
                        Bem-vindo ao extrator de dados do PWA
        ==========================================================================
        Requisitos:
          - O navegador Chrome deve estar instalado.            
          - Possuir em mãos o link da instância do PWA 
            (Ex.: https://<TENANT_NAME>.sharepoint.com/sites/<PWA_SITE>/).
        ---------------------------------------------------------------------------
        ATENÇÃO: O script fechará o Chrome após inserir o link. Salve seu trabalho!
        ---------------------------------------------------------------------------
        Você poderá usar seu perfil do Chrome se desejar aproveitar a sessão logada 
        no PWA ou, entrar em um sessão não autenticada (convidado).
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