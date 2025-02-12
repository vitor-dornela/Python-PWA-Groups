import os
import re
import json
import psutil
import logging

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

