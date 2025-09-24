"""
User data extraction from downloaded Excel files.
"""
import pandas as pd
import logging
import os
from typing import List, Dict, Optional

def process_downloaded_users_excel(file_path: str) -> List[Dict]:
    """
    Process the downloaded Excel file from ManageUsers.aspx Export to Excel button.
    Returns a list of user dictionaries compatible with existing data structure.
    """
    if not file_path or not os.path.exists(file_path):
        logging.error(f"Arquivo Excel não encontrado: {file_path}")
        return []
    
    try:
        logging.info(f"Processando arquivo Excel de usuários: {file_path}")
        
        # Read the Excel file
        df = pd.read_excel(file_path)
        
        logging.info(f"Arquivo Excel carregado com {len(df)} registros")
        logging.info(f"Colunas disponíveis: {list(df.columns)}")
        
        # Convert DataFrame to list of dictionaries
        users_data = []
        for _, row in df.iterrows():
            user_dict = {}
            
            # Map common column names to our expected format
            # Note: Adjust these mappings based on actual SharePoint export format
            column_mappings = {
                # Common SharePoint user export columns
                'Name': 'User Name',
                'User Name': 'User Name', 
                'Full Name': 'User Name',
                'Display Name': 'User Name',
                'Email': 'Email',
                'E-mail': 'Email',
                'Email Address': 'Email',
                'Account': 'Account',
                'Login Name': 'Account',
                'User Account': 'Account',
                'Groups': 'Groups',
                'Security Groups': 'Groups',
                'Group Membership': 'Groups',
                'Department': 'Department',
                'Title': 'Title',
                'Job Title': 'Title',
            }
            
            # Apply mappings and extract data
            for excel_col, our_col in column_mappings.items():
                if excel_col in df.columns:
                    value = row[excel_col]
                    if pd.notna(value):  # Skip NaN values
                        user_dict[our_col] = str(value).strip()
            
            # Ensure we have at least a user identifier
            if user_dict.get('User Name') or user_dict.get('Email'):
                users_data.append(user_dict)
        
        logging.info(f"Processados {len(users_data)} usuários do arquivo Excel")
        
        # Clean up - remove the downloaded file after processing
        try:
            os.remove(file_path)
            logging.info(f"Arquivo temporário removido: {file_path}")
        except Exception as e:
            logging.warning(f"Não foi possível remover arquivo temporário: {e}")
        
        return users_data
        
    except Exception as e:
        logging.error(f"Erro ao processar arquivo Excel: {e}")
        return []


def extract_users_from_excel_export(driver, pwa_base_url: str) -> List[Dict]:
    """
    High-level function to extract user data using Excel export method.
    Returns list of user dictionaries.
    """
    from .browser_helpers import export_users_to_excel
    from .config import MANAGE_USERS_PATH
    
    # Construct the ManageUsers URL
    manage_users_url = pwa_base_url + MANAGE_USERS_PATH
    
    logging.info("Iniciando extração de usuários via Export to Excel...")
    
    # Download the Excel file
    downloaded_file = export_users_to_excel(driver, manage_users_url)
    
    if not downloaded_file:
        logging.error("Falha no download do arquivo Excel de usuários")
        return []
    
    # Process the downloaded file
    users_data = process_downloaded_users_excel(downloaded_file)
    
    if users_data:
        logging.info(f"✅ Extração via Excel concluída: {len(users_data)} usuários encontrados")
    else:
        logging.warning("⚠️ Nenhum usuário encontrado no arquivo Excel")
    
    return users_data