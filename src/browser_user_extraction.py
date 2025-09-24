"""
Browser-based user data extraction - alternative to file download
"""
import logging
import time
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import pandas as pd


def extract_users_from_browser_export(driver, manage_users_url):
    """
    Navigate to ManageUsers page, click Export to Excel, and extract data from browser.
    This is an alternative to file download when SharePoint opens data in browser.
    """
    try:
        logging.info("Navegando para a página de gerenciamento de usuários...")
        driver.get(manage_users_url)
        
        # Wait for page to load
        time.sleep(5)
        
        # Look for Export to Excel button with extensive selectors
        export_button = None
        possible_selectors = [
            "//input[@value='Export to Excel']",
            "//button[contains(text(), 'Export to Excel')]",
            "//a[contains(text(), 'Export to Excel')]",
            "//input[contains(@id, 'Export')]",
            "//input[contains(@name, 'Export')]",
            "//*[contains(text(), 'Export to Excel')]",
            "//span[contains(text(), 'Export to Excel')]/parent::*",
            "//*[@title='Export to Excel']",
            "//input[contains(@onclick, 'Export')]",
            "//*[contains(@onclick, 'Export')]",
            "//a[contains(@href, 'Export')]"
        ]
        
        for selector in possible_selectors:
            try:
                export_button = driver.find_element(By.XPATH, selector)
                logging.info(f"✅ Botão Export encontrado: {selector}")
                break
            except NoSuchElementException:
                continue
        
        if not export_button:
            logging.error("❌ Botão 'Export to Excel' não encontrado")
            return []
        
        # Click the export button
        logging.info("Clicando no botão Export to Excel...")
        original_window = driver.current_window_handle
        export_button.click()
        
        # Wait for data to load (either in same page or new window)
        time.sleep(5)
        
        # Check if new window opened
        windows = driver.window_handles
        target_window = original_window
        
        if len(windows) > 1:
            # Switch to new window
            for window in windows:
                if window != original_window:
                    driver.switch_to.window(window)
                    target_window = window
                    logging.info("🪟 Nova janela detectada, mudando para ela")
                    break
        
        # Extract data from current page
        users_data = _extract_table_data_from_page(driver)
        
        # Switch back to original window if needed
        if target_window != original_window:
            driver.switch_to.window(original_window)
        
        if users_data:
            logging.info(f"✅ Dados extraídos do navegador: {len(users_data)} usuários")
        else:
            logging.warning("⚠️ Nenhum dado de usuário encontrado no navegador")
        
        return users_data
        
    except Exception as e:
        logging.error(f"❌ Erro na extração via navegador: {e}")
        return []


def _extract_table_data_from_page(driver):
    """
    Extract tabular user data from the current page.
    Handles various table formats that SharePoint might use.
    """
    users_data = []
    
    try:
        # Method 1: Look for HTML tables
        tables = driver.find_elements(By.TAG_NAME, "table")
        
        if tables:
            logging.info(f"Encontradas {len(tables)} tabelas na página")
            
            for i, table in enumerate(tables):
                try:
                    # Convert table to pandas DataFrame
                    table_html = table.get_attribute('outerHTML')
                    dfs = pd.read_html(table_html)
                    
                    if dfs:
                        df = dfs[0]
                        logging.info(f"Tabela {i+1}: {len(df)} linhas, {len(df.columns)} colunas")
                        
                        # Check if this looks like user data
                        if _is_user_data_table(df):
                            users_data.extend(_convert_df_to_user_records(df))
                            logging.info(f"✅ Dados de usuário extraídos da tabela {i+1}")
                            
                except Exception as e:
                    logging.warning(f"Erro ao processar tabela {i+1}: {e}")
                    continue
        
        # Method 2: Look for structured data in page source (fallback)
        if not users_data:
            logging.info("Tentando extração alternativa do código fonte...")
            users_data = _extract_from_page_source(driver)
            
    except Exception as e:
        logging.error(f"Erro na extração de dados tabulares: {e}")
    
    return users_data


def _is_user_data_table(df):
    """
    Check if DataFrame contains user data based on column names and content.
    """
    if df.empty:
        return False
    
    # Convert column names to lowercase for comparison
    columns_lower = [str(col).lower() for col in df.columns]
    
    # Look for user-related column indicators
    user_indicators = [
        'name', 'user', 'email', 'account', 'login',
        'status', 'active', 'department', 'title'
    ]
    
    # Check if any columns contain user indicators
    indicator_count = sum(1 for col in columns_lower 
                         for indicator in user_indicators 
                         if indicator in col)
    
    # Also check data content for user patterns
    content_check = False
    try:
        # Convert first few rows to string and check for email patterns
        sample_data = str(df.head(3).to_string()).lower()
        if '@' in sample_data and ('.com' in sample_data or '.org' in sample_data):
            content_check = True
    except:
        pass
    
    # Table likely contains user data if it has user-related columns or email content
    is_user_table = indicator_count >= 2 or content_check
    
    if is_user_table:
        logging.info(f"Tabela identificada como dados de usuário (indicadores: {indicator_count}, emails: {content_check})")
    
    return is_user_table


def _convert_df_to_user_records(df):
    """
    Convert pandas DataFrame to list of user dictionaries.
    """
    users = []
    
    try:
        # Create column mapping
        column_mapping = {}
        for col in df.columns:
            col_lower = str(col).lower()
            
            if any(indicator in col_lower for indicator in ['name', 'full name', 'display name']):
                column_mapping['User Name'] = col
            elif any(indicator in col_lower for indicator in ['email', 'e-mail', 'mail']):
                column_mapping['Email'] = col
            elif any(indicator in col_lower for indicator in ['account', 'login', 'user account']):
                column_mapping['Account'] = col
            elif 'status' in col_lower:
                column_mapping['Status'] = col
            elif any(indicator in col_lower for indicator in ['title', 'job title', 'position']):
                column_mapping['Title'] = col
            elif 'department' in col_lower:
                column_mapping['Department'] = col
        
        # Convert rows to user records
        for _, row in df.iterrows():
            user_record = {}
            
            # Map known columns
            for our_field, df_column in column_mapping.items():
                value = row[df_column]
                if pd.notna(value) and str(value).strip():
                    user_record[our_field] = str(value).strip()
            
            # Add any unmapped columns that might be useful
            for col in df.columns:
                if col not in column_mapping.values():
                    value = row[col]
                    if pd.notna(value) and str(value).strip():
                        # Use original column name for unmapped fields
                        user_record[str(col)] = str(value).strip()
            
            # Only add record if it has meaningful data
            if user_record.get('User Name') or user_record.get('Email') or user_record.get('Account'):
                users.append(user_record)
        
        logging.info(f"Convertidos {len(users)} registros de usuário")
        
    except Exception as e:
        logging.error(f"Erro ao converter DataFrame: {e}")
    
    return users


def _extract_from_page_source(driver):
    """
    Fallback method to extract user data from page source when tables aren't found.
    """
    users_data = []
    
    try:
        page_source = driver.page_source
        
        # Look for email patterns as user indicators
        import re
        email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
        emails = re.findall(email_pattern, page_source)
        
        if emails:
            logging.info(f"Encontrados {len(emails)} emails na página")
            
            # Create basic user records from emails
            seen_emails = set()
            for email in emails:
                if email not in seen_emails and not email.endswith('.js'):  # Filter out JS files
                    seen_emails.add(email)
                    
                    # Create basic user record
                    user_name = email.split('@')[0].replace('.', ' ').title()
                    user_record = {
                        'Email': email,
                        'User Name': user_name,
                        'Account': email,
                        'Status': 'Unknown'  # Can't determine from email alone
                    }
                    users_data.append(user_record)
            
            logging.info(f"Criados {len(users_data)} registros básicos de usuário")
    
    except Exception as e:
        logging.error(f"Erro na extração do código fonte: {e}")
    
    return users_data