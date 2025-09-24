"""
Test Excel Export Button Functionality
This script will test if the Export to Excel button is found and if it actually downloads a file.
"""
import os
import time
import tempfile
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import glob

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

def create_test_browser():
    """Create a browser instance for testing."""
    options = webdriver.EdgeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--guest")
    
    # Suppress logs
    import sys
    from contextlib import redirect_stderr
    import io
    
    with redirect_stderr(io.StringIO()):
        driver = webdriver.Edge(options=options)
    
    return driver

def configure_browser_downloads(driver, download_dir):
    """Configure browser to download files to specific directory."""
    try:
        driver.execute_cdp_cmd('Page.setDownloadBehavior', {
            'behavior': 'allow',
            'downloadPath': download_dir
        })
        logging.info(f"✅ Configurado diretório de download: {download_dir}")
        return True
    except Exception as e:
        logging.error(f"❌ Erro ao configurar downloads: {e}")
        return False

def wait_for_download(download_dir, timeout=30):
    """Wait for any file to appear in download directory."""
    start_time = time.time()
    
    while time.time() - start_time < timeout:
        files = os.listdir(download_dir)
        if files:
            # Check for Excel files specifically
            excel_files = [f for f in files if f.endswith('.xlsx') or f.endswith('.xls')]
            if excel_files:
                full_path = os.path.join(download_dir, excel_files[0])
                # Check if file is complete (not growing)
                initial_size = os.path.getsize(full_path)
                time.sleep(2)
                final_size = os.path.getsize(full_path)
                if initial_size == final_size and initial_size > 0:
                    logging.info(f"✅ Arquivo baixado: {excel_files[0]} ({final_size} bytes)")
                    return full_path
        
        time.sleep(1)
    
    return None

def test_export_button():
    """Test the Export to Excel functionality step by step."""
    print("🧪 TESTE: Funcionalidade do botão Export to Excel")
    print("=" * 60)
    
    # Create temporary download directory
    temp_dir = tempfile.mkdtemp(prefix="pwa_export_test_")
    logging.info(f"📁 Diretório temporário criado: {temp_dir}")
    
    driver = None
    try:
        # Step 1: Create browser
        print("\n🌐 Passo 1: Criando navegador...")
        driver = create_test_browser()
        logging.info("✅ Browser Edge criado com sucesso")
        
        # Step 2: Configure downloads
        print("\n⚙️ Passo 2: Configurando downloads...")
        download_configured = configure_browser_downloads(driver, temp_dir)
        
        if not download_configured:
            print("❌ Falha na configuração de downloads. Continuando mesmo assim...")
        
        # Step 3: Use known ManageUsers URL
        print("\n🔗 Passo 3: Usando URL conhecida do PWA...")
        manage_users_url = "https://prosperi.sharepoint.com/sites/PWA_VITOR_MASCARENHAS/_layouts/15/pwa/Admin/ManageUsers.aspx"
        print(f"URL: {manage_users_url}")
        
        # Step 4: Navigate to page
        print(f"\n📄 Passo 4: Navegando para: {manage_users_url}")
        driver.get(manage_users_url)
        
        print("⏳ Aguardando carregamento da página...")
        time.sleep(5)  # Wait for page load
        
        print("🔐 Por favor, faça login manualmente se necessário e pressione Enter quando a página estiver carregada...")
        input("Pressione Enter para continuar...")
        
        # Step 5: Look for Export button with detailed analysis
        print("\n🔍 Passo 5: Procurando botão Export to Excel...")
        
        # Extended list of selectors
        selectors = [
            "//input[@value='Export to Excel']",
            "//button[contains(text(), 'Export to Excel')]", 
            "//a[contains(text(), 'Export to Excel')]",
            "//input[contains(@id, 'Export')]",
            "//input[contains(@name, 'Export')]",
            "//button[contains(@class, 'export')]",
            "//*[contains(text(), 'Export to Excel')]",
            "//span[contains(text(), 'Export to Excel')]/parent::*",
            "//*[@title='Export to Excel']",
            "//input[contains(@onclick, 'Export')]",
            "//*[contains(@onclick, 'Export')]",
            "//a[contains(@href, 'Export')]",
            "//input[@value='Export To Excel']",
            "//*[contains(text(), 'Export To Excel')]",
            "//*[contains(text(), 'EXPORT TO EXCEL')]",
            "//*[contains(text(), 'export to excel')]"
        ]
        
        found_button = None
        found_selector = None
        
        for i, selector in enumerate(selectors):
            try:
                button = driver.find_element(By.XPATH, selector)
                found_button = button
                found_selector = selector
                logging.info(f"✅ Botão encontrado com seletor #{i+1}: {selector}")
                break
            except NoSuchElementException:
                continue
        
        if not found_button:
            print("❌ Nenhum botão Export to Excel encontrado!")
            
            # Detailed page analysis
            print("\n🔍 ANÁLISE DA PÁGINA:")
            try:
                # Look for all buttons
                buttons = driver.find_elements(By.TAG_NAME, "button")
                inputs = driver.find_elements(By.TAG_NAME, "input")
                links = driver.find_elements(By.TAG_NAME, "a")
                
                print(f"📊 Elementos encontrados: {len(buttons)} botões, {len(inputs)} inputs, {len(links)} links")
                
                print("\n🔘 BOTÕES na página:")
                for i, btn in enumerate(buttons[:10]):  # First 10 buttons
                    try:
                        text = btn.text or btn.get_attribute('value') or btn.get_attribute('title') or btn.get_attribute('onclick')
                        if text:
                            print(f"   {i+1}. '{text[:50]}...' " if len(str(text)) > 50 else f"   {i+1}. '{text}'")
                    except:
                        pass
                
                print("\n📝 INPUTS na página:")
                for i, inp in enumerate(inputs[:10]):  # First 10 inputs
                    try:
                        value = inp.get_attribute('value') or inp.get_attribute('id') or inp.get_attribute('name')
                        inp_type = inp.get_attribute('type')
                        if value or inp_type:
                            print(f"   {i+1}. Type: {inp_type}, Value/ID/Name: '{value}'")
                    except:
                        pass
                
                # Check page source for export-related text
                page_source = driver.page_source.lower()
                export_terms = ['export', 'excel', 'download', 'save as']
                found_terms = [(term, page_source.count(term)) for term in export_terms if term in page_source]
                
                if found_terms:
                    print(f"\n🔤 Termos relacionados encontrados:")
                    for term, count in found_terms:
                        print(f"   '{term}': {count} ocorrências")
                
            except Exception as e:
                print(f"❌ Erro na análise da página: {e}")
            
            return
        
        # Step 6: Test button click and extract data from browser
        print("\n🖱️ Passo 6: Testando clique no botão...")
        print(f"   Seletor usado: {found_selector}")
        
        # Click the button
        print("   Clicando no botão...")
        found_button.click()
        
        # Wait for page to load/change after click
        print("   Aguardando carregamento da página com dados...")
        time.sleep(5)
        
        # Check if we have data in the browser now
        print("   Verificando se dados foram carregados no navegador...")
        
        try:
            # Look for table data or Excel-like content in the page
            page_source = driver.page_source
            
            # Check if page contains tabular data indicators
            table_indicators = [
                '<table', '<thead', '<tbody', '<tr>', '<td>',
                'excel', 'worksheet', 'workbook', 
                'Name', 'Email', 'User', 'Account'
            ]
            
            found_indicators = []
            for indicator in table_indicators:
                if indicator.lower() in page_source.lower():
                    count = page_source.lower().count(indicator.lower())
                    found_indicators.append((indicator, count))
            
            if found_indicators:
                print("✅ DADOS ENCONTRADOS NO NAVEGADOR!")
                print("   📊 Indicadores de dados tabulares encontrados:")
                for indicator, count in found_indicators[:10]:  # First 10
                    print(f"     - '{indicator}': {count} ocorrências")
                
                # Try to find actual table elements
                try:
                    tables = driver.find_elements(By.TAG_NAME, "table")
                    print(f"   📋 Tabelas HTML encontradas: {len(tables)}")
                    
                    for i, table in enumerate(tables[:3]):  # First 3 tables
                        try:
                            rows = table.find_elements(By.TAG_NAME, "tr")
                            if rows:
                                print(f"     Tabela {i+1}: {len(rows)} linhas")
                                
                                # Try to get header info from first row
                                if len(rows) > 0:
                                    headers = rows[0].find_elements(By.TAG_NAME, "th")
                                    cells = rows[0].find_elements(By.TAG_NAME, "td")
                                    
                                    if headers:
                                        header_texts = [h.text.strip() for h in headers[:5]]
                                        print(f"       Cabeçalhos: {header_texts}")
                                    elif cells:
                                        cell_texts = [c.text.strip() for c in cells[:5]]
                                        print(f"       Primeira linha: {cell_texts}")
                                        
                        except Exception as e:
                            print(f"       Erro ao analisar tabela {i+1}: {e}")
                            
                except Exception as e:
                    print(f"   ⚠️ Erro ao buscar tabelas: {e}")
                
                # Check for specific user data patterns
                user_patterns = ['@', '.com', 'Active', 'Inactive', 'User', 'Account']
                found_patterns = []
                for pattern in user_patterns:
                    count = page_source.count(pattern)
                    if count > 0:
                        found_patterns.append((pattern, count))
                
                if found_patterns:
                    print("   👤 Padrões de dados de usuário encontrados:")
                    for pattern, count in found_patterns:
                        print(f"     - '{pattern}': {count} ocorrências")
                        
            else:
                print("❌ Nenhum dado tabular detectado no navegador")
                print("   Pode ser que o botão tenha aberto uma nova janela/aba")
                
                # Check for new windows/tabs
                windows = driver.window_handles
                print(f"   🪟 Janelas/abas abertas: {len(windows)}")
                
                if len(windows) > 1:
                    print("   Tentando verificar outras janelas...")
                    original_window = driver.current_window_handle
                    
                    for i, window in enumerate(windows):
                        if window != original_window:
                            try:
                                driver.switch_to.window(window)
                                print(f"   📋 Verificando janela {i+1}: {driver.title}")
                                
                                # Quick check for data in new window
                                new_source = driver.page_source
                                if any(ind.lower() in new_source.lower() for ind, _ in table_indicators[:4]):
                                    print("   ✅ Dados encontrados na nova janela!")
                                
                            except Exception as e:
                                print(f"   ⚠️ Erro ao verificar janela {i+1}: {e}")
                    
                    # Switch back to original window
                    driver.switch_to.window(original_window)
                
        except Exception as e:
            print(f"❌ Erro ao verificar dados no navegador: {e}")
        
    except Exception as e:
        logging.error(f"❌ Erro durante o teste: {e}")
        
    finally:
        # Cleanup
        if driver:
            print("\n🧹 Fechando navegador...")
            driver.quit()
        
        # Clean up temp directory
        try:
            import shutil
            shutil.rmtree(temp_dir, ignore_errors=True)
            print(f"🗑️ Diretório temporário removido: {temp_dir}")
        except:
            pass
        
        print("\n✅ Teste concluído!")

if __name__ == "__main__":
    test_export_button()