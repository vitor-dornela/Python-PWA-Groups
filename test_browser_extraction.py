"""
Simple test for browser-based user extraction
"""
import sys
import os
sys.path.append('src')

from browser_user_extraction import extract_users_from_browser_export
from selenium import webdriver
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

def test_browser_extraction():
    print("🧪 Teste simplificado: Extração de usuários via navegador")
    print("=" * 50)
    
    driver = None
    try:
        # Create browser
        print("🌐 Criando navegador Edge...")
        options = webdriver.EdgeOptions()
        options.add_argument("--start-maximized")
        options.add_argument("--guest")
        
        driver = webdriver.Edge(options=options)
        print("✅ Navegador criado")
        
        # URL do ManageUsers
        manage_users_url = "https://prosperi.sharepoint.com/sites/PWA_VITOR_MASCARENHAS/_layouts/15/pwa/Admin/ManageUsers.aspx"
        
        print(f"\n📄 Navegando para: {manage_users_url}")
        
        # Extract users using browser method
        users_data = extract_users_from_browser_export(driver, manage_users_url)
        
        print(f"\n📊 RESULTADO:")
        print(f"   Usuários encontrados: {len(users_data)}")
        
        if users_data:
            print("✅ SUCESSO! Dados extraídos do navegador")
            
            # Show sample data
            for i, user in enumerate(users_data[:3]):  # First 3 users
                print(f"\n   👤 Usuário {i+1}:")
                for key, value in user.items():
                    print(f"     {key}: {value}")
                    
            # Show column summary
            all_columns = set()
            for user in users_data:
                all_columns.update(user.keys())
            
            print(f"\n   📋 Colunas disponíveis: {sorted(all_columns)}")
        else:
            print("❌ Nenhum usuário encontrado")
            
    except Exception as e:
        print(f"❌ Erro: {e}")
        import traceback
        traceback.print_exc()
        
    finally:
        if driver:
            print("\n🧹 Fechando navegador...")
            driver.quit()
        
    print("\n✅ Teste concluído!")

if __name__ == "__main__":
    test_browser_extraction()