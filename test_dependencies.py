#!/usr/bin/env python3
"""
Teste das dependências do projeto PWA-Extractor
"""

def test_imports():
    """Testa se todas as dependências necessárias estão instaladas"""
    try:
        print("🔍 Testando importações...")
        
        # Dependências principais
        import selenium
        print(f"✅ Selenium {selenium.__version__}")
        
        import pandas
        print(f"✅ Pandas {pandas.__version__}")
        
        import bs4
        print(f"✅ BeautifulSoup4 {bs4.__version__}")
        
        import numpy
        print(f"✅ NumPy {numpy.__version__}")
        
        import openpyxl
        print(f"✅ openpyxl {openpyxl.__version__}")
        
        # Teste específico do Selenium WebDriver
        from selenium import webdriver
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.support.ui import WebDriverWait
        print("✅ Selenium WebDriver imports OK")
        
        # Teste do BeautifulSoup
        from bs4 import BeautifulSoup
        print("✅ BeautifulSoup imports OK")
        
        print("\n🎉 Todas as dependências foram instaladas com sucesso!")
        return True
        
    except ImportError as e:
        print(f"❌ Erro na importação: {e}")
        return False
    except Exception as e:
        print(f"❌ Erro inesperado: {e}")
        return False

if __name__ == "__main__":
    test_imports()