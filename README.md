# 📊 Extrator de Dados do PWA

Bem-vindo ao projeto **Extrator de Dados do PWA**, um script automatizado em Python que utiliza Selenium para extrair informações de Grupos, Usuários e Categorias do PWA (Project Web App) e exportá-las em uma planilha Excel.

---

## 🚀 **Funcionalidades Principais**
- ✅ Autenticação automatizada no PWA via navegador Chrome
- ✅ Extração de informações de Grupos, Usuários e Categorias
- ✅ Geração automática de arquivo Excel com os dados
- ✅ Suporte para múltiplos perfis do Chrome

---

## 🛑 **Pré-requisitos**
- Python 3.10 ou superior
- Google Chrome instalado
- WebDriver compatível com a versão do Chrome

---

## ⚙️ **Como Usar**
1. **Criar o ambiente virtual:**
```bash
python -m venv .venv
source .venv/bin/activate  # Linux/macOS
.\.venv\Scripts\Activate.ps1  # Windows (PowerShell)
```
2. **Instalar as dependências:**
```bash
pip install -r requirements.txt
```
3. **Executar o script:**
```bash
python main.py
```
4. **Siga as instruções na tela:**
   - Informe a URL da instância do PWA.
   - Escolha usar um perfil existente do Chrome ou modo convidado.

---

## 📂 **Saída (Output)**
Um arquivo Excel (`pwa_data.xlsx`) será gerado na pasta `Downloads`, contendo:
- **Users:** Lista de usuários associados a cada grupo
- **Groups:** Lista de grupos com nome, descrição e última sincronização
- **Categories:** Lista de categorias associadas a cada grupo

---

## 🛡️ **Segurança**
- O script não armazena credenciais. A autenticação é feita diretamente via navegador Chrome.
- Nenhum dado sensível é gravado fora do arquivo de saída.

---

## 🛠️ **Tecnologias Utilizadas**
- Python 3.10+
- Selenium WebDriver
- BeautifulSoup
- Pandas
- JSON

---



## 📝 **Licença**
Este projeto está licenciado sob a licença MIT. Sinta-se livre para usar e modificar conforme necessário.

---

