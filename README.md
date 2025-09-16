# 📊 Extrator de Dados do PWA

Bem-vindo ao projeto **Extrator de Dados do PWA**, um script automatizado em Python que utiliza Selenium para extrair informações de Grupos, Usuários e Categorias do PWA (Project Web App) e exportá-las em uma planilha Excel com tabelas formatadas.

---

## 🚀 **Funcionalidades Principais**
- ✅ Autenticação automatizada no PWA via navegador Chrome
- ✅ Extração de informações de Grupos, Usuários e Categorias
- ✅ Geração automática de arquivo Excel com tabelas nativas
- ✅ Filtros automáticos e formatação profissional
- ✅ Mensagens de progresso durante a extração
- ✅ Suporte para modo convidado do Chrome

---

## 🛑 **Pré-requisitos**

- **Python 3.11 ou superior**
- **Google Chrome** instalado
- **Acesso ao PWA** (Project Web App) do SharePoint

> **💡 Nota:** O WebDriver do Chrome é gerenciado automaticamente pelo Selenium.

---

## ⚙️ **Como Usar**

### 1. **Criar o ambiente virtual:**
```bash
# Criar o ambiente virtual (execute apenas uma vez)
python -m venv .venv
```

### 2. **Ativar o ambiente virtual:**

**Windows (PowerShell):**
```powershell
.venv\Scripts\Activate.ps1
```

**Windows (Command Prompt):**
```cmd
.\.venv\Scripts\activate.bat
```

**Linux/macOS:**
```bash
source .venv/bin/activate
```

### 3. **Instalar as dependências:**
```bash
pip install -r requirements.txt
```

### 4. **Executar o script:**
```bash
python main.py
```

### 5. **Siga as instruções na tela:**
   - Informe a URL da instância do PWA
   - Complete o login no navegador que será aberto automaticamente
   - Aguarde a extração dos dados

> **💡 Dica:** Para execuções futuras, você só precisa repetir os passos 2, 4 e 5.

---

## 📦 **Gerar Executável (Opcional)**

Para criar um arquivo executável (.exe) que não requer Python instalado:

### 1. **Instalar PyInstaller:**
```bash
pip install pyinstaller
```

### 2. **Gerar o executável:**
```bash
pyinstaller --onefile --console --name="PWA_EXTRACTOR" main.py
```

### 3. **Localizar o executável:**
- O arquivo `PWA_EXTRACTOR.exe` será criado na pasta `dist/`
- Copie este arquivo para qualquer computador Windows
- Execute com duplo clique (não requer Python instalado)

> **💡 Vantagens do executável:**
> - ✅ Não requer Python instalado no computador de destino
> - ✅ Todas as dependências incluídas automaticamente
> - ✅ Facilita distribuição em ambientes corporativos
> - ✅ Execução simples com duplo clique

---

## 📂 **Saída (Output)**

Um arquivo Excel (`pwa_data.xlsx`) será gerado na pasta `Downloads`, contendo tabelas formatadas com:

- **Users:** Lista de usuários associados a cada grupo
- **Groups:** Lista de grupos com nome, descrição e última sincronização  
- **Categories:** Lista de categorias associadas a cada grupo

> **💡 Nota:** O arquivo Excel é gerado com tabelas nativas, incluindo filtros automáticos e formatação profissional.

---

## 🛡️ **Segurança**

- O script não armazena credenciais. A autenticação é feita diretamente via navegador Chrome.
- Nenhum dado sensível é gravado fora do arquivo de saída.

---

## 🛠️ **Tecnologias Utilizadas**

- **Python 3.11+**
- **Selenium WebDriver** - Automação do navegador
- **BeautifulSoup4** - Parsing de HTML
- **Pandas** - Manipulação de dados
- **openpyxl** - Criação de arquivos Excel com tabelas
- **psutil** - Gerenciamento de processos do Chrome

---

## 📝 **Licença**

Este projeto está licenciado sob a licença MIT. Sinta-se livre para usar e modificar conforme necessário.

---

