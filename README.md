# Mailjet Multi-Send: Email Automation, Document Generation & Campaign Tool

---

## 🌍 Language / Idioma
* [Português (#-português)](#-português)
* [English (#-english)](#-english)

---

## 🇵🇹 Português

Esta aplicação profissional desenvolvida em Python oferece uma solução robusta para automatizar o envio de e-mails em massa utilizando a API oficial da **Mailjet**. O ecossistema combina a geração dinâmica de documentos (Word/Template) com base em tabelas de dados e uma interface gráfica moderna para controlo de campanhas.

### Arquitetura da Aplicação
O fluxo de execução do projeto está dividido de forma modular:
* **`run_app.py`**: O ponto de entrada da aplicação. Contém a **interface gráfica interativa** (Streamlit) onde o utilizador faz o upload dos ficheiros, configura os parâmetros e visualiza o progresso.
* **`app_v3.py`**: O motor de execução principal. É chamado diretamente pela interface gráfica para processar os dados, gerar os documentos e realizar o envio em massa através da Mailjet.

### Funcionalidades Principais
A aplicação disponibiliza quatro módulos especializados de atuação:
1. **CSV + PDF Único (BCC):** Permite enviar um único documento PDF em anexo para uma lista de contactos carregada via CSV, utilizando o campo BCC (Cópia Oculta) para proteger a privacidade dos destinatários.
2. **Mail Merge (Word + Excel):** Cruza uma tabela do Excel com um modelo de documento em Word (`template.docx`). A aplicação gera automaticamente um documento personalizado para cada linha do Excel antes de efetuar o envio.
3. **Campanha via Template Mailjet:** Integra-se diretamente com os templates visuais guardados na sua conta Mailjet, automatizando o disparo de campanhas estruturadas.
4. **CSV com corpo de email variável:** Permite que cada e-mail enviado possua um texto e variáveis completamente personalizadas por destinatário, extraídas diretamente das colunas do ficheiro CSV.

### Encerramento Seguro da Aplicação (Processo Python)
* **Botão "Sair" / "Fechar Aplicação":** Foi adicionado um botão específico na interface gráfica que permite **encerrar o processo Python** que controla a execução do servidor. 
* **Fluxo correto:** O utilizador deve **clicar primeiro neste botão** para matar o processo em segundo plano e só depois fechar o separador do browser. 
* *Nota:* Se fechar apenas a janela do navegador sem clicar no botão, o processo Python continuará a correr oculto na sua máquina, consumindo recursos e mantendo a porta de rede ocupada.

> **Dica de Idioma / Tradução:** Embora a interface gráfica integrada esteja originalmente desenvolvida em **Português**, por correr num ambiente web (Streamlit), pode utilizar a funcionalidade de tradução nativa do seu navegador de internet (Google Chrome, Edge, etc.) para traduzir instantaneamente todos os botões e etiquetas (*labels*) para a língua que pretender.

### Componentes & Stack Python (`requirements.txt`)
* **Interface Gráfica:** `streamlit` (Painel web intuitivo).
* **Processamento de Dados:** `pandas` e `openpyxl` (Manipulação de Excel/CSV).
* **Geração de Documentos:** `python-docx` e `docxtpl` (Preenchimento de Word `.docx`).
* **Motor de Templates & Utilitários:** `jinja2`, `markupsafe`, `tqdm` e `importlib-metadata`.
* **Envio de E-mail:** `mailjet-rest` (Integração com a API da Mailjet).

### Pré-requisitos & Conta Mailjet
Para utilizar esta aplicação, **é obrigatório contratar o serviço da Mailjet** para obter as suas credenciais de acesso exclusivas. A aplicação necessita destas chaves (geridas num ficheiro de configuração próprio dentro da pasta `.streamlit`) para autenticar os envios:
* `API_KEY` (Chave pública) e `API_SECRET` (Chave secreta).

---

### 1. Configurar o Ambiente Python Local
Se pretender executar ou testar o código fonte diretamente no seu computador:

```bash
# 1. Criar o ambiente virtual
python -m venv venv

# 2. Ativar o ambiente virtual
# No Windows (PowerShell):
.\venv\Scripts\activate
# No Linux/macOS:
source venv/bin/activate

# 3. Instalar as dependências do projeto
pip install --upgrade pip
pip install -r requirements.txt
```

---

### 2. Como Criar o Executável com PyInstaller (Windows e Linux)
O projeto inclui as configurações avançadas e mapeamento nos ficheiros **`.spec`** (`run_app.spec` e `run_app2.spec`) para empacotar o servidor do Streamlit e os componentes base.

#### Compilação no Windows ou 🐧 Linux:
1. Certifique-se de que o ambiente virtual está ativo.
2. Execute o comando do PyInstaller apontando para a especificação do projeto:
   ```bash
   pyinstaller run_app.spec
   ```

#### Estrutura de Pastas Obrigatória Pós-Compilação
Após a compilação terminar, o PyInstaller irá gerar os ficheiros dentro da diretoria **`dist/run_app/`**. 

Para que a aplicação funcione, a pasta **`_internal`** **deve permanecer intacta na mesma diretoria do executável**. Adicionalmente, deve copiar manualmente para esse mesmo local os recursos visuais e configurações do projeto, resultando na seguinte árvore estrutural:

```text
📂 dist/run_app/ (Subpasta resultante da compilação)
 ├── 📂 _internal/             <-- OBRIGATÓRIA (Gerada pelo PyInstaller, não mover nem apagar)
 ├── 📄 run_app.exe            <-- Executável principal (ou binário Linux)
 ├── 📄 1757057210411.svg      <-- Copiar manualmente para aqui (Recurso Visual)
 ├── 📄 EuroSorte.gif          <-- Copiar manualmente para aqui (Recurso Visual)
 ├── 📂 output/                <-- Criar/Copiar pasta vazia de destino dos relatórios
 └── 📂 .streamlit/            <-- Criar/Copiar esta pasta
      └── 📄 [ficheiro_config] <-- Ficheiro com a sua Key e Secret da Mailjet
```

---

## 🇬🇧 English

This professional Python application provides a robust solution for automating bulk email dispatching using the official **Mailjet API**. The ecosystem combines dynamic document generation (Word template automation) based on structured sheets with a modern graphical user interface for campaign control.

### Application Architecture
The project's execution flow is split modularly:
* **`run_app.py`**: The main entry point of the application. It hosts the **interactive graphical user interface** (Streamlit) where users upload data sheets, configure parameters, and monitor live progress.
* **`app_v3.py`**: The core execution engine. It is triggered by the GUI to process raw data, generate templated documents, and orchestrate bulk dispatches via Mailjet.

### Core Features
The application delivers four specialized operating modules:
1. **CSV + Single PDF (BCC):** Allows broadcasting a single PDF attachment to a contact list imported via CSV, utilizing the BCC (Blind Carbon Copy) field to preserve recipient privacy.
2. **Mail Merge (Word + Excel):** Merges an Excel spreadsheet with a Word template file (`template.docx`). The app automatically compiles a personalized document for each row in the Excel sheet prior to dispatch.
3. **Campaign via Mailjet Template:** Integrates directly with pre-designed visual layouts stored in your Mailjet account, automating structured marketing campaigns.
4. **CSV with dynamic email body:** Ensures each dispatched email features completely tailored copy and custom variables per recipient, extracted live from the CSV file columns.

### Safe Application Shutdown (Python Process Control)
* **"Exit" / "Close App" Button:** A dedicated button has been integrated into the graphical UI to **terminate the backend Python process** orchestrating the execution server.
* **Correct Workflow:** Users must **click this button first** to safely kill the background processes, and only then close the active browser tab.
* *Note:* Simply closing the web browser window without triggering this button will leave a phantom Python process running hidden in your machine, wasting resources and locking the network port.

> **Language / Translation Tip:** Although the built-in graphical user interface is natively written in **Portuguese**, since it runs in a web environment (Streamlit), you can use your internet browser's built-in translation feature (Google Chrome, Edge, etc.) to instantly translate all buttons and labels into your preferred language.

### Components & Python Stack (`requirements.txt`)
* **Graphical UI:** `streamlit` (Intuitive local web dashboard).
* **Data Processing:** `pandas` and `openpyxl` (Excel/CSV spreadsheet manipulation).
* **Document Generation:** `python-docx` and `docxtpl` (Automated population of Word `.docx` templates).
* **Template Engine & Utilities:** `jinja2`, `markupsafe`, `tqdm` and `importlib-metadata`.
* **Email Delivery:** `mailjet-rest` (Official integration wrapper for the Mailjet API).

### Prerequisites & Mailjet Account
To use this application, **you must sign up for a Mailjet service account** to acquire your unique API credentials. The application requires these keys (stored locally inside the `.streamlit` configuration folder) to authenticate your requests:
* `API_KEY` (Public Key) and `API_SECRET` (Secret Key).

---

### 1. Local Python Environment Setup
To run or develop the source code directly on your host machine:

```bash
# 1. Create the virtual environment
python -m venv venv

# 2. Activate the virtual environment
# On Windows (PowerShell):
.\venv\Scripts\activate
# On Linux/macOS:
source venv/bin/activate

# 3. Install all project python packages
pip install --upgrade pip
pip install -r requirements.txt
```

---

### 2. How to Compile the Executable with PyInstaller (Windows & Linux)
The project includes advanced configurations within the **`.spec`** files (`run_app.spec` and `run_app2.spec`) to bundle the Streamlit server orchestration and core components.

#### Compilation on Windows or 🐧 Linux:
1. Ensure your virtual environment is active.

2. Run the PyInstaller build command using the specification file  bash pyinstaller run_app.spec
3. Mandatory Post-Build Folder StructureOnce the compilation finishes, PyInstaller will output the build files inside the dist/run_app/ directory.
4. For the application to work, the _internal folder must remain intact within the same directory as the executable. Additionally, you must manually copy your custom assets and configuration folders to this exact spot, resulting in the following tree layout:
   
📂 dist/run_app/ (Subfolder generated by compilation) 

    ├── 📂 _internal/<-- MANDATORY (Generated by PyInstaller, do not move or delete) 
    ├── 📄 run_app.exe <-- Main standalone executable (or Linux binary) 
    ├── 📄 1757057210411.svg      <-- Copy manually here (Visual Asset) 
    ├── 📄 EuroSorte.gif          <-- Copy manually here (Visual Asset) 
    ├── 📂 output/                <-- Create/Copy empty reports target directory 
    └── 📂 .streamlit/            <-- Create/Copy this folder 
      └── 📄 [config_file]     <-- File containing your Mailjet Key and Secret

License / Licença This software automation workflow is shared under the MIT License.
