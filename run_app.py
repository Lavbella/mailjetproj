import streamlit.web.cli as stcli
import os
import sys
import webbrowser
from threading import Timer

def get_resource_path(relative_path):
    """ Obtém o caminho correto para recursos dentro do executável """
    if getattr(sys, 'frozen', False):
        # Caminho interno onde o PyInstaller extrai os ficheiros
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

def open_browser():
    webbrowser.open_new("http://localhost:8501")

if __name__ == "__main__":
    # 1. Localizar o app.py corretamente
    path_to_app = get_resource_path("app_v3.py")
    
    # 2. Definir o diretório de trabalho para a pasta do .exe (para ler CSVs/XLSX externos)
    if getattr(sys, 'frozen', False):
        os.chdir(os.path.dirname(sys.executable))

    # 3. COMENTE estas linhas para conseguir ver o erro real no terminal
    # sys.stdout = open(os.devnull, 'w')
    # sys.stderr = open(os.devnull, 'w')

    # 4. Abre o browser
    Timer(3, open_browser).start()

    # 5. Configuração do Streamlit
    sys.argv = [
        "streamlit",
        "run",
        path_to_app,
        "--server.port=8501",
        "--server.headless=true",
        "--browser.gatherUsageStats=false",
        "--global.developmentMode=false",
    ]
    
    sys.exit(stcli.main())