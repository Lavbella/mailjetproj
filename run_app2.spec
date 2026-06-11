# -*- mode: python ; coding: utf-8 -*-
import os
import importlib.util
from PyInstaller.utils.hooks import copy_metadata, collect_all, collect_submodules

# 1. Identificar o Sistema Operativo
import platform
is_windows = platform.system() == "Windows"

def get_package_path(package_name):
    try:
        spec = importlib.util.find_spec(package_name)
        if spec and spec.origin:
            return os.path.dirname(spec.origin)
    except:
        pass
    return None

# 2. Hidden Imports Universais
hiddenimports = [
    "openpyxl", "openpyxl.cell", "openpyxl.reader.excel",
    "tqdm", "tqdm.auto", "tqdm.std", "tqdm.notebook",
    "jinja2", "jinja2.meta", "markupsafe",
    "docx", "docx.oxml", "docx.oxml.ns",
    "mailjet_rest", "mailjet_rest.client",
    "pkg_resources", "importlib_metadata",
    "streamlit.runtime.scriptrunner", # Necessário para Streamlit moderno
]

# 3. Adicionar imports específicos do Windows apenas se estiver no Windows
if is_windows:
    hiddenimports += ["pythoncom", "win32com", "win32com.client"]
    for mod in collect_submodules("win32com"):
        hiddenimports.append(mod)

# 4. Datas (Ficheiros do projeto)
datas = [
    ("app_v3.py", "."),  
    # Se tiver uma pasta 'templates' ou 'logos', adicione aqui:
    # ("logos", "logos"), 
]

# 5. Coletar tudo dos pacotes principais (removido docx2pdf que já não usamos)
pkgs_to_collect = ["mailjet_rest", "docxtpl", "streamlit", "pandas", "docx", "jinja2", "tqdm", "openpyxl"]

for pkg in pkgs_to_collect:
    hi, d, _ = collect_all(pkg)
    
    # IMPORTANTE para Linux: filtrar apenas nomes de módulos
    # Evita o erro "Invalid hiddenimport ... site-packages/..."
    for item in hi:
        if isinstance(item, str) and "/" not in item and "\\" not in item:
            if item not in hiddenimports:
                hiddenimports.append(item)
    
    datas.extend(d)
    
    pkg_path = get_package_path(pkg)
    if pkg_path:
        if (pkg_path, pkg) not in datas:
            datas.append((pkg_path, pkg))

datas += copy_metadata("streamlit")
datas += copy_metadata("pandas")

# Limpeza final de duplicados
hiddenimports = list(set([str(h) for h in hiddenimports if h]))

a = Analysis(
    ["run_app.py"],
    pathex=["."],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["docx2pdf"], # Excluímos explicitamente para evitar erros no Linux
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="run_app", # Nome do executável
    debug=False,
    strip=False,
    upx=True,
    console=True, 
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=["pytest"],
    name="run_app",
)