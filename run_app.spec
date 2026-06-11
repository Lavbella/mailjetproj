# -*- mode: python ; coding: utf-8 -*-
import os
import importlib.util
from PyInstaller.utils.hooks import copy_metadata, collect_all, collect_submodules

def get_package_path(package_name):
    try:
        spec = importlib.util.find_spec(package_name)
        if spec and spec.origin:
            return os.path.dirname(spec.origin)
    except:
        pass
    return None

hiddenimports = [
    "openpyxl",
    "openpyxl.cell", 
    "openpyxl.reader.excel",
    "tqdm",
    "tqdm.auto",
    "tqdm.std",
    "tqdm.notebook",
    "tqdm.gui",
    "tqdm.cli",
    "jinja2",
    "jinja2.meta",
    "markupsafe",
    "docx",
    "docx.oxml",
    "docx.oxml.ns",
    "mailjet_rest",
    "mailjet_rest.client",
    "pythoncom",
    "win32com",
    "win32com.client",
    "pkg_resources",
    "importlib_metadata",
]

# ADICIONE O SEU APP.PY AQUI
datas = [
    ("app_v2.py", "."),  # Copia o app.py para a raiz do executável
]

for pkg in ("mailjet_rest", "docxtpl", "docx2pdf", "streamlit", "pandas", "docx", "jinja2", "tqdm", "openpyxl"):
    hi, d, _ = collect_all(pkg)
    for item in hi:
        if isinstance(item, str):
            hiddenimports.append(item)
    datas += d
    
    pkg_path = get_package_path(pkg)
    if pkg_path:
        datas.append((pkg_path, pkg))

datas += copy_metadata("streamlit")
datas += copy_metadata("pandas")
datas += copy_metadata("docx2pdf")

for mod in collect_submodules("win32com"):
    hiddenimports.append(mod)

hiddenimports = list(set([str(h) for h in hiddenimports if isinstance(h, str)]))

a = Analysis(
    ["run_app.py"],
    pathex=["."],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="run_app",
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
    upx_exclude=[],
    name="run_app",
)