import os
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

# Lista de arquivos de dados necessários
datas = [('img/logo.webp', 'img')]

block_cipher = None

# Identificar submódulos usados no script main.py
hiddenimports = (
    collect_submodules('cv2') +
    collect_submodules('PIL') +
    collect_submodules('tkinter') +
    [
        'openpyxl', 'numpy', 'pandas', 'fitz', 'platform', 'hashlib',
        'hmac', 'shutil', 'subprocess', 'threading', 'json', 'datetime',
        'io', 're', 'queue', 'concurrent.futures', 'tkinter.ttk', 'tkinter.filedialog',
        'pytesseract'  # Incluindo pytesseract já que aparece em main.py
    ]
)

a = Analysis(
    ['main.py'],  # Arquivo principal do seu projeto
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='PDFAnalyzer',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    icon='img/icon.ico'
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='PDFAnalyzer'
)
