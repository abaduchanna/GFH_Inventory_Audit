# -*- mode: python ; coding: utf-8 -*-
import datetime as _dt
_year = _dt.date.today().year

SPEC_DOC = f"""PyInstaller spec
Developed by www.3SVerse.com \u00a9 {_year}
Build command: pyinstaller GFH_Inventory_Audit.spec
"""


block_cipher = None

a = Analysis(
    ['GFH_Inventory_Audit.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('native_titlebar.py', '.'),
        ('assets', 'assets'),
        ('GFH_Telecom_Logo.png', '.'),
        ('GFH_icon.ico', '.'),
        ('stores.json', '.'),
        ('theme_manager.py', '.'),
        ('logo_handler.py', '.'),
        ('header_manager.py', '.'),
    ],
    hiddenimports=[
        'tkinter',
        '_tkinter',
        'theme_manager',
        'logo_handler',
        'PIL',
        'openpyxl',
        'pyautogui',
    ],
    hookspath=[],
    hooksconfig={
        'matplotlib': {'backends': ''},
    },
    runtime_hooks=[],
    excludedimports=[
        'doctest',
        'pdb',
        'selenium',
        'pandas',
        'numpy',
        'pyperclip',
        'webdriver_manager',
        'pdfplumber',
        'pypdf',
        'requests',
        'torch',
        'torchvision',
        'torchaudio',
        'matplotlib',
        'matplotlib.pyplot',
        'numba',
        'llvmlite',
        'sympy',
        'tensorflow',
        'scipy',
        'sklearn',
        'scikit-learn',
        'gi',
        'pygments',
        'fsspec',
        'tensorboard',
        'IPython',
        'ipython',
        'jupyter',
        'notebook',
        'PyQt6',
        'PyQt5',
        'PySide6',
        'PySide2',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)


splash = Splash(
    'splash.png',
    binaries=a.binaries,
    datas=a.datas,
    text_pos=(60, 316),
    text_size=11,
    text_color='#96a8cd',
    text_default='Starting...',
    minify_script=True,
    always_on_top=True,
)

exe = EXE(
    pyz,
    a.scripts,
    splash,
    splash.binaries,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='GFH_Inventory_Audit',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='GFH_icon.ico',
)
