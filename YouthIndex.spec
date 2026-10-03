# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec for Youth Consumer Spend Index.
Works on both macOS (produces .app) and Windows (produces folder with .exe).
Not used by Streamlit Cloud.
"""

from PyInstaller.utils.hooks import collect_all, copy_metadata

datas = []
binaries = []
hiddenimports = []

for pkg in [
    "streamlit",
    "yfinance",
    "pandas",
    "numpy",
    "plotly",
    "altair",
    "pyarrow",
    "docx",
    "kaleido",
]:
    try:
        d, b, h = collect_all(pkg)
        datas += d
        binaries += b
        hiddenimports += h
    except Exception:
        pass

for pkg in ["streamlit", "yfinance", "pandas", "numpy", "plotly", "python-docx", "kaleido"]:
    try:
        datas += copy_metadata(pkg)
    except Exception:
        pass

datas += [("app.py", ".")]

block_cipher = None

a = Analysis(
    ["run_app.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports + [
        "streamlit.web.cli",
        "streamlit.runtime.scriptrunner.magic_funcs",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="YouthIndex",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="YouthIndex",
)

app = BUNDLE(
    coll,
    name="Youth Consumer Spend Index.app",
    icon=None,
    bundle_identifier="com.youthindex.app",
    info_plist={
        "CFBundleName": "Youth Consumer Spend Index",
        "CFBundleDisplayName": "Youth Consumer Spend Index",
        "CFBundleShortVersionString": "1.0.0",
        "CFBundleVersion": "1.0.0",
        "NSHighResolutionCapable": True,
        "LSBackgroundOnly": False,
    },
)
