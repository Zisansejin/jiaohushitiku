import os
import sys

a = Analysis(
    ["start_app.py"],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        "streamlit",
        "streamlit.web.cli",
        "sqlite3",
        "pypdf",
        "docx",
        "openpyxl",
        "reportlab",
        "pandas",
        "requests",
        "markdown",
        "dotenv",
        "json",
        "datetime",
        "io"
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=["pathlib"], # 主动排除冲突包
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="药学规培AI题库生成工具",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=False, # 隐藏黑窗口，调试改为True
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)