# -*- mode: python ; coding: utf-8 -*-
# Ejecutable para equipos sin Python: Construir_Ejecutable.bat -> dist\ValidadorRenovaciones\ValidadorRenovaciones.exe
# Se reparte la carpeta completa (comprimida en .zip); el analista la descomprime en su usuario y abre el .exe.

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

datas = collect_data_files("customtkinter") + collect_data_files("tkcalendar") + collect_data_files("babel")
# oracledb carga cryptography desde código compilado (PyInstaller no lo ve): sin ella, el modo thin falla con DPY-3016
hiddenimports = (["win32timezone", "win32cred", "win32clipboard", "babel.numbers", "babel.dates", "PIL.Image",
                  "PIL.ImageOps", "PIL.ImageDraw", "PIL.ImageFont", "lxml.etree", "autoprueba"]
                 + collect_submodules("oracledb") + collect_submodules("cryptography"))

a = Analysis(
    ["validador_renovaciones_app.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=["lxml.isoschematron", "matplotlib", "scipy", "IPython", "notebook", "pytest", "unittest.mock", "PyQt5", "PySide6"],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="ValidadorRenovaciones",
    debug=False,
    strip=False,
    upx=False,
    console=False,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="ValidadorRenovaciones",
)
