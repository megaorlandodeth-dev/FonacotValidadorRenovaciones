@echo off
chcp 65001 >nul
title FONACOT - Validador Diario de Cancelacion de Renovaciones
cd /d "%~dp0"

if exist "dist\FonacotValidadorRenovaciones_v1\FonacotValidadorRenovaciones_v1.exe" (
    start "" "dist\FonacotValidadorRenovaciones_v1\FonacotValidadorRenovaciones_v1.exe"
    exit
)

rem Usa el Python instalado en este equipo (antes apuntaba a la ruta de otro usuario)
set "PYEXE="
where py >nul 2>nul && set "PYEXE=py -3"
if not defined PYEXE (
    where python >nul 2>nul && set "PYEXE=python"
)
if not defined PYEXE (
    echo No se encontro Python. Instala Python 3.10+ desde https://www.python.org/ y marca "Add to PATH".
    pause
    exit /b 1
)

%PYEXE% -c "import customtkinter, tkcalendar, bs4, pandas, openpyxl, win32com" >nul 2>nul
if errorlevel 1 (
    echo Instalando dependencias...
    %PYEXE% -m pip install -r requirements.txt
)

rem Opcionales: lectura de capturas (pillow) y validacion en la base de datos (oracledb). Sin ellos la app funciona igual.
%PYEXE% -c "import PIL, oracledb" >nul 2>nul
if errorlevel 1 (
    echo Instalando complementos opcionales ^(pillow, oracledb^)...
    %PYEXE% -m pip install pillow "oracledb>=3.4,<4"
)

rem oracledb 4 ya no acepta el Oracle Client 12.2 (el de SQL Developer): se usa la 3.4, que funciona con los dos modos
%PYEXE% -c "import oracledb,sys; sys.exit(0 if int(oracledb.__version__.split('.')[0]) < 4 else 1)" >nul 2>nul
if errorlevel 1 (
    echo Instalando la version de oracledb compatible con el Oracle Client de este equipo...
    %PYEXE% -m pip install "oracledb>=3.4,<4"
)

%PYEXE% validador_renovaciones_app.py
if errorlevel 1 (
    echo.
    echo Ocurrio un error al iniciar la aplicacion.
    pause
)
