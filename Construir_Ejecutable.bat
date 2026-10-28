@echo off
chcp 65001 >nul
title FONACOT - Construir el ejecutable del Validador
cd /d "%~dp0"

rem Solo para quien reparte la app: arma dist\ValidadorRenovaciones (con Python y todas las librerias adentro) y
rem ValidadorRenovaciones.zip. Los analistas no necesitan Python: descomprimen el .zip y abren ValidadorRenovaciones.exe.
set "PYEXE="
where py >nul 2>nul && set "PYEXE=py -3"
if not defined PYEXE (
    where python >nul 2>nul && set "PYEXE=python"
)
if not defined PYEXE (
    echo Para construir el ejecutable se necesita Python 3.10+ en este equipo ^(solo en este^).
    pause
    exit /b 1
)

echo Instalando lo necesario para construir...
%PYEXE% -m pip install -r requirements.txt "pyinstaller>=6.10,<7"
if errorlevel 1 goto error

%PYEXE% -m PyInstaller ValidadorRenovaciones.spec --noconfirm --clean
if errorlevel 1 goto error

copy /y GUIA_DE_USO.md dist\ValidadorRenovaciones\ >nul
if exist ValidadorRenovaciones.zip del ValidadorRenovaciones.zip
powershell -NoProfile -Command "Compress-Archive -Path 'dist\ValidadorRenovaciones' -DestinationPath 'ValidadorRenovaciones.zip'"
if errorlevel 1 goto error

echo.
echo Listo: ValidadorRenovaciones.zip
echo Cada analista lo descomprime en Documentos o en el Escritorio ^(no en Archivos de programa^) y abre
echo ValidadorRenovaciones\ValidadorRenovaciones.exe
pause
exit /b 0

:error
echo.
echo No se pudo construir el ejecutable. Revisa los mensajes de arriba.
pause
exit /b 1
