@echo off
chcp 65001 >nul
title FONACOT - Validador Diario de Cancelacion de Renovaciones
cd /d "%~dp0"

if exist "dist\FonacotValidadorRenovaciones_v1\FonacotValidadorRenovaciones_v1.exe" (
    start "" "dist\FonacotValidadorRenovaciones_v1\FonacotValidadorRenovaciones_v1.exe"
    exit
)

"C:\Users\sammy.avila\AppData\Local\Programs\Python\Python312\python.exe" validador_renovaciones_app.py
if errorlevel 1 (
    echo.
    echo Ocurrio un error al iniciar la aplicacion.
    pause
)

