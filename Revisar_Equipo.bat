@echo off
chcp 65001 >nul
rem Revisa que este equipo tenga todo para el Validador (Outlook, OCR, PDF, conexion a la base...). Deja autoprueba.txt.
cd /d "%~dp0"
start "" "%~dp0ValidadorRenovaciones.exe" --autoprueba
