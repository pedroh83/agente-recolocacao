@echo off
chcp 65001 >nul
title Adaptador de Curriculo - vagas da pasta
cd /d "%~dp0"
python adaptador.py
echo.
start "" "%~dp0saida"
pause
