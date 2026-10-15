@echo off
chcp 65001 >nul
title Adaptador de Curriculo - vaga copiada
cd /d "%~dp0"
echo Usando o texto da vaga que voce copiou (Ctrl+C)...
echo.
python adaptador.py --colar
echo.
start "" "%~dp0saida"
pause
