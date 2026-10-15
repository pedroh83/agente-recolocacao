@echo off
chcp 65001 >nul
title Teste do Adaptador de Curriculo
cd /d "%~dp0"
python adaptador.py --teste
echo.
pause
