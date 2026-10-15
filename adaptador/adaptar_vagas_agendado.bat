@echo off
chcp 65001 >nul
REM Versao para o Agendador de Tarefas: sem pausa, grava o andamento em log_adaptador.txt
cd /d "%~dp0"
echo ===== %date% %time% ===== >> log_adaptador.txt
python adaptador.py >> log_adaptador.txt 2>&1
