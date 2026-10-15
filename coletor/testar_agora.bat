@echo off
REM ================================================
REM  Roda o agente na hora, mostrando tudo na tela.
REM  Use este para testar; o outro e o automatico.
REM ================================================
cd /d "%~dp0"
python agente_vagas.py --debug
echo.
echo ---------------------------------------------
echo Terminou. Aperte qualquer tecla para fechar.
pause > nul
