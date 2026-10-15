@echo off
REM ================================================
REM  AGENTE DE VAGAS - execucao diaria automatica
REM
REM  As configuracoes ficam no config.txt.
REM  Este arquivo so dispara o agente e guarda um
REM  registro do que aconteceu em registro.txt.
REM ================================================

cd /d "%~dp0"

echo. >> registro.txt
echo ============================================ >> registro.txt
echo Execucao de %date% as %time% >> registro.txt
echo ============================================ >> registro.txt

python agente_vagas.py >> registro.txt 2>&1

if errorlevel 1 (
    echo *** A EXECUCAO FALHOU *** >> registro.txt
    echo.
    echo Algo deu errado. Abra o registro.txt para ver o motivo.
    timeout /t 20
) else (
    echo Execucao concluida com sucesso. >> registro.txt
)
