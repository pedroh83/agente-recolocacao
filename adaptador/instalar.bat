@echo off
chcp 65001 >nul
title Instalar Adaptador de Curriculo
cd /d "%~dp0"
echo ============================================================
echo  INSTALADOR DO ADAPTADOR DE CURRICULO
echo ============================================================
echo.
python --version >nul 2>&1
if errorlevel 1 goto sem_python

echo [1/2] Instalando as bibliotecas (pode levar 1 ou 2 minutos)...
python -m pip install --upgrade --quiet pip
python -m pip install --upgrade --quiet google-genai python-docx
if errorlevel 1 (
  echo.
  echo Nao consegui instalar as bibliotecas. Confira sua internet e rode de novo.
  pause
  exit /b
)
echo       OK
echo.
echo [2/2] Abrindo o assistente de configuracao...
echo.
python adaptador.py --configurar
echo.
pause
exit /b

:sem_python
echo O Python nao foi encontrado neste computador.
echo.
echo Vou abrir a pagina de download. Instale assim:
echo   1. Clique no botao amarelo "Download Python".
echo   2. Abra o arquivo baixado.
echo   3. MARQUE a caixa "Add python.exe to PATH" (embaixo, na primeira tela).
echo   4. Clique em "Install Now" e espere terminar.
echo   5. Feche esta janela e de dois cliques em instalar.bat de novo.
echo.
pause
start "" "https://www.python.org/downloads/"
exit /b
