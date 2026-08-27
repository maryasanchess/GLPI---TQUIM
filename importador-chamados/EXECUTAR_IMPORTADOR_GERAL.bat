@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if not errorlevel 1 set "PY_LAUNCHER=py -3"
if not defined PY_LAUNCHER (
  where python >nul 2>nul
  if not errorlevel 1 set "PY_LAUNCHER=python"
)
if not defined PY_LAUNCHER goto :sem_python

if not exist ".venv_glpi_geral\Scripts\python.exe" (
  echo Criando o ambiente do importador...
  %PY_LAUNCHER% -m venv .venv_glpi_geral
  if errorlevel 1 goto :erro
)

".venv_glpi_geral\Scripts\python.exe" -c "import pandas, openpyxl" >nul 2>nul
if errorlevel 1 (
  echo Instalando as dependencias no primeiro uso...
  ".venv_glpi_geral\Scripts\python.exe" -m pip install -r requirements.txt
  if errorlevel 1 goto :erro
)

".venv_glpi_geral\Scripts\python.exe" importador_geral.py
exit /b %errorlevel%

:erro
echo.
echo Nao foi possivel preparar o importador.
pause
exit /b 1

:sem_python
echo Python nao foi encontrado. Instale o Python 3 e marque Add Python to PATH.
pause
exit /b 1
