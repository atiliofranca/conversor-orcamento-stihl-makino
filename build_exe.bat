@echo off
REM Build script to produce a single .exe using PyInstaller
REM This script will:
REM  - create a local virtual environment in .venv (if missing)
REM  - activate the venv
REM  - upgrade pip and install requirements
REM  - run PyInstaller to create a one-file, windowed executable
REM  - copy the produced exe to the \release folder

setlocal enabledelayedexpansion

echo === Orçamento-Stihl - Build .exe ===

REM Ensure python is available
where python >nul 2>&1
if errorlevel 1 (
  echo Python nao encontrado no PATH. Instale Python 3.8+ e tente novamente.
  pause
  exit /b 1
)

if not exist .venv (
  echo Criando virtualenv em .venv...
  python -m venv .venv
)

echo Ativando venv...
call .venv\Scripts\activate.bat

echo Atualizando pip e instalando dependencias...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo Executando PyInstaller...
REM Use name without non-ascii characters
pyinstaller --clean --onefile --windowed --name orcamento-stihl main.py

if exist dist\orcamento-stihl.exe (
  if not exist release mkdir release
  copy /Y dist\orcamento-stihl.exe release\orcamento-stihl.exe >nul
  echo Executavel criado em release\orcamento-stihl.exe
else
  echo Build falhou. Verifique a saída do PyInstaller acima.
fi

echo Fim.
pause
endlocal
