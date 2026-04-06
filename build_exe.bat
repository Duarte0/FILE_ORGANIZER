@echo off
setlocal

cd /d "%~dp0"

if not exist ".venv" (
    py -3 -m venv .venv
)

call ".venv\Scripts\activate.bat"

python -m pip install --upgrade pip
python -m pip install -r requirements.txt

if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"

python -m PyInstaller --noconfirm "DistribuidorArquivos.spec"

copy /Y "config.env" "dist\config.env" >nul
copy /Y "regras.xlsx" "dist\regras.xlsx" >nul

if not exist "dist\relatorios" mkdir "dist\relatorios"
if not exist "dist\logs" mkdir "dist\logs"

echo Build concluido! O executavel final esta em dist\DistribuidorArquivos.exe
pause

endlocal