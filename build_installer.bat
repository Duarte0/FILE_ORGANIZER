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
if exist "release" rmdir /s /q "release"

python -m PyInstaller --noconfirm "DistribuidorArquivos.spec"

mkdir "release"
mkdir "release\app"

copy /Y "dist\DistribuidorArquivos.exe" "release\app\DistribuidorArquivos.exe" >nul
copy /Y "config.env" "release\app\config.env" >nul
copy /Y "regras.xlsx" "release\app\regras.xlsx" >nul

mkdir "release\app\logs"
mkdir "release\app\relatorios"

set "ISCC_PATH=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not exist "%ISCC_PATH%" set "ISCC_PATH=%ProgramFiles%\Inno Setup 6\ISCC.exe"

if not exist "%ISCC_PATH%" (
    echo Inno Setup nao encontrado. Instale o Inno Setup 6 antes de executar este script.
    pause
    exit /b 1
)

"%ISCC_PATH%" "/DMyAppSource=release\app" "/Orelease" "/FDistribuidorArquivos_Setup" "installer.iss"

echo Instalador gerado com sucesso em release\DistribuidorArquivos_Setup.exe
pause

endlocal