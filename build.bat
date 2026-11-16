@echo off
REM Build dist\Questionnaire_Converter.exe (standalone, no Python needed to run it).
REM Uses a dedicated virtual environment so only the required packages are bundled.
setlocal
cd /d "%~dp0"

if not exist .venv-build\Scripts\python.exe (
    echo Creating build environment .venv-build ...
    py -3 -m venv .venv-build 2>nul || python -m venv .venv-build || goto :error
)
set PY=.venv-build\Scripts\python.exe

%PY% -m pip install --quiet --upgrade pip || goto :error
%PY% -m pip install --quiet -r requirements.txt pyinstaller || goto :error

%PY% -m PyInstaller --noconfirm --clean --onefile --windowed ^
    --name Questionnaire_Converter ^
    --hidden-import openpyxl ^
    Questionnaire_Converter.py || goto :error

echo.
echo Build OK: dist\Questionnaire_Converter.exe
exit /b 0

:error
echo.
echo Build FAILED.
exit /b 1
