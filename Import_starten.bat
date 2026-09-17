@echo off
REM Per Doppelklick startbar: legt bei Bedarf die virtuelle Umgebung an, installiert
REM fehlende Python-Pakete, loescht eine vorhandene Zieldatenbank und fuehrt den
REM RWK-Import neu aus.
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Virtuelle Umgebung .venv wird angelegt ...
    where py >nul 2>nul
    if errorlevel 1 (
        python -m venv .venv
    ) else (
        py -3 -m venv .venv
    )
    if not exist ".venv\Scripts\python.exe" (
        echo Anlegen der virtuellen Umgebung fehlgeschlagen. Ist Python installiert und im PATH?
        pause
        exit /b 1
    )
)

".venv\Scripts\python.exe" -c "import pyodbc" >nul 2>nul
if errorlevel 1 (
    echo Benoetigte Python-Pakete werden installiert ...
    ".venv\Scripts\python.exe" -m pip install --upgrade pip >nul
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
    if errorlevel 1 (
        echo Installation der Pakete fehlgeschlagen.
        pause
        exit /b 1
    )
)

".venv\Scripts\python.exe" import_rwk.py --neu
echo.
pause
