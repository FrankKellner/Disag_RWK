@echo off
REM Per Doppelklick startbar: erzeugt aus den RWK-CSV-Dateien pro eigener
REM Mannschaft eine ICS-Kalenderdatei mit allen Heim- und Auswaertswettkaempfen.
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
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
    if errorlevel 1 (
        echo Installation der Pakete fehlgeschlagen.
        pause
        exit /b 1
    )
)
".venv\Scripts\python.exe" erstelle_ics.py
echo.
pause
