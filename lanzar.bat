@echo off
cd /d "%~dp0"

:: Usar el entorno virtual si existe, sino el Python global
if exist "venv\Scripts\pythonw.exe" (
    start "" "venv\Scripts\pythonw.exe" main.py
) else if exist "venv\Scripts\python.exe" (
    "venv\Scripts\python.exe" main.py
) else (
    pythonw.exe main.py 2>nul || python.exe main.py
)
