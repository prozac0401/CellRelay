@echo off
setlocal
cd /d "%~dp0"

if exist ".venv312\Scripts\pythonw.exe" (
    start "" ".venv312\Scripts\pythonw.exe" "main.py"
    exit /b 0
)

if exist ".venv\Scripts\pythonw.exe" (
    start "" ".venv\Scripts\pythonw.exe" "main.py"
    exit /b 0
)

echo CellRelay virtual environment was not found.
echo See README.md for installation instructions.
pause
exit /b 1
