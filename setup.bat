@echo off
rem Ekxel setup: makes .venv next to this file, installs the libraries into it, then checks everything.
rem Safe to run again. Nothing is installed system-wide.
cd /d "%~dp0"
set "PIP_CACHE_DIR=%~dp0.pip-cache"
where python >nul 2>nul || (echo Python not found. Install Python 3.11+ ^(winget install Python.Python.3.12^), then run this again. & exit /b 1)
if not exist ".venv\Scripts\python.exe" (
    echo Creating .venv ...
    python -m venv .venv || (echo Could not create .venv. & exit /b 1)
)
echo Installing libraries ...
".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -q -r requirements.txt || (echo pip install failed. & exit /b 1)
".venv\Scripts\python.exe" tools\doctor.py
exit /b %errorlevel%
