@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  py -m venv .venv
  ".venv\Scripts\python.exe" -m pip install openpyxl
)
".venv\Scripts\python.exe" build_worker_dashboard.py
if errorlevel 1 exit /b 1
call PUBLISH_WORKER_DASHBOARD.bat
