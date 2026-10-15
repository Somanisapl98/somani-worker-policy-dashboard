@echo off
setlocal
cd /d "%~dp0"
if not exist .git (echo ERROR: GitHub is not configured. Run CONFIGURE_WORKER_GITHUB.bat first.& exit /b 1)
copy /y "output\index.html" "index.html" >nul
git add index.html
for /f %%A in ('git diff --cached --name-only ^| find /c /v ""') do set CHANGED=%%A
if "%CHANGED%"=="0" (echo No worker dashboard changes to publish.& exit /b 0)
git commit -m "Update worker policy dashboard"
if errorlevel 1 exit /b 1
git push origin main
if errorlevel 1 exit /b 1
".venv\Scripts\python.exe" cleanup_worker_inputs.py
