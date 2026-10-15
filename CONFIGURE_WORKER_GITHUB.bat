@echo off
setlocal
cd /d "%~dp0"
set /p REPO=Paste the NEW worker dashboard GitHub repository URL: 
if "%REPO%"=="" exit /b 1
if exist .git rmdir /s /q .git
git init
git branch -M main
git remote add origin "%REPO%"
git add .gitignore index.html 2>nul
git commit -m "Initial worker dashboard"
git push -u origin main
