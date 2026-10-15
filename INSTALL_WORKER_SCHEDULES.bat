@echo off
setlocal
set ROOT=%~dp0
set ROOT=%ROOT:~0,-1%
schtasks /Create /TN "Somani Worker Policy Dashboard Morning" /TR "cmd.exe /c \"%ROOT%\UPDATE_WORKER_DASHBOARD.bat\"" /SC DAILY /ST 10:30 /F
schtasks /Create /TN "Somani Worker Policy Dashboard Evening" /TR "cmd.exe /c \"%ROOT%\UPDATE_WORKER_DASHBOARD.bat\"" /SC DAILY /ST 17:05 /F
echo Worker dashboard schedules created at 10:30 AM and 5:05 PM.
pause
