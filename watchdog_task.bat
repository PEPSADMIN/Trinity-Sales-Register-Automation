@echo off
setlocal
REM ============================================================
REM watchdog_task.bat
REM Registered in Windows Task Scheduler to run daily at 7:40 AM,
REM ~40 minutes after run_task.bat (7:00 AM). Deliberately minimal:
REM all the actual logic lives in watchdog.py (in Python) rather
REM than batch, since run_task.bat's own batch-syntax fragility is
REM exactly what caused the incident this watchdog exists to catch.
REM ============================================================

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

set "PYTHON=%SCRIPT_DIR%venv\Scripts\python.exe"
if not exist "%PYTHON%" set "PYTHON=python"

set "PLAYWRIGHT_BROWSERS_PATH=%SCRIPT_DIR%.pw-browsers"

echo [%DATE% %TIME%] Watchdog check starting... >> "%SCRIPT_DIR%run_log.txt"
"%PYTHON%" -u "%SCRIPT_DIR%watchdog.py" >> "%SCRIPT_DIR%run_log.txt" 2>&1
echo [%DATE% %TIME%] Watchdog check finished (exit %ERRORLEVEL%). >> "%SCRIPT_DIR%run_log.txt"
