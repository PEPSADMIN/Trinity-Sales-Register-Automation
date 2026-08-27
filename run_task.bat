@echo off
setlocal enabledelayedexpansion
REM ============================================================
REM run_task.bat
REM Trinity Sales Register - download + email, no WhatsApp.
REM Registered in Windows Task Scheduler to run daily at 7:00 AM.
REM Retries since the ERP login/page-load is flaky (no one watching).
REM
REM Hardened: all paths are absolute (script dir) and python is
REM located explicitly, so the job works under SYSTEM regardless of
REM the Task Scheduler "Start In" / working directory.
REM ============================================================

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

REM Use the self-contained venv (works under SYSTEM; avoids per-user
REM site-packages / Playwright browser-cache visibility issues).
set "PYTHON=%SCRIPT_DIR%venv\Scripts\python.exe"
if not exist "%PYTHON%" set "PYTHON=python"

REM Playwright browsers live in the project folder (SYSTEM-readable)
set "PLAYWRIGHT_BROWSERS_PATH=%SCRIPT_DIR%.pw-browsers"

set "LOGFILE=%SCRIPT_DIR%run_log.txt"
set "MAX_ATTEMPTS=6"

echo [%DATE% %TIME%] Starting Trinity Sales Register (download + email)... >> "%LOGFILE%"

for /L %%A in (1,1,%MAX_ATTEMPTS%) do (
    echo [%DATE% %TIME%] Download attempt %%A of %MAX_ATTEMPTS%... >> "%LOGFILE%"
    "%PYTHON%" -u "%SCRIPT_DIR%automate_report.py" >> "%LOGFILE%" 2>&1
    if !ERRORLEVEL! EQU 0 (
        echo [%DATE% %TIME%] Download completed successfully on attempt %%A. >> "%LOGFILE%"
        goto :download_done
    )
    echo [%DATE% %TIME%] Download attempt %%A failed with error code !ERRORLEVEL!. >> "%LOGFILE%"
)

echo [%DATE% %TIME%] All %MAX_ATTEMPTS% download attempts failed. >> "%LOGFILE%"
echo [%DATE% %TIME%] Sending error alert email to hariit@pepsindia.com... >> "%LOGFILE%"
"%PYTHON%" -u "%SCRIPT_DIR%email_sender.py" --error "Automated Trinity Sales Register download failed after %MAX_ATTEMPTS% attempts. The ERP server may be slow, busy, or temporarily unavailable. Please check the ERP connection and re-run, or generate the report manually." --error-file "%SCRIPT_DIR%last_error.txt" >> "%LOGFILE%" 2>&1
if !ERRORLEVEL! EQU 0 (
    echo [%DATE% %TIME%] Error alert email sent. >> "%LOGFILE%"
) else (
    echo [%DATE% %TIME%] Error alert email FAILED to send. >> "%LOGFILE%"
)
goto :done

:download_done
echo [%DATE% %TIME%] Emailing downloaded Sales Register (no edits)... >> "%LOGFILE%"
"%PYTHON%" -u "%SCRIPT_DIR%email_sender.py" >> "%LOGFILE%" 2>&1
if !ERRORLEVEL! EQU 0 (
    echo [%DATE% %TIME%] Email sent. >> "%LOGFILE%"
) else (
    echo [%DATE% %TIME%] Email FAILED with error code !ERRORLEVEL!. >> "%LOGFILE%"
)

:done
