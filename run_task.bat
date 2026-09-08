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

set "LOGFILE=%SCRIPT_DIR%run_log.txt"

REM Sync the scripts to the latest committed/pushed code before running.
REM This is what actually closes off the Sep-06/07-2026 stale-file
REM incident: that bug shipped because a local, uncommitted edit on disk
REM had silently reverted an already-fixed function. Resetting to
REM origin/main here means the 7 AM run always executes the real,
REM reviewed version — a forgotten local edit can no longer diverge from
REM it. config.env/downloads/venv are all gitignored, so this doesn't
REM touch secrets or data. Best-effort: if offline, just run with
REM whatever is already on disk rather than failing the whole job.
set "GIT=C:\Program Files\Git\cmd\git.exe"
if not exist "%GIT%" set "GIT=git"

REM NOTE: no "-C %SCRIPT_DIR%" here — SCRIPT_DIR ends in a trailing
REM backslash (from %~dp0), and "\"" right before the closing quote is
REM parsed by cmd.exe as an escaped quote, not a closing one. That
REM swallowed the rest of the line into one argument and crashed this
REM script outright on 08-Sep-2026 (no download, no email sent at all).
REM We already "cd /d" into SCRIPT_DIR above, so plain git commands
REM already operate on the right repo — no -C needed.
REM Plain if/goto, not nested if/else(...) blocks — this file already
REM uses goto for the download-retry loop, and nested parenthesized
REM if/else blocks are fragile in cmd.exe (easy to trip a silent
REM "... was unexpected at this time" parse failure once a file has
REM several such blocks). goto keeps each step a single flat line.
echo [%DATE% %TIME%] Syncing to origin/main... >> "%LOGFILE%"
"%GIT%" fetch origin main --quiet >> "%LOGFILE%" 2>&1
if not !ERRORLEVEL! EQU 0 goto :sync_fetch_failed
"%GIT%" reset --hard origin/main --quiet >> "%LOGFILE%" 2>&1
if not !ERRORLEVEL! EQU 0 goto :sync_reset_failed
echo [%DATE% %TIME%] Synced to origin/main. >> "%LOGFILE%"
goto :sync_done

:sync_fetch_failed
echo [%DATE% %TIME%] WARNING: git fetch failed (offline?); running with existing local files. >> "%LOGFILE%"
goto :sync_done

:sync_reset_failed
echo [%DATE% %TIME%] WARNING: git reset failed; running with existing local files. >> "%LOGFILE%"

:sync_done

REM Use the self-contained venv (works under SYSTEM; avoids per-user
REM site-packages / Playwright browser-cache visibility issues).
set "PYTHON=%SCRIPT_DIR%venv\Scripts\python.exe"
if not exist "%PYTHON%" set "PYTHON=python"

REM Playwright browsers live in the project folder (SYSTEM-readable)
set "PLAYWRIGHT_BROWSERS_PATH=%SCRIPT_DIR%.pw-browsers"

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
