@echo off
REM ============================================================
REM register_watchdog_task.bat
REM Registers watchdog_task.bat in Windows Task Scheduler to run
REM daily at 7:40 AM (~40 min after the main 7:00 AM job, giving
REM it time to finish all 6 download retries if the ERP is slow).
REM Runs under the SYSTEM account, same as the main task.
REM
REM Run this ONCE as Administrator.
REM ============================================================

setlocal
set TASK_NAME=Trinity Sales Register Watchdog
set BAT_PATH=%~dp0watchdog_task.bat

schtasks /Create ^
 /TN "%TASK_NAME%" ^
 /TR "%BAT_PATH%" ^
 /SC DAILY ^
 /ST 07:40 ^
 /RU SYSTEM ^
 /RL HIGHEST ^
 /F

REM Fix the action: schtasks /Create mishandles paths with spaces and
REM splits them, causing 0x80070002. Re-set the full path + working
REM directory via COM (same workaround as register_task.bat).
powershell -NoProfile -Command "$d='%~dp0'.TrimEnd('\'); $bat=$d+'\watchdog_task.bat'; $s=New-Object -ComObject Schedule.Service; $s.Connect(); $f=$s.GetFolder('\'); $t=$f.GetTask('%TASK_NAME%'); $def=$t.Definition; foreach($a in $def.Actions){$a.Path=$bat; $a.Arguments=''; $a.WorkingDirectory=$d}; $f.RegisterTaskDefinition('%TASK_NAME%',$def,4,'SYSTEM',$null,5)"

if %ERRORLEVEL% EQU 0 (
    echo.
    echo Task "%TASK_NAME%" scheduled daily at 07:40 (SYSTEM account).
    echo Verify with:  schtasks /Query /TN "%TASK_NAME%"
) else (
    echo.
    echo Failed to create the task. Run this .bat as Administrator.
)
pause
