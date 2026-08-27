@echo off
REM ============================================================
REM register_task.bat
REM Registers run_task.bat in Windows Task Scheduler to run daily
REM at 7:00 AM. Runs under the SYSTEM account (no password needed,
REM works even when no user is logged on).
REM
REM Run this ONCE as Administrator.
REM Edit the /ST time or /TN name below if needed.
REM ============================================================

setlocal
set TASK_NAME=Trinity Sales Register Email
set BAT_PATH=%~dp0run_task.bat

schtasks /Create ^
 /TN "%TASK_NAME%" ^
 /TR "%BAT_PATH%" ^
 /SC DAILY ^
 /ST 07:00 ^
 /RU SYSTEM ^
 /RL HIGHEST ^
 /F

REM Fix the action: schtasks /Create mishandles paths with spaces and
REM splits them (Command="D:\Hari", Arguments="JR. DATA\..."), causing
REM 0x80070002. Re-set the full path + working directory via COM.
powershell -NoProfile -Command "$d='%~dp0'.TrimEnd('\'); $bat=$d+'\run_task.bat'; $s=New-Object -ComObject Schedule.Service; $s.Connect(); $f=$s.GetFolder('\'); $t=$f.GetTask('%TASK_NAME%'); $def=$t.Definition; foreach($a in $def.Actions){$a.Path=$bat; $a.Arguments=''; $a.WorkingDirectory=$d}; $f.RegisterTaskDefinition('%TASK_NAME%',$def,4,'SYSTEM',$null,5)"

if %ERRORLEVEL% EQU 0 (
    echo.
    echo Task "%TASK_NAME%" scheduled daily at 07:00 (SYSTEM account).
    echo Verify with:  schtasks /Query /TN "%TASK_NAME%"
) else (
    echo.
    echo Failed to create the task. Run this .bat as Administrator.
)
pause
