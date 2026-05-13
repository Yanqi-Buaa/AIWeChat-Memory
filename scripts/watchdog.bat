@echo off
chcp 65001 >nul
title CowAgent Watchdog

set "ROOT=%~dp0.."
set "PID_FILE=%ROOT%\.cow.pid"
set "NOTIFY_FILE=%ROOT%\scripts\pending_notify.json"
set "PYTHON=python"
set "MAX_WAIT=60"

echo [Watchdog] Triggered at %date% %time%

:: Step 1: Check if process is alive
set "ALIVE=0"
if exist "%PID_FILE%" (
    set /p PID=<"%PID_FILE%"
    tasklist /FI "PID eq %PID%" /NH 2>nul | findstr "%PID%" >nul
    if not errorlevel 1 set "ALIVE=1"
)

if "%ALIVE%"=="1" (
    echo [Watchdog] Process PID=%PID% is alive, skipping restart
    goto :CLEANUP
)

:: Step 2: Process is dead, restart it
echo [Watchdog] Process is not running, starting CowAgent...
cd /d "%ROOT%"
start /B "" "%PYTHON%" app.py >nohup.out 2>&1

:: Step 3: Wait for process to come alive
set /a WAITED=0
:WAIT_LOOP
if %WAITED% geq %MAX_WAIT% (
    echo [Watchdog] Timeout waiting for process to start
    goto :CLEANUP
)
set /a WAITED+=1
timeout /t 1 /nobreak >nul

if exist "%PID_FILE%" (
    set /p NEWPID=<"%PID_FILE%"
    tasklist /FI "PID eq %NEWPID%" /NH 2>nul | findstr "%NEWPID%" >nul
    if not errorlevel 1 (
        echo [Watchdog] Process started successfully (PID=%NEWPID%)
        goto :WAIT_READY
    )
)
goto :WAIT_LOOP

:: Step 4: Wait a bit for agent to initialize, then send notification
:WAIT_READY
echo [Watchdog] Waiting for agent to initialize...
timeout /t 15 /nobreak >nul

:: Send pending notification via the Python script
if exist "%NOTIFY_FILE%" (
    echo [Watchdog] Sending pending notification...
    "%PYTHON%" "%SCRIPT_DIR%\send_notify.py"
    if errorlevel 1 (
        echo [Watchdog] Warning: Notification send script returned error
    ) else (
        echo [Watchdog] Notification processed
    )
) else (
    echo [Watchdog] No pending notification found
)

:: Step 5: Cleanup
:CLEANUP
echo [Watchdog] Cleaning up watchdog task...
schtasks /delete /tn "CowAgent_Watchdog" /f >nul 2>&1
echo [Watchdog] Watchdog finished at %date% %time%
