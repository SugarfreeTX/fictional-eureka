@echo off
setlocal EnableExtensions

rem Project lives in WSL. Task Scheduler is Windows. Run the script
rem inside Ubuntu so it uses that environment / venv / openpyxl.
set "WSL_DISTRO=Ubuntu"
set "WSL_DIR=/home/khemra/projects/budget-manager"
set "LOG=%~dp0savings_tracker.log"

echo [%date% %time%] Starting savings_tracker via WSL>> "%LOG%"

where wsl >nul 2>&1
if errorlevel 1 (
    echo [%date% %time%] ERROR: wsl.exe not found.>> "%LOG%"
    exit /b 1
)

wsl -d %WSL_DISTRO% --cd "%WSL_DIR%" -- bash -lc "if [ -f budget-manager/bin/activate ]; then . .venv/bin/activate; elif [ -f venv/bin/activate ]; then . venv/bin/activate; fi; python savings_tracker.py" >> "%LOG%" 2>&1
set "RC=%ERRORLEVEL%"

echo [%date% %time%] Finished with exit code %RC%>> "%LOG%"
exit /b %RC%
