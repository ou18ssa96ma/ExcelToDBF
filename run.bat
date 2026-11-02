@echo off
setlocal
cd /d "%~dp0"

if exist "%~dp0..\antigravity\.venv\Scripts\python.exe" (
    "%~dp0..\antigravity\.venv\Scripts\python.exe" main.py %1
) else if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" main.py %1
) else (
    python main.py %1
)

if %ERRORLEVEL% neq 0 (
    echo.
    echo Application exited with code %ERRORLEVEL%.
    pause
)

