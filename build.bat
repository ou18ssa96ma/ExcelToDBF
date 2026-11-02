
@echo off
title Excel to DBF - Windows Build

echo =====================================
echo       Excel to DBF - Build
echo =====================================

if not exist ".venv\Scripts\python.exe" (
    echo ERROR: Virtual environment not found.
    pause
    exit /b 1
)

echo.
echo Installing dependencies...
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 goto error

echo.
echo Installing PyInstaller...
.venv\Scripts\python.exe -m pip install pyinstaller
if errorlevel 1 goto error

echo.
echo Building application...

.venv\Scripts\python.exe -m PyInstaller ^
    --noconfirm ^
    --clean ^
    --onefile ^
    --windowed ^
    --name ExcelToDBF ^
    main.py

if errorlevel 1 goto error

echo.
echo =====================================
echo BUILD SUCCESSFUL
echo Output: dist\ExcelToDBF\ExcelToDBF.exe
echo =====================================
pause
exit /b 0

:error
echo.
echo BUILD FAILED!
pause
exit /b 1
