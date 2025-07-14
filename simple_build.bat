@echo off
echo ========================================
echo Simple PyInstaller Build Script
echo ========================================
echo.

REM Check Python installation
python --version
if errorlevel 1 (
    echo ERROR: Python not found in PATH!
    pause
    exit /b 1
)

REM Install PyInstaller if needed
echo Checking PyInstaller...
python -m pip install pyinstaller

REM Install requirements
echo Installing requirements...
python -m pip install -r requirements.txt

REM Clean previous builds
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"
if exist "*.spec" del "*.spec"

REM Build with PyInstaller using module syntax
echo.
echo Building executable...
python -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name ExcelSummaryMaker ^
    --add-data "src;src" ^
    --hidden-import tkinter ^
    --hidden-import pandas ^
    --hidden-import openpyxl ^
    --hidden-import xlsxwriter ^
    --clean ^
    main.py

if errorlevel 1 (
    echo Build failed!
    pause
    exit /b 1
)

echo.
echo Build completed successfully!
echo Executable: dist\ExcelSummaryMaker.exe
pause
