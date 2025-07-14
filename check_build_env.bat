@echo off
echo ========================================
echo PyInstaller Environment Check
echo ========================================
echo.

echo 1. Checking Python...
python --version
if errorlevel 1 (
    echo ERROR: Python not found!
    goto :error
)

echo.
echo 2. Checking pip...
python -m pip --version
if errorlevel 1 (
    echo ERROR: pip not working!
    goto :error
)

echo.
echo 3. Checking PyInstaller installation...
python -m pip show pyinstaller
if errorlevel 1 (
    echo PyInstaller not installed, installing now...
    python -m pip install pyinstaller
    if errorlevel 1 (
        echo ERROR: Failed to install PyInstaller!
        goto :error
    )
)

echo.
echo 4. Testing PyInstaller command...
python -m PyInstaller --version
if errorlevel 1 (
    echo ERROR: PyInstaller not working!
    goto :error
)

echo.
echo 5. Checking required packages...
python -c "import tkinter; print('tkinter: OK')"
python -c "import pandas; print('pandas: OK')"
python -c "import openpyxl; print('openpyxl: OK')"
python -c "import xlsxwriter; print('xlsxwriter: OK')"

echo.
echo ========================================
echo All checks passed! PyInstaller is ready.
echo You can now run build.bat or simple_build.bat
echo ========================================
goto :end

:error
echo.
echo ========================================
echo Environment check failed!
echo Please fix the errors above before building.
echo ========================================

:end
pause
