@echo off
setlocal EnableDelayedExpansion
echo ============================================================
echo Building Excel Summary Maker
echo ============================================================
echo.

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH!
    echo Please install Python and try again.
    pause
    exit /b 1
)

REM Check if PyInstaller is available
python -m pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo PyInstaller not found, installing...
    python -m pip install pyinstaller
    if errorlevel 1 (
        echo ERROR: Failed to install PyInstaller!
        pause
        exit /b 1
    )
)

REM Install/upgrade required packages
echo Installing/upgrading required packages...
python -m pip install --upgrade -r requirements.txt
if errorlevel 1 (
    echo ERROR: Failed to install required packages!
    pause
    exit /b 1
)

REM Clean previous builds
echo.
echo Cleaning previous builds...
if exist "build" (
    echo Removing build directory...
    rmdir /s /q "build"
)
if exist "dist" (
    echo Removing dist directory...
    rmdir /s /q "dist"
)
if exist "*.spec" (
    echo Removing spec files...
    del "*.spec"
)

REM Check if main.py exists
if not exist "main.py" (
    echo ERROR: main.py not found!
    pause
    exit /b 1
)

REM Run PyInstaller
echo.
echo Starting PyInstaller build...
pyinstaller --onefile --windowed --name=ExcelSummaryMaker --add-data="src;src" --add-data="original_excel;original_excel" --add-data="processed_excel;processed_excel" --add-data="logs;logs" --hidden-import=tkinter --hidden-import=tkinter.ttk --hidden-import=pandas --hidden-import=openpyxl --hidden-import=numpy --hidden-import=xlsxwriter --hidden-import=dateutil --collect-all=tkinter --collect-all=openpyxl --collect-all=pandas --clean main.py

if errorlevel 1 (
    echo.
    echo ERROR: PyInstaller build failed!
    pause
    exit /b 1
)

REM Check if executable was created
if not exist "dist\ExcelSummaryMaker.exe" (
    echo ERROR: Executable not found!
    pause
    exit /b 1
)

REM Create necessary directories in dist
echo.
echo Setting up directories...
if not exist "dist\original_excel" mkdir "dist\original_excel"
if not exist "dist\processed_excel" mkdir "dist\processed_excel"
if not exist "dist\logs" mkdir "dist\logs"

REM Copy additional files
echo Copying additional files...
if exist "README.md" copy "README.md" "dist\" >nul
if exist "requirements.txt" copy "requirements.txt" "dist\" >nul

REM Get file size
for %%F in ("dist\ExcelSummaryMaker.exe") do set size=%%~zF
set /a sizeMB=!size!/1024/1024

echo.
echo ============================================================
echo BUILD COMPLETED SUCCESSFULLY!
echo ============================================================
echo.
echo Executable: dist\ExcelSummaryMaker.exe
echo Size: %sizeMB% MB
echo.
echo You can distribute the entire 'dist' folder to users.
echo Opening dist folder...
echo.
explorer dist
pause
