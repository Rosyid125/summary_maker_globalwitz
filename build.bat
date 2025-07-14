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

REM Run PyInstaller (using --onedir for faster startup and smaller main exe)
echo.
echo Starting PyInstaller build...
python -m PyInstaller --onedir --windowed --name=ExcelSummaryMaker --add-data="src;src" --add-data="original_excel;original_excel" --add-data="processed_excel;processed_excel" --add-data="logs;logs" --hidden-import=tkinter --hidden-import=tkinter.ttk --hidden-import=pandas --hidden-import=openpyxl --hidden-import=numpy --hidden-import=xlsxwriter --hidden-import=dateutil --collect-all=tkinter --collect-all=openpyxl --collect-all=pandas --clean main.py

if errorlevel 1 (
    echo.
    echo ERROR: PyInstaller build failed!
    pause
    exit /b 1
)

REM Check if executable was created
if not exist "dist\ExcelSummaryMaker\ExcelSummaryMaker.exe" (
    echo ERROR: Executable not found!
    pause
    exit /b 1
)

REM Create necessary directories in dist
echo.
echo Setting up directories...
if not exist "dist\ExcelSummaryMaker\original_excel" mkdir "dist\ExcelSummaryMaker\original_excel"
if not exist "dist\ExcelSummaryMaker\processed_excel" mkdir "dist\ExcelSummaryMaker\processed_excel"
if not exist "dist\ExcelSummaryMaker\logs" mkdir "dist\ExcelSummaryMaker\logs"

REM Copy additional files
echo Copying additional files...
if exist "README.md" copy "README.md" "dist\ExcelSummaryMaker\" >nul
if exist "requirements.txt" copy "requirements.txt" "dist\ExcelSummaryMaker\" >nul
if exist "DISTRIBUTION_README.md" copy "DISTRIBUTION_README.md" "dist\ExcelSummaryMaker\README.md" >nul

REM Create launcher files
echo Creating launcher files...
if exist "launcher_template.bat" (
    copy "launcher_template.bat" "dist\ExcelSummaryMaker\Launch_ExcelSummaryMaker.bat" >nul
) else (
    echo @echo off > "dist\ExcelSummaryMaker\Launch_ExcelSummaryMaker.bat"
    echo cd /d "%%~dp0" >> "dist\ExcelSummaryMaker\Launch_ExcelSummaryMaker.bat"
    echo if not exist "original_excel" mkdir "original_excel" >> "dist\ExcelSummaryMaker\Launch_ExcelSummaryMaker.bat"
    echo start "" "ExcelSummaryMaker.exe" >> "dist\ExcelSummaryMaker\Launch_ExcelSummaryMaker.bat"
)

REM Create shortcut creation script
echo Set oWS = WScript.CreateObject("WScript.Shell") > "dist\ExcelSummaryMaker\CreateShortcut.vbs"
echo sLinkFile = oWS.ExpandEnvironmentStrings("%%%%USERPROFILE%%%%\Desktop\Excel Summary Maker.lnk") >> "dist\ExcelSummaryMaker\CreateShortcut.vbs"
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> "dist\ExcelSummaryMaker\CreateShortcut.vbs"
echo oLink.TargetPath = oWS.CurrentDirectory ^& "\ExcelSummaryMaker.exe" >> "dist\ExcelSummaryMaker\CreateShortcut.vbs"
echo oLink.WorkingDirectory = oWS.CurrentDirectory >> "dist\ExcelSummaryMaker\CreateShortcut.vbs"
echo oLink.Description = "Excel Summary Maker - GlobalWitz X Volza" >> "dist\ExcelSummaryMaker\CreateShortcut.vbs"
echo oLink.Save >> "dist\ExcelSummaryMaker\CreateShortcut.vbs"

REM Get file size
for %%F in ("dist\ExcelSummaryMaker\ExcelSummaryMaker.exe") do set size=%%~zF
set /a sizeMB=!size!/1024/1024

REM Get total folder size
echo Calculating total distribution size...
for /f "tokens=3" %%a in ('dir "dist\ExcelSummaryMaker" /s /-c ^| findstr /C:" bytes"') do set totalSize=%%a
set /a totalSizeMB=!totalSize!/1024/1024

echo.
echo ============================================================
echo BUILD COMPLETED SUCCESSFULLY!
echo ============================================================
echo.
echo Main executable: dist\ExcelSummaryMaker\ExcelSummaryMaker.exe
echo Main exe size: %sizeMB% MB
echo Total distribution size: %totalSizeMB% MB
echo.
echo IMPORTANT CHANGES FOR OUTPUT FILES:
echo - Input files: Place in 'original_excel' folder next to exe
echo - Output files: Saved to user's Documents\ExcelSummaryMaker_Output\
echo - This prevents permission issues in Program Files
echo.
echo BENEFITS OF THIS BUILD:
echo - Faster startup time (no extraction needed)
echo - Smaller main executable file
echo - Libraries organized in separate folder
echo - Easy to distribute entire folder
echo - Better file path handling for different Windows versions
echo.
echo You can distribute the entire 'dist\ExcelSummaryMaker' folder to users.
echo Opening dist folder...
echo.
explorer "dist\ExcelSummaryMaker"
pause
