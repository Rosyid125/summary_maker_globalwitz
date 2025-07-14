@echo off
echo ========================================
echo Excel Summary Maker - GlobalWitz X Volza
echo ========================================
echo.
echo Starting application...
echo.
echo FILE LOCATIONS:
echo - Input files: Place Excel files in the 'original_excel' folder
echo - Output files: Will be saved to Documents\ExcelSummaryMaker_Output\
echo.

REM Change to the directory where the script is located
cd /d "%~dp0"

REM Check if required folders exist
if not exist "original_excel" (
    echo Creating original_excel folder...
    mkdir "original_excel"
)

REM Start the application
start "" "ExcelSummaryMaker.exe"

echo Application started!
echo.
echo If you encounter any issues:
echo 1. Make sure your Excel files are in the 'original_excel' folder
echo 2. Check that you have write permissions to Documents folder
echo 3. Review the application logs for detailed error information
echo.
timeout /t 3 >nul
