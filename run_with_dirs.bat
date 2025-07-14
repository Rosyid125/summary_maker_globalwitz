@echo off
echo ========================================
echo Excel Summary Maker - GlobalWitz X Volza
echo ========================================
echo.
echo Checking and creating required directories...

if not exist "original_excel" (
    echo Creating original_excel directory...
    mkdir "original_excel"
    echo - Created: original_excel\
) else (
    echo - Found: original_excel\
)

if not exist "processed_excel" (
    echo Creating processed_excel directory...
    mkdir "processed_excel"
    echo - Created: processed_excel\
) else (
    echo - Found: processed_excel\
)

echo.
echo Directory setup complete!
echo.
echo Starting the application...
echo ========================================
python main.py

echo.
echo ========================================
echo Application finished.
echo ========================================
pause
