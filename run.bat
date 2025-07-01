@echo off
title Excel Summary Maker - GlobalWitz X Volza
echo Starting Excel Summary Maker...
echo.

REM Create necessary directories
if not exist "original_excel" mkdir "original_excel"
if not exist "processed_excel" mkdir "processed_excel"
if not exist "logs" mkdir "logs"

REM Run the application
echo Starting Excel Summary Maker GUI...
python main.py

pause
