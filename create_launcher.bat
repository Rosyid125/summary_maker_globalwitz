@echo off
echo Creating launcher script...

REM Create a simple launcher script in the dist folder
echo @echo off > "dist\ExcelSummaryMaker\Launch_ExcelSummaryMaker.bat"
echo cd /d "%%~dp0" >> "dist\ExcelSummaryMaker\Launch_ExcelSummaryMaker.bat"
echo start "" "ExcelSummaryMaker.exe" >> "dist\ExcelSummaryMaker\Launch_ExcelSummaryMaker.bat"

REM Create a shortcut creation script
echo Creating shortcut helper...
echo Set oWS = WScript.CreateObject("WScript.Shell") > "dist\ExcelSummaryMaker\CreateShortcut.vbs"
echo sLinkFile = oWS.ExpandEnvironmentStrings("%%USERPROFILE%%\Desktop\Excel Summary Maker.lnk") >> "dist\ExcelSummaryMaker\CreateShortcut.vbs"
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> "dist\ExcelSummaryMaker\CreateShortcut.vbs"
echo oLink.TargetPath = oWS.CurrentDirectory ^& "\ExcelSummaryMaker.exe" >> "dist\ExcelSummaryMaker\CreateShortcut.vbs"
echo oLink.WorkingDirectory = oWS.CurrentDirectory >> "dist\ExcelSummaryMaker\CreateShortcut.vbs"
echo oLink.Description = "Excel Summary Maker - GlobalWitz X Volza" >> "dist\ExcelSummaryMaker\CreateShortcut.vbs"
echo oLink.Save >> "dist\ExcelSummaryMaker\CreateShortcut.vbs"
echo WScript.Echo "Desktop shortcut created successfully!" >> "dist\ExcelSummaryMaker\CreateShortcut.vbs"

echo.
echo Launcher files created successfully!
echo - Launch_ExcelSummaryMaker.bat: Simple launcher
echo - CreateShortcut.vbs: Creates desktop shortcut
echo.
