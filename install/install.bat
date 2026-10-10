@echo off
title Install ChartLite
set "DEST=%LOCALAPPDATA%\ChartLite"
if not exist "%DEST%" mkdir "%DEST%"
copy /Y "%~dp0manifest.xml" "%DEST%\manifest.xml" >nul
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "d4b7d33a-dc63-4cdc-8fca-5107d2df5af5" /t REG_SZ /d "%DEST%\manifest.xml" /f >nul
if errorlevel 1 (
  echo Could not register ChartLite. Your PC policy may block add-in sideloading.
  pause & exit /b 1
)
echo.
echo  ChartLite installed.
echo  Close PowerPoint completely, open it again, then: Home tab ^> ChartLite.
echo  (If the button is missing: Home ^> Add-ins ^> More Add-ins ^> Developer Add-ins ^> ChartLite.)
echo.
pause
