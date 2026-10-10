@echo off
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "d4b7d33a-dc63-4cdc-8fca-5107d2df5af5" /f >nul 2>nul
del "%LOCALAPPDATA%\ChartLite\manifest.xml" >nul 2>nul
echo ChartLite removed. Restart PowerPoint.
pause
