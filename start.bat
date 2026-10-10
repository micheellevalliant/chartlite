@echo off
title ChartLite for PowerPoint
cd /d "%~dp0"
where node >nul 2>nul
if errorlevel 1 (
  echo.
  echo  Node.js is not installed.
  echo  Download the LTS version from https://nodejs.org , install it, then double-click start.bat again.
  echo.
  pause
  exit /b 1
)
if not exist node_modules (
  echo First run: installing ChartLite tools, this takes a minute...
  call npm install --no-audit --no-fund || goto :fail
)
if not exist "%USERPROFILE%\.office-addin-dev-certs\localhost.crt" (
  echo Installing the local HTTPS certificate - click YES if Windows asks.
  call npm run certs || goto :fail
)
echo Starting ChartLite and opening PowerPoint...
call npm start || goto :fail
echo.
echo ChartLite is running. Keep the "dev server" window open while you use it.
echo In PowerPoint: Home tab ^> ChartLite.
pause
exit /b 0
:fail
echo.
echo Something went wrong - see the messages above.
pause
exit /b 1
