@echo off
REM Double-click this to re-read the roster sheet, re-pull stats, and open the tool.
REM The roster .xlsx is only ever read -- never written to.
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0build_data.ps1" %*
if errorlevel 1 (
  echo.
  echo *** Build failed - the tool still has the previous data. ***
  echo.
  pause
  exit /b 1
)
start "" "%~dp0index.html"
echo.
pause
