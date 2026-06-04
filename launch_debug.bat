@echo off
cd /d "%~dp0"
echo Running app in debug mode...
py app.py
if errorlevel 1 (
  echo.
  echo If an error appeared, send me the full text from this window.
)
pause
