@echo off
rem Script author: NaF
setlocal EnableExtensions

title PC and Monitor Serial Export

echo Launching show_serials...

set "SCRIPT_DIR=%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%collect_serials.ps1"
set "RESULT=%ERRORLEVEL%"

echo.
echo Script author: NaF
set /p "DUMMY=Press Enter to exit..."
exit /b %RESULT%
