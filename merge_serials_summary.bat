@echo off
rem Script author: NaF
setlocal EnableExtensions

title Merge PC and Monitor Serial CSV Files

echo Launching merge_serials_summary...

set "SCRIPT_DIR=%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%merge_serials_summary.ps1"
set "RESULT=%ERRORLEVEL%"

echo.
echo Script author: NaF
set /p "DUMMY=Press Enter to exit..."
exit /b %RESULT%
