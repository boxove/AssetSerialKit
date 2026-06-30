@echo off
rem Script author: NaF
setlocal EnableExtensions

title Remote PC and Monitor Serial Collection

echo Launching remote_collect...

set "SCRIPT_DIR=%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%remote_collect.ps1"
set "RESULT=%ERRORLEVEL%"

echo.
echo Script author: NaF
set /p "DUMMY=Press Enter to exit..."
exit /b %RESULT%
