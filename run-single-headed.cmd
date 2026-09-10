@echo off
setlocal
title YanJia Single Case Runner

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\run-single-headed.ps1" %*
set "EXIT_CODE=%ERRORLEVEL%"

echo.
echo Run finished. Exit code: %EXIT_CODE%
if /I not "%YANJIA_LAUNCHER_NO_PAUSE%"=="1" pause
exit /b %EXIT_CODE%
