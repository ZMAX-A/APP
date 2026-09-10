@echo off
setlocal
title YanJia Full Suite Runner

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\run-full-visible.ps1" %*
set "EXIT_CODE=%ERRORLEVEL%"

echo.
echo Full run finished. Exit code: %EXIT_CODE%
if /I not "%YANJIA_LAUNCHER_NO_PAUSE%"=="1" pause
exit /b %EXIT_CODE%
