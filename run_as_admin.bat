@echo off
setlocal enabledelayedexpansion

REM Check if running as administrator
net session >nul 2>&1
if errorlevel 1 (
    REM Re-run this batch file with admin privileges
    powershell -NoProfile -Command "Start-Process cmd.exe -ArgumentList '/c %~f0' -Verb RunAs -WindowStyle Hidden"
    exit /b
)

REM Get parent directory where main.js and package.json are located
cd /d "%~dp0.." 2>nul

REM Start npm with completely hidden window
REM This ensures Electron and all child processes run as admin
start /B npm start

REM Exit immediately - npm runs in background with all admin privileges inherited
exit /b







