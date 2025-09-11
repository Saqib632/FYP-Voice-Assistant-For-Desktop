@echo off
title Voice Assistant - Administrator Mode
color 0A

echo ========================================
echo    Voice Assistant - Administrator Mode
echo ========================================
echo.

REM Check if running as administrator
net session >nul 2>&1
if %errorLevel% == 0 (
    echo [SUCCESS] Running as Administrator
    echo.
) else (
    echo [ERROR] Not running as Administrator
    echo [INFO] Requesting Administrator privileges...
    echo.
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

echo [INFO] Navigating to project directory...
cd /d "D:\My Projects\New folder (2)"

echo [INFO] Starting Voice Assistant...
echo [INFO] You can now use WiFi control commands!
echo.
echo Available WiFi Commands:
echo   - "turn on wifi" or "enable wifi"
echo   - "turn off wifi" or "disable wifi" 
echo   - "wifi status" or "check wifi"
echo.
echo Press Ctrl+C to stop the assistant
echo ========================================
echo.

python voice.py

echo.
echo [INFO] Voice Assistant has stopped.
pause







