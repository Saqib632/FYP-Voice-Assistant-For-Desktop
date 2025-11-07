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
echo [INFO] Navigating to project directory...
REM Change to the directory where this batch file is located (works with spaces)
pushd "%~dp0"  >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Failed to change directory to "%~dp0". The script may not run from the correct folder.
)

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

REM Run the voice script using the py launcher if available, otherwise fallback to python.
REM Using the full path ensures the script is loaded from the project folder even if the current working directory differs.
if exist "%~dp0voice.py" (
    py "%~dp0voice.py" || python "%~dp0voice.py"
) else (
    echo [WARN] "%~dp0voice.py" not found. Trying to run by relative name...
    py voice.py || python voice.py
)

REM Return to original directory
popd >nul 2>&1

echo.
echo [INFO] Voice Assistant has stopped.
pause







