@echo off
title NCC OD Form Automator
color 0A
cls

echo.
echo  ============================================================
echo   NCC OD FORM AUTOMATOR  -  Sri Ramakrishna College
echo  ============================================================
echo.
echo  Starting server, please wait...
echo.

cd /d "%~dp0"

:: Kill any existing server on port 3000
for /f "tokens=5" %%a in ('netstat -aon ^| findstr :3000 ^| findstr LISTENING') do (
    taskkill /PID %%a /F >nul 2>&1
)

:: Start server in background
start /B node server.js > server.log 2>&1

:: Wait for server to start (poll up to 10 seconds)
set /a tries=0
:WAIT_LOOP
timeout /t 1 /nobreak >nul
set /a tries+=1
powershell -Command "try { Invoke-WebRequest http://localhost:3000 -UseBasicParsing -TimeoutSec 1 | Out-Null; exit 0 } catch { exit 1 }" >nul 2>&1
if %errorlevel%==0 goto SERVER_READY
if %tries% GEQ 10 goto SERVER_TIMEOUT
goto WAIT_LOOP

:SERVER_READY
echo.
echo  ============================================================
echo   SERVER IS RUNNING!
echo   Open your browser at:  http://localhost:3000
echo  ============================================================
echo.
echo  The app is now open in your browser.
echo  Keep this window open while using the app.
echo  Press Ctrl+C to stop the server when done.
echo.
start "" "http://localhost:3000"
goto WAIT_END

:SERVER_TIMEOUT
echo.
echo  WARNING: Server took too long to start.
echo  Try opening http://localhost:3000 manually.
echo.

:WAIT_END
pause
