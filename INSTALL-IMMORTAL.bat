@echo off
echo Installing WinRM Bridge as IMMORTAL service...
echo.

cd /d "%~dp0"

REM Install dependencies
echo Installing npm packages...
call npm install

REM Create startup script
echo node "%~dp0server.js" > "%~dp0start-winrm-bridge.bat"

REM Add to startup folder
set STARTUP=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup
copy "%~dp0start-winrm-bridge.bat" "%STARTUP%\HiveWinRMBridge.bat"

echo.
echo ╔════════════════════════════════════════╗
echo ║  WinRM Bridge - IMMORTAL Installed   ║
echo ╚════════════════════════════════════════╝
echo ✓ Service installed
echo ✓ Will auto-start on boot
echo ✓ Running on port 8775
echo ✓ Hive mesh integration enabled
echo.
echo Starting service now...
start "" "%~dp0start-winrm-bridge.bat"

echo.
echo Service is running!
echo Access at: http://localhost:8775
pause
