@echo off
cd /d "%~dp0"
where node >nul 2>&1
if errorlevel 1 (
  echo Node.js が必要です。https://nodejs.org からインストールしてください。
  pause
  exit /b 1
)
start "" "http://127.0.0.1:8765/"
node server.js
pause
