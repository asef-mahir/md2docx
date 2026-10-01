@echo off
title Study Guide Watcher
cd /d "%~dp0"
".venv\Scripts\python.exe" "watcher.py"
echo.
echo Watcher stopped. Press any key to close.
pause >nul