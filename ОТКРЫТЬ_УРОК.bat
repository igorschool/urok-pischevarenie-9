@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo  Урок откроется в браузере. Это окно не закрывайте.
echo  Интернет не нужен.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0start_lesson.ps1"
if errorlevel 1 pause
