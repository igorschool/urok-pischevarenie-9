@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo  Урок 10 — сервер для доски и камеры
echo  Открой: http://127.0.0.1:8770/
echo.
echo  Не закрывай это окно до конца урока.
echo.
start "" "http://127.0.0.1:8770/"
py -3 -m http.server 8770
if errorlevel 1 (
  echo Python не найден. Открой index.html вручную.
  echo Камера через file:// часто не работает.
  pause
)
