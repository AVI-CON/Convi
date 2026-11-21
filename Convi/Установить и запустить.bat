@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title Convi - запуск
echo.
echo  CONVI
echo ---------------------------------------------
if not exist ".venv\Scripts\python.exe" (
    echo Создаю виртуальное окружение...
    py -3.13 -m venv .venv 2>nul
    if errorlevel 1 py -3.12 -m venv .venv 2>nul
    if errorlevel 1 py -3.11 -m venv .venv 2>nul
    if errorlevel 1 python -m venv .venv 2>nul
    if errorlevel 1 (
        echo.
        echo Не найден Python 3.11-3.13 x64. Установи его с python.org и включи "Add Python to PATH".
        pause
        exit /b 1
    )
)
echo Проверяю библиотеки ^(при первом запуске нужен интернет, это займёт несколько минут^)...
".venv\Scripts\python.exe" -m pip install --upgrade pip -q
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto :error
echo Запускаю...
start "" ".venv\Scripts\pythonw.exe" convi.py
exit /b 0

:error
echo.
echo Не удалось установить библиотеки. Проверь интернет и запусти снова.
pause
exit /b 1
