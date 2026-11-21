@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title Convi - сборка EXE
echo.
echo  CONVI - BUILD
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
".venv\Scripts\python.exe" -m pip install --upgrade pip -q
".venv\Scripts\python.exe" -m pip install -r requirements.txt pyinstaller
if errorlevel 1 goto :error
if exist "build" rmdir /s /q "build"
if exist "dist\Convi.exe" del /q "dist\Convi.exe"
set UPXARG=
if exist "upx\upx.exe" set UPXARG=--upx-dir upx
echo Собираю один EXE ^(лишние части Qt отброшены, несколько минут^)...
".venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean %UPXARG% convi.spec
if errorlevel 1 goto :error
echo.
echo Готово: %CD%\dist\Convi.exe
explorer dist
pause
exit /b 0

:error
echo.
echo Сборка не удалась. Пришли текст ошибки выше.
pause
exit /b 1
