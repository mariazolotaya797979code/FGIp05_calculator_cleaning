@echo off
chcp 65001 >nul
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install_shortcut.ps1"
if errorlevel 1 (
    echo.
    echo Не удалось создать ярлык.
    pause
    exit /b 1
)
echo.
echo Готово. На рабочем столе появилась иконка «Калькулятор клининга».
echo Файл calculator_data.xlsx должен лежать в этой же папке.
pause
