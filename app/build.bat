@echo off
cd /d "%~dp0.."
pip install pyinstaller
pyinstaller --onefile --windowed ^
    --name "Conciliacao Mercado" ^
    --add-data "AB;AB" ^
    app/main.py
echo.
echo Build concluido: dist\Conciliacao Mercado.exe
pause
