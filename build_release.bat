@echo off
setlocal
set "PROJECT=%~dp0"
cd /d "%PROJECT%"

set "APP_NAME=TuanziConverter"
python -m PyInstaller --noconsole --onefile --clean --name "%APP_NAME%" --icon "assets\app.ico" --add-data "templates;templates" --add-data "bin;bin" --add-data "assets;assets" "app_qt.py"

if exist build rmdir /s /q build
if exist __pycache__ rmdir /s /q __pycache__
if exist "%APP_NAME%.spec" del /q "%APP_NAME%.spec"

echo Done.
pause
