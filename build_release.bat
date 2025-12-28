@echo off
setlocal
set PROJECT=%~dp0
cd /d %PROJECT%

python -m PyInstaller --noconsole --onefile --clean --name "Markdown转Word" --icon "assets\app.ico" --add-data "templates;templates" --add-data "bin;bin" --add-data "assets;assets" "app_qt.py"

if exist build rmdir /s /q build
if exist __pycache__ rmdir /s /q __pycache__
if exist Markdown转Word.spec del /q Markdown转Word.spec

echo Done.
pause
