@echo off
setlocal
cd /d %~dp0
py -m PyInstaller --noconfirm --clean ^
    --name "Excel整理工具" ^
    --windowed ^
    --onefile ^
    --paths . ^
    --exclude-module pandas ^
    --exclude-module numpy ^
    --exclude-module PySide6 ^
    --exclude-module PyQt5 ^
    --exclude-module PyQt6 ^
    --exclude-module matplotlib ^
    --exclude-module scipy ^
    --exclude-module IPython ^
    --exclude-module notebook ^
    --exclude-module pytest ^
    --exclude-module tkinter.test ^
    run.py
echo.
echo === done. exe at dist\Excel整理工具.exe ===
endlocal
