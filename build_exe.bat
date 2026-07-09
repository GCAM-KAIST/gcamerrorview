@echo off
REM ============================================================
REM  Build gcamlog into a single double-click .exe
REM  Needs Python 3.9+ on PATH. Run once, then share the .exe.
REM ============================================================

python -m pip install --upgrade pyinstaller openpyxl
if errorlevel 1 goto :err

python -m PyInstaller --onefile --windowed --name gcamlog --icon assets\gcamlog.ico --add-data "assets\gcamlog.ico;assets" --add-data "assets\gcamlog.png;assets" --add-data "data\gcam_regions.csv;data" --add-data "data\gcam_systems.csv;data" --distpath . gcamlog.py
if errorlevel 1 goto :err

echo.
echo Done. gcamlog.exe is in this folder.
goto :end

:err
echo.
echo Build failed. Make sure Python 3.9+ is installed and on your PATH.

:end
pause
