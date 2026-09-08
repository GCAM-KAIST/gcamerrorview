@echo off
REM ============================================================
REM  Build gcamerrorview into a single double-click .exe
REM  Needs Python 3.9+ on PATH. Run once, then share the .exe.
REM ============================================================

python -m pip install --upgrade pyinstaller openpyxl pillow
if errorlevel 1 goto :err

python -m PyInstaller --onefile --windowed --name gcamerrorview --icon assets\gcamerrorview.ico --add-data "assets\gcamerrorview.ico;assets" --add-data "assets\gcamerrorview.png;assets" --add-data "data\gcam_regions.csv;data" --add-data "data\gcam_systems.csv;data" --distpath . gcamerrorview.py
if errorlevel 1 goto :err

echo.
echo Done. gcamerrorview.exe is in this folder.
goto :end

:err
echo.
echo Build failed. Make sure Python 3.9+ is installed and on your PATH.

:end
pause
