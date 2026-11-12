@echo off
:: build_win.bat — builds StripChartDashboard.exe on Windows
::
:: Run from the Excel_dash_UI\ folder:
::   build_win.bat   (double-click or run from Command Prompt)
::
:: Output: dist\StripChartDashboard\StripChartDashboard.exe
:: -------------------------------------------------------

echo ===================================================
echo   StripChart Dashboard — Windows build
echo ===================================================

:: ── 1. Locate StripChart_macros.xlsm ─────────────────────
set XLSM=..\Excel_dashboard\StripChart_macros.xlsm
if not exist "%XLSM%" (
    echo.
    echo ERROR: StripChart_macros.xlsm not found.
    echo Expected location: %XLSM%
    echo Make sure the Excel_dashboard folder is next to Excel_dash_UI\.
    pause
    exit /b 1
)
echo Found: %XLSM%

:: ── 2. Install / upgrade build dependencies ───────────────
echo.
echo Installing build dependencies...
pip install pyinstaller --quiet
pip install -r requirements.txt --quiet

:: ── 3. Clean previous builds ──────────────────────────────
echo Cleaning previous builds...
if exist build rmdir /s /q build
if exist dist  rmdir /s /q dist

:: ── 4. Build ──────────────────────────────────────────────
echo.
echo Building... (this takes 2-5 minutes)
pyinstaller stripchart.spec

:: ── 5. Done ───────────────────────────────────────────────
echo.
echo ===================================================
echo   Build complete!
echo.
echo   Folder: dist\StripChartDashboard\
echo   Exe:    dist\StripChartDashboard\StripChartDashboard.exe
echo.
echo   To distribute to Windows users:
echo     Zip the dist\StripChartDashboard folder.
echo     Share StripChartDashboard_win.zip
echo.
echo   User instructions:
echo     1. Unzip ^& run StripChartDashboard.exe
echo     2. Open StripChart_macros.xlsm in Excel
echo        (copied to %%USERPROFILE%%\Documents\ExcelDashboard\)
echo     3. Open your data workbook in Excel
echo     4. Use the StripChart tab in the browser dashboard
echo ===================================================
pause
