@echo off
setlocal

rem === CONFIGURE THESE PATHS ===
set "SRC=G:\My Drive\personal\finances\HOME.QDF"
set "DST=C:\tmp\HOME_nightly.QDF"

rem === AutoHotkey paths ===
set "AHKEXE=C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe"
set "AHK=\\10.0.0.214\pi-nas\openclaw\quicken_tools\QuickenPortfolioExport.ahk"

rem === EXPORTED CSV PRODUCED BY THE AHK SCRIPT ===
set "EXPORT_CSV=\\10.0.0.214\pi-nas\openclaw\quicken_tools\portfolio_nightly.csv"

rem === FINAL RENAMED COPY ===
set "FINAL_DIR=\\10.0.0.214\pi-nas\openclaw\quicken_tools\archive"

for /f %%I in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd"') do set "TODAY=%%I"
set "FINAL_CSV=portfolio_%TODAY%.csv"

rem === GUARD: source must exist ===
if not exist "%SRC%" (
    echo ERROR: Source file not found:
    echo   %SRC%
    exit /b 2
)

rem === GUARD: AutoHotkey executable must exist ===
if not exist "%AHKEXE%" (
    echo ERROR: AutoHotkey executable not found:
    echo   %AHKEXE%
    exit /b 3
)

rem === GUARD: AHK script must exist ===
if not exist "%AHK%" (
    echo ERROR: AutoHotkey script not found:
    echo   %AHK%
    exit /b 4
)

rem === ENSURE DESTINATION FOLDER EXISTS ===
for %%I in ("%DST%") do set "DSTDIR=%%~dpI"
if not exist "%DSTDIR%" (
    mkdir "%DSTDIR%"
    if errorlevel 1 (
        echo ERROR: Could not create destination folder:
        echo   %DSTDIR%
        exit /b 5
    )
)

rem === COPY QDF TO LOCAL WORKING FILE ===
copy /Y "%SRC%" "%DST%" >nul
if errorlevel 1 (
    echo ERROR: Copy of QDF failed.
    exit /b 6
)

echo Copy successful:
echo   %SRC%
echo   ^>
echo   %DST%

rem === RUN AHK EXPORT SCRIPT AND WAIT FOR IT ===
"%AHKEXE%" "%AHK%"
set "RC=%ERRORLEVEL%"

if not "%RC%"=="0" (
    echo ERROR: AutoHotkey export script failed with code %RC%.
    exit /b %RC%
)

rem === CHECK FOR EXPECTED EXPORT OUTPUT ===
if not exist "%EXPORT_CSV%" (
    echo ERROR: Expected export CSV not found:
    echo   %EXPORT_CSV%
    exit /b 7
)

rem === ENSURE ARCHIVE DIRECTORY EXISTS ===
if not exist "%FINAL_DIR%" (
    mkdir "%FINAL_DIR%"
    if errorlevel 1 (
        echo ERROR: Could not create archive directory:
        echo   %FINAL_DIR%
        exit /b 8
    )
)

rem === COPY/RENAME EXPORT TO ARCHIVE FILE ===
copy /Y "%EXPORT_CSV%" "%FINAL_DIR%\%FINAL_CSV%" >nul
if errorlevel 1 (
    echo ERROR: Failed copying final CSV to archive name.
    exit /b 9
)

echo Success:
echo   Export created: %EXPORT_CSV%
echo   Archived as:    %FINAL_DIR%\%FINAL_CSV%

exit /b 0
