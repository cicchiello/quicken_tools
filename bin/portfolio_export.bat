@echo off
setlocal

rem === CONFIGURE THESE PATHS ===
set "SRC=G:\My Drive\personal\finances\HOME.QDF"
set "DST=C:\tmp\HOME_nightly.QDF"

rem == AuthHotkey paths == 
set "AHKEXE=C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe"
set "AHK=\\10.0.0.214\pi-nas\openclaw\quicken_tools\QuickenPortfolioExport.ahk"

rem === EXPORTED CSV PRODUCED BY THE AHK SCRIPT ===
set "EXPORT_CSV=\\10.0.0.214\pi-nas\openclaw\quicken_tools\portfolio_nightly.csv"

rem === FINAL RENAMED COPY ===
set "FINAL_DIR=\\10.0.0.214\pi-nas\openclaw\quicken_tools\archive"
set "FINAL_CSV=portfolio_%DATE:~10,4%-%DATE:~4,2%-%DATE:~7,2%.csv"


rem === GUARD: source must exist ===
if not exist "%SRC%" (
    echo ERROR: Source file not found:
    echo   %SRC%
    exit /b 2
)

rem === ENSURE DESTINATION FOLDER EXISTS ===
for %%I in ("%DST%") do set "DSTDIR=%%~dpI"
if not exist "%DSTDIR%" (
    mkdir "%DSTDIR%"
    if errorlevel 1 (
        echo ERROR: Could not create destination folder:
        echo   %DSTDIR%
        exit /b 3
    )
)

rem === COPY ===
copy /Y "%SRC%" "%DST%" >nul
if errorlevel 1 (
    echo ERROR: Copy failed.
    exit /b 4
)

echo Copy successful:
echo   %SRC%
echo   ^>
echo   %DST%


"%AHKEXE%" "%AHK%"
set "RC=%ERRORLEVEL%"


if not "%RC%"=="0" (
    echo ERROR: AutoHotkey export script failed with code %RC%.
    exit /b %RC%
)

if not exist "%EXPORT_CSV%" (
    echo ERROR: Expected export CSV not found:
    echo   %EXPORT_CSV%
    exit /b 6
)

if not exist "%FINAL_DIR%" (
    mkdir "%FINAL_DIR%"
    if errorlevel 1 (
        echo ERROR: Could not create archive directory:
        echo   %FINAL_DIR%
        exit /b 7
    )
)

copy /Y "%EXPORT_CSV%" "%FINAL_DIR%\%FINAL_CSV%"
if errorlevel 1 (
    echo ERROR: Failed copying final CSV to archive name.
    exit /b 8
)

echo Success:
echo   Export created: %EXPORT_CSV%
echo   Archived as:    %FINAL_DIR%\%FINAL_CSV%

exit /b 0


