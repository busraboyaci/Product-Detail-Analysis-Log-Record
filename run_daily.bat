@echo off
REM ==========================================================================
REM  run_daily.bat  -  one-click daily runner for the log analysis tool
REM
REM  What it does:
REM    1. Makes sure input\ output\ archive\ folders exist.
REM    2. Analyses EVERY Excel file currently in the input\ folder.
REM    3. Writes a timestamped report per file into output\.
REM    4. Moves each processed input file into archive\ so it is not
REM       analysed again tomorrow.
REM
REM  Daily use: drop today's log export files into input\, then double-click
REM  this file.
REM ==========================================================================

setlocal enabledelayedexpansion
chcp 65001 >nul
set "PYTHONUTF8=1"
cd /d "%~dp0"

echo ============================================================
echo   Product Detail Log Analysis - Daily Run
echo ============================================================
echo.

if not exist "input"   mkdir "input"
if not exist "output"  mkdir "output"
if not exist "archive" mkdir "archive"

REM Pick a Python launcher: prefer the "py" launcher, fall back to "python".
where py >nul 2>&1 && (set "PY=py") || (set "PY=python")

set "found=0"
for %%F in ("input\*.xlsx" "input\*.xlsm" "input\*.xls") do (
    set "name=%%~nxF"
    REM Skip Excel lock files (they start with ~$)
    if not "!name:~0,2!"=="~$" (
        set /a found+=1
        echo [RUN] Analysing: %%~nxF
        %PY% "product_detail_log_analysis.py" --input "%%~nxF"
        if !errorlevel! equ 0 (
            move /y "%%F" "archive\" >nul
            echo [DONE] Moved to archive\: %%~nxF
        ) else (
            echo [FAIL] Analysis failed - file left in input\: %%~nxF
        )
        echo.
    )
)

if "!found!"=="0" (
    echo [INFO] No Excel files found in the input\ folder.
    echo        Put today's log export files into the input\ folder,
    echo        then run this again.
    echo.
)

echo ============================================================
echo   Finished. Your reports are in the output\ folder.
echo ============================================================
echo.
pause
endlocal
