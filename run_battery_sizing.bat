@echo off
setlocal enabledelayedexpansion

echo === Battery Analysis: Starting Sizing Sweep ===
echo.

where conda >nul 2>nul
if errorlevel 1 (
    REM Same auto-detection as in setup.bat -- "conda" is not known by
    REM default when double-clicked, even if Anaconda/Miniconda is installed.
    echo "conda" is not known in this window -- searching the usual install locations ...
    set "CONDA_FOUND="
    for %%P in (
        "%USERPROFILE%\anaconda3"
        "%USERPROFILE%\miniconda3"
        "%LOCALAPPDATA%\anaconda3"
        "%LOCALAPPDATA%\miniconda3"
        "%LOCALAPPDATA%\Continuum\anaconda3"
        "%LOCALAPPDATA%\Continuum\miniconda3"
        "C:\ProgramData\Anaconda3"
        "C:\ProgramData\Miniconda3"
        "C:\Anaconda3"
        "C:\Miniconda3"
    ) do (
        if not defined CONDA_FOUND (
            if exist "%%~P\Scripts\conda.exe" (
                set "CONDA_FOUND=%%~P"
            )
        )
    )

    if defined CONDA_FOUND (
        echo Found at: !CONDA_FOUND! -- activating Conda ...
        call "!CONDA_FOUND!\Scripts\activate.bat" "!CONDA_FOUND!"
        where conda >nul 2>nul
    )

    if errorlevel 1 (
        echo.
        echo ERROR: "conda" was not found. Please run setup.bat first,
        echo or start this script from the "Anaconda Prompt" ^(search the Start menu^).
        pause
        exit /b 1
    )
)

echo Activating environment "batterieanalyse" ...
call conda activate batterieanalyse
if errorlevel 1 (
    echo.
    echo ERROR: Environment "batterieanalyse" not found. Please run setup.bat first.
    pause
    exit /b 1
)

REM This .bat lives in the Battery Analysis main folder (same level as
REM setup.bat), while battery_sizing.py lives in the "Code" subfolder --
REM so switch into that subfolder, regardless of where this script was
REM launched from (e.g. double-click vs. file drag & drop).
if not exist "%~dp0Code\battery_sizing.py" (
    echo.
    echo ERROR: "Code\battery_sizing.py" was not found under:
    echo   %~dp0Code
    echo Is run_battery_sizing.bat located in the Battery Analysis main
    echo folder, right next to the "Code" folder? If the Code folder has a
    echo different name or location, please let me know.
    pause
    exit /b 1
)
cd /d "%~dp0Code"

REM Same console-based file input as run_battery_analysis.bat (a tkinter
REM pop-up is unreliable from a .bat-launched console) -- drag your
REM Inputs.xlsx into this window and press Enter, or type the path.
set "INPUT_FILE=%~1"
if "%INPUT_FILE%"=="" (
    echo.
    echo No file was passed via drag ^& drop onto this .bat icon.
    set /p "INPUT_FILE=Drag your Inputs.xlsx into this window and press Enter (or type the path): "
)
REM Windows wraps a drag&dropped path in quotes automatically -- remove
REM them here, since we add our own quotes below.
set "INPUT_FILE=%INPUT_FILE:"=%"

if "%INPUT_FILE%"=="" (
    echo.
    echo ERROR: No file path given. Aborting.
    pause
    exit /b 1
)

echo.
echo Starting battery_sizing.py with:
echo   %INPUT_FILE%
echo.
echo NOTE: this runs a full optimization for several battery size combinations
echo (capacity x power), so it takes noticeably longer than a normal single run.
echo.

python battery_sizing.py "%INPUT_FILE%"

echo.
if errorlevel 1 (
    echo === ERROR during execution. See messages above. ===
) else (
    echo === Done! Sweep Excel and Heatmap PDF are in the Output/Inputs folder. ===
)
pause