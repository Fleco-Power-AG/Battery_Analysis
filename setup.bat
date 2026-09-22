@echo off
setlocal enabledelayedexpansion

echo === Battery Analysis: Environment Setup ===
echo.

where conda >nul 2>nul
if errorlevel 1 (
    REM Anaconda/Miniconda does NOT add itself to the system PATH by
    REM default -- a plain cmd window opened by double-clicking this file
    REM therefore doesn't know "conda" even if it is installed. Instead of
    REM failing immediately, search the usual install locations first and
    REM activate Conda from there.
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
        echo ERROR: "conda" was not found -- not even in the usual install locations.
        echo Please install Miniconda first: https://docs.conda.io/en/latest/miniconda.html
        echo Then either run this script from the "Anaconda Prompt" ^(search the Start
        echo menu^), or reinstall Anaconda/Miniconda and check "Add to PATH".
        pause
        exit /b 1
    )
)

echo Conda found. Creating/updating environment "batterieanalyse" from environment.yml ...
call conda env create -f environment.yml
if errorlevel 1 (
    echo Environment may already exist - trying update instead ...
    call conda env update -f environment.yml --prune
)

if errorlevel 1 (
    echo.
    echo ERROR while creating/updating the environment. See messages above.
    pause
    exit /b 1
)

echo.
echo === Done! ===
echo The "batterieanalyse" environment is ready.
echo You can now run run_battery_analysis.bat.
pause