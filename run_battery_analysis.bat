@echo off
setlocal enabledelayedexpansion

echo === Battery Analysis: Starting Optimization ===
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
REM setup.bat), while run_battery_analysis.py lives in the "Code"
REM subfolder -- so switch into that subfolder, regardless of where this
REM script was launched from (e.g. double-click vs. file drag & drop).
if not exist "%~dp0Code\run_battery_analysis.py" (
    echo.
    echo ERROR: "Code\run_battery_analysis.py" was not found under:
    echo   %~dp0Code
    echo Is run_battery_analysis.bat located in the Battery Analysis main
    echo folder, right next to the "Code" folder? If the Code folder has a
    echo different name or location, please let me know.
    pause
    exit /b 1
)
cd /d "%~dp0Code"

REM NEU (Beat, 22.9.2026: "leider kam kein Pop-up"): der tkinter-Datei-
REM Dialog aus run_battery_analysis.py ist in einer per .bat gestarteten
REM Konsole unzuverlaessig (Fokus-/Sichtbarkeitsprobleme je nach Windows-
REM Session). Statt darauf zu warten, wird jetzt direkt in der Konsole nach
REM dem Pfad gefragt, falls keine Datei per Drag&Drop aufs .bat-Icon
REM uebergeben wurde -- die Inputs.xlsx kann dafuer einfach in dieses
REM Fenster hineingezogen werden (Windows fuegt den Pfad automatisch ein).
set "INPUT_FILE=%~1"
if "%INPUT_FILE%"=="" (
    echo.
    echo No file was passed via drag ^& drop onto this .bat icon.
    set /p "INPUT_FILE=Drag your Inputs.xlsx into this window and press Enter (or type the path): "
)
REM Windows umschliesst einen per Drag&Drop eingefuegten Pfad automatisch
REM mit Anfuehrungszeichen -- die entfernen wir hier, da wir selbst unten
REM quoten.
set "INPUT_FILE=%INPUT_FILE:"=%"

if "%INPUT_FILE%"=="" (
    echo.
    echo ERROR: No file path given. Aborting.
    pause
    exit /b 1
)

echo.
echo Starting run_battery_analysis.py with:
echo   %INPUT_FILE%
echo.

python run_battery_analysis.py "%INPUT_FILE%"

echo.
if errorlevel 1 (
    echo === ERROR during execution. See messages above. ===
) else (
    echo === Done! Result Excel and PDF report are in the Output/Inputs folder. ===
)
pause