@echo off
setlocal

echo === Batterieanalyse: Environment-Setup ===
echo.

where conda >nul 2>nul
if errorlevel 1 (
    echo FEHLER: "conda" wurde nicht gefunden.
    echo Bitte zuerst Miniconda installieren: https://docs.conda.io/en/latest/miniconda.html
    echo Danach dieses Skript in einer neuen "Anaconda Prompt" erneut ausfuehren.
    pause
    exit /b 1
)

echo Conda gefunden. Erstelle/aktualisiere Environment "batterieanalyse" aus environment.yml ...
call conda env create -f environment.yml
if errorlevel 1 (
    echo Environment existiert evtl. schon - versuche Update stattdessen ...
    call conda env update -f environment.yml --prune
)

if errorlevel 1 (
    echo.
    echo FEHLER beim Erstellen/Aktualisieren des Environments. Siehe Meldungen oben.
    pause
    exit /b 1
)

echo.
echo === Fertig! ===
echo Das Environment "batterieanalyse" ist bereit.
echo Du kannst jetzt run_battery_analysis.bat ausfuehren.
pause
