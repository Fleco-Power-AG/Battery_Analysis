@echo off
setlocal
chcp 65001 >nul
title Batterieanalyse - Optimierung

rem ============================================================================
rem  Ablage : Hauptordner "17 Batterieanalyse"
rem  Start  : Inputs_<Name>.xlsx auf diese Datei ziehen ODER doppelklicken und
rem           die Datei ins Fenster ziehen.
rem  Ergebnisse landen im Unterordner "Output" neben der Inputs-Datei.
rem ============================================================================

set "ROOT=%~dp0"
set "CODE=%ROOT%05 Code"
set "SETUP=%ROOT%06 Setup"
rem Stammdaten (SwissIX, Solar_RMP, SRL-Marktpreise) zentral fuer alle Projekte
set "BATTERIE_STAMMDATEN_DIR=%ROOT%03 Stammdaten"
set "SKRIPT=run_battery_analysis.py"

echo === Batterieanalyse: Optimierung ===
echo.

if not exist "%CODE%\%SKRIPT%" (
    echo [FEHLER] "%CODE%\%SKRIPT%" nicht gefunden.
    echo          Liegt diese Datei im Hauptordner neben "05 Code"?
    goto :ende
)
if not exist "%SETUP%\conda_aktivieren.bat" (
    echo [FEHLER] "%SETUP%\conda_aktivieren.bat" nicht gefunden.
    goto :ende
)

call "%SETUP%\conda_aktivieren.bat"
if errorlevel 1 goto :ende

rem ---- Inputs-Datei: per Drag and Drop aufs Icon oder ins Fenster ------------
set "INPUT_FILE=%~1"
if defined INPUT_FILE goto :input_da
echo.
echo Inputs-Datei ins Fenster ziehen und Enter druecken:
set /p "INPUT_FILE=> "
:input_da
if not defined INPUT_FILE (
    echo [FEHLER] Kein Dateipfad angegeben.
    goto :ende
)
set "INPUT_FILE=%INPUT_FILE:"=%"
if not exist "%INPUT_FILE%" (
    echo [FEHLER] Datei nicht gefunden: "%INPUT_FILE%"
    goto :ende
)
for %%F in ("%INPUT_FILE%") do set "INPUT_DIR=%%~dpF"

rem ---- Vorlagen nicht direkt rechnen, Output-Ordner sicherstellen ------------
echo "%INPUT_DIR%" | findstr /I /L /C:"02 Vorlagen" >nul
if not errorlevel 1 (
    echo [FEHLER] Die Datei liegt im Ordner "02 Vorlagen".
    echo          Vorlagen bitte nicht direkt rechnen. Mit run_alles.bat und
    echo          Option [1] ein neues Projekt anlegen.
    goto :ende
)
if not exist "%INPUT_DIR%Output" (
    mkdir "%INPUT_DIR%Output"
    echo [INFO] Output-Ordner angelegt.
)

echo.
echo Starte %SKRIPT% mit:
echo   %INPUT_FILE%
echo.

cd /d "%CODE%"
python "%SKRIPT%" "%INPUT_FILE%"
if errorlevel 1 (
    echo.
    echo === FEHLER bei der Ausfuehrung, siehe Meldungen oben. ===
) else (
    echo.
    echo === Fertig. Ergebnisse liegen im Unterordner "Output" der Inputs-Datei. ===
)

:ende
echo.
pause
endlocal
