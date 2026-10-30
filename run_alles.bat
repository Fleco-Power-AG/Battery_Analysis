@echo off
setlocal
chcp 65001 >nul
set "PYTHONIOENCODING=utf-8"
title Batterieanalyse - kompletter Ablauf

rem ============================================================================
rem  run_alles.bat  -  Batterieanalyse -> Sizing-Sweep -> Kundenbericht
rem
rem  Ablage : Hauptordner "17 Batterieanalyse"
rem  Start  : Inputs_<Name>.xlsx auf diese Datei ziehen ODER doppelklicken.
rem           Beim Doppelklick: [1] neues Projekt anlegen oder
rem                             [2] bestehende Inputs-Datei rechnen.
rem
rem  Neues Projekt: legt "01 Durchgefuehrte Analysen\JJJJ-MM <Name>\" an, kopiert
rem  die Vorlagen hinein, oeffnet sie zum Ausfuellen und rechnet danach direkt weiter.
rem
rem  Projektordner (Ordner der Inputs-Datei):
rem    Inputs_<Name>.xlsx                 Eingabe
rem    kunde_<Name>.json                  Kundendaten, wird bei Bedarf aus Vorlage erstellt
rem    Kundenbericht_<Name>.pdf + _QA.txt aus fleco_kundenreport.py
rem    Output\                            Bericht_, Ergebnis_, Sweep_ ... aus den Rechnungen
rem ============================================================================

set "ROOT=%~dp0"
set "CODE=%ROOT%05 Code"
set "SETUP=%ROOT%06 Setup"
rem Stammdaten (SwissIX, Solar_RMP, SRL-Marktpreise) zentral fuer alle Projekte
set "BATTERIE_STAMMDATEN_DIR=%ROOT%03 Stammdaten"
set "VORLAGE=%ROOT%02 Vorlagen\kunde_vorlage.json"

echo.
echo  =====================================================
echo   Batterieanalyse  -  kompletter Ablauf
echo  =====================================================
echo.

for %%D in ("run_battery_analysis.py" "battery_sizing.py" "fleco_kundenreport.py") do (
    if not exist "%CODE%\%%~D" (
        echo [FEHLER] "%CODE%\%%~D" nicht gefunden.
        goto :ende
    )
)
if not exist "%SETUP%\conda_aktivieren.bat" (
    echo [FEHLER] "%SETUP%\conda_aktivieren.bat" nicht gefunden.
    goto :ende
)

call "%SETUP%\conda_aktivieren.bat"
if errorlevel 1 goto :ende

rem ---------------------------------------------------------- Inputs-Datei ---
set "INPUT=%~1"
if defined INPUT goto :input_da
echo.
echo   [1] Neues Projekt anlegen
echo   [2] Bestehende Inputs-Datei rechnen
echo.
choice /C 12 /M "Auswahl"
if errorlevel 2 goto :frage_input
goto :neues_projekt
:frage_input
echo.
echo Inputs-Datei ins Fenster ziehen und Enter druecken:
set /p "INPUT=> "
:input_da
if not defined INPUT (
    echo [FEHLER] Kein Dateipfad angegeben.
    goto :ende
)
set "INPUT=%INPUT:"=%"
if not exist "%INPUT%" (
    echo [FEHLER] Datei nicht gefunden: "%INPUT%"
    goto :ende
)
for %%F in ("%INPUT%") do set "DIR=%%~dpF"
for %%F in ("%INPUT%") do set "BASE=%%~nF"
set "NAME=%BASE%"
if /i "%BASE:~0,7%"=="Inputs_" set "NAME=%BASE:~7%"

set "KUNDE=%DIR%kunde_%NAME%.json"
set "OUT=%DIR%Kundenbericht_%NAME%.pdf"
set "QA=%DIR%Kundenbericht_%NAME%_QA.txt"

echo.
echo  Projekt : %NAME%
echo  Ordner  : %DIR%
echo.

rem ---- Vorlagen nicht direkt rechnen, Output-Ordner sicherstellen ------------
echo "%DIR%" | findstr /I /L /C:"02 Vorlagen" >nul
if not errorlevel 1 (
    echo [FEHLER] Die Datei liegt im Ordner "02 Vorlagen".
    echo          Vorlagen bitte nicht direkt rechnen. Mit run_alles.bat und
    echo          Option [1] ein neues Projekt anlegen.
    goto :ende
)
if not exist "%DIR%Output" (
    mkdir "%DIR%Output"
    echo [INFO] Output-Ordner angelegt.
)

rem ------------------------------------- Kundendaten vor den Rechnungen klaeren ---
if exist "%KUNDE%" goto :kunde_da
if not exist "%VORLAGE%" (
    echo [FEHLER] Weder "%KUNDE%" noch die Vorlage "%VORLAGE%" gefunden.
    goto :ende
)
copy "%VORLAGE%" "%KUNDE%" >nul
echo [INFO] kunde_%NAME%.json wurde aus der Vorlage erstellt und wird geoeffnet.
echo        Auftraggeber, Projekt, Standort und Datum ausfuellen, speichern und
echo        Notepad schliessen. Danach geht es automatisch weiter.
start /wait notepad "%KUNDE%"
:kunde_da

rem --------------------------------------------------------------- Ablauf -----
choice /C JN /M "1) Batterieanalyse neu rechnen"
set "DO_ANA=%errorlevel%"
choice /C JN /M "2) Sizing-Sweep neu rechnen, dauert deutlich laenger"
set "DO_SWEEP=%errorlevel%"
echo.

for /f %%T in ('python -c "import time;print(int(time.time()))"') do set "START=%%T"
cd /d "%CODE%"

if "%DO_ANA%"=="2" goto :nach_ana
echo ---- Schritt 1: Batterieanalyse -------------------------------------------
python run_battery_analysis.py "%INPUT%"
if errorlevel 1 (
    echo [FEHLER] Batterieanalyse fehlgeschlagen. Ablauf abgebrochen.
    goto :ende
)
:nach_ana
call :finde_datei "Bericht_%NAME%.pdf" BERICHT
if not defined BERICHT (
    echo [FEHLER] Bericht_%NAME%.pdf nicht gefunden. Bitte Batterieanalyse rechnen.
    goto :ende
)
if "%DO_ANA%"=="1" (
    call :pruefe_neu "%BERICHT%" "Batterieanalyse"
    if errorlevel 1 goto :ende
)

if "%DO_SWEEP%"=="2" goto :nach_sweep
echo.
echo ---- Schritt 2: Sizing-Sweep ----------------------------------------------
python battery_sizing.py "%INPUT%"
if errorlevel 1 (
    echo [FEHLER] Sizing-Sweep fehlgeschlagen. Ablauf abgebrochen.
    goto :ende
)
:nach_sweep
call :finde_datei "Sweep_%NAME%.pdf" SWEEP
if "%DO_SWEEP%"=="1" (
    call :pruefe_neu "%SWEEP%" "Sizing-Sweep"
    if errorlevel 1 goto :ende
)

echo.
echo ---- Schritt 3: Kundenbericht ---------------------------------------------
if defined SWEEP goto :mit_sweep
echo [INFO] Kein Sweep vorhanden, Bericht ohne Kapitel Dimensionierung.
python fleco_kundenreport.py --bericht "%BERICHT%" --kunde "%KUNDE%" --out "%OUT%"
goto :nach_bericht
:mit_sweep
echo [INFO] Sweep gefunden, Kapitel Dimensionierung wird erstellt.
python fleco_kundenreport.py --bericht "%BERICHT%" --sweep "%SWEEP%" --kunde "%KUNDE%" --out "%OUT%"
:nach_bericht
if errorlevel 1 (
    echo [FEHLER] Kundenbericht konnte nicht erstellt werden.
    goto :ende
)

rem ------------------------------------------------------- Freigabe-Check -----
echo.
findstr /L /C:"[!! ]" "%QA%" >nul
if errorlevel 1 goto :qa_ok
echo  *****************************************************************
echo   ACHTUNG: Offene Punkte. Bericht vor dem Versand pruefen:
echo  *****************************************************************
findstr /L /C:"[!! ]" "%QA%"
goto :qa_fertig
:qa_ok
echo [OK] Alle Plausibilitaetschecks bestanden.
:qa_fertig
echo.
echo Fertig: "%OUT%"
start "" "%OUT%"
goto :ende

rem ------------------------------------------------------------------------------
rem  :finde_datei <Dateiname> <Variable>  ->  sucht in "Output\" und im Projekt-
rem  ordner, setzt die Variable auf die neueste gefundene Datei (sonst leer).
rem ------------------------------------------------------------------------------
:finde_datei
set "%~2="
for /f "delims=" %%P in ('python -c "import os,sys;c=[p for p in sys.argv[1:] if os.path.exists(p)];c and print(max(c,key=os.path.getmtime))" "%DIR%Output\%~1" "%DIR%%~1"') do set "%~2=%%P"
exit /b 0

rem ------------------------------------------------------------------------------
rem  :pruefe_neu <Datei> <Schritt>  ->  errorlevel 1, wenn die Datei fehlt oder
rem  nicht in diesem Lauf geschrieben wurde. Verhindert Berichte auf alten Daten.
rem ------------------------------------------------------------------------------
:pruefe_neu
python -c "import os,sys;p=sys.argv[1];sys.exit(0 if os.path.exists(p) and os.path.getmtime(p)>=int(sys.argv[2])-5 else 1)" "%~1" %START%
if errorlevel 1 (
    echo [FEHLER] %~2: "%~1" wurde nicht neu erstellt. Ablauf abgebrochen.
    exit /b 1
)
echo [OK] %~2 abgeschlossen.
exit /b 0

rem ------------------------------------------------------------------------------
rem  :neues_projekt  ->  Projektordner mit Vorlagen anlegen
rem ------------------------------------------------------------------------------
:neues_projekt
rem Analysen-Ordner ueber Platzhalter suchen (vermeidet Umlaut im Skript)
set "ANALYSEN="
for /d %%A in ("%ROOT%01 *") do if not defined ANALYSEN set "ANALYSEN=%%~fA"
if not defined ANALYSEN (
    echo [FEHLER] Ordner "01 ..." fuer die Analysen nicht gefunden.
    goto :ende
)

echo.
set "PNAME="
set /p "PNAME=Projektname, z.B. Meccano: "
if not defined PNAME (
    echo [FEHLER] Kein Projektname angegeben.
    goto :ende
)
set "FNAME=%PNAME: =_%"

rem ---- Inputs-Vorlage waehlen: bei mehreren Dateien Auswahl per Nummer ----
set "INPUTS_VORLAGE="
set "N=0"
echo.
echo Inputs-Vorlagen in "02 Vorlagen":
for %%V in ("%ROOT%02 Vorlagen\Inputs*.xlsx") do (
    set /a N+=1
    call set "VL_%%N%%=%%~fV"
    call echo   [%%N%%] %%~nxV
)
if "%N%"=="0" goto :vorlage_gewaehlt
if "%N%"=="1" (
    set "INPUTS_VORLAGE=%VL_1%"
    goto :vorlage_gewaehlt
)
set "WAHL="
set /p "WAHL=Nummer der Vorlage: "
if defined WAHL call set "INPUTS_VORLAGE=%%VL_%WAHL%%%"
if not defined INPUTS_VORLAGE (
    echo [FEHLER] Ungueltige Auswahl.
    goto :ende
)
:vorlage_gewaehlt
for /f %%M in ('python -c "import datetime;print(datetime.date.today().isoformat()[:7])"') do set "MONAT=%%M"
set "PDIR=%ANALYSEN%\%MONAT% %PNAME%"

if exist "%PDIR%" (
    echo [FEHLER] Den Projektordner gibt es bereits:
    echo          "%PDIR%"
    goto :ende
)
mkdir "%PDIR%"
mkdir "%PDIR%\Output"
copy "%VORLAGE%" "%PDIR%\kunde_%FNAME%.json" >nul
if defined INPUTS_VORLAGE (
    copy "%INPUTS_VORLAGE%" "%PDIR%\Inputs_%FNAME%.xlsx" >nul
) else (
    echo [WARNUNG] Keine Inputs-Vorlage "02 Vorlagen\Inputs*.xlsx" gefunden.
    echo           Inputs_%FNAME%.xlsx bitte selbst in den Projektordner legen.
)

rem Projektname und heutiges Datum in die Kundendaten vorbelegen
python -c "import json,sys,datetime;p=sys.argv[1];d=json.load(open(p,encoding='utf-8'));d.setdefault('projekt',{})['bezeichnung']=sys.argv[2];d['bericht_datum']=datetime.date.today().strftime('%%d.%%m.%%Y');json.dump(d,open(p,'w',encoding='utf-8'),ensure_ascii=False,indent=2)" "%PDIR%\kunde_%FNAME%.json" "%PNAME%"

echo.
echo [OK] Projektordner angelegt:
echo      "%PDIR%"
if not exist "%PDIR%\Inputs_%FNAME%.xlsx" (
    echo.
    echo Inputs_%FNAME%.xlsx bitte in den Projektordner legen und danach
    echo die Datei auf run_alles.bat ziehen.
    start "" "%PDIR%"
    goto :ende
)
start "" notepad "%PDIR%\kunde_%FNAME%.json"
start "" "%PDIR%\Inputs_%FNAME%.xlsx"

:warte_ausfuellen
echo.
echo  -----------------------------------------------------------------
echo   Jetzt ausfuellen und SPEICHERN:
echo     - kunde_%FNAME%.json   in Notepad  - Auftraggeber, Standort, Kontakt
echo     - Inputs_%FNAME%.xlsx  in Excel    - Parameter und Zeitreihen
echo   Excel danach schliessen und hier eine Taste druecken.
echo   Die Rechnung startet dann automatisch.
echo  -----------------------------------------------------------------
pause >nul
if exist "%PDIR%\~$*.xlsx" (
    echo [HINWEIS] Die Inputs-Datei ist noch in Excel geoeffnet.
    echo           Bitte speichern und Excel schliessen.
    goto :warte_ausfuellen
)
set "INPUT=%PDIR%\Inputs_%FNAME%.xlsx"
goto :input_da

:ende

:ende
echo.
pause
endlocal
