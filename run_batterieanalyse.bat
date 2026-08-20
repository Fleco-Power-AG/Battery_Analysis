@echo off
setlocal

rem run_battery_analysis.bat
rem =========================
rem Doppelklick-Startpunkt fuer die Batterieanalyse.
rem
rem Verwendung:
rem   - Einfach doppelklicken -> es oeffnet sich ein Datei-Auswahl-Dialog,
rem     in dem du deine Inputs.xlsx auswaehlst (egal wo sie liegt).
rem   - Oder: deine Inputs.xlsx-Datei auf dieses .bat-Icon ziehen
rem     (Drag & Drop) -> der Pfad wird automatisch uebernommen.
rem
rem Dieses Skript findet automatisch eine lokale virtuelle Umgebung
rem (.venv oder venv im selben Ordner), falls vorhanden -- sonst wird das
rem system-globale "python" verwendet. Falls deine virtuelle Umgebung
rem anders heisst, passe PYTHON_EXE unten entsprechend an.

rem In den Ordner wechseln, in dem diese .bat-Datei liegt (damit die
rem Python-Module unabhaengig vom Ausfuehrungsort gefunden werden).
cd /d "%~dp0"

set PYTHON_EXE=python
if exist "%~dp0.venv\Scripts\python.exe" set PYTHON_EXE=%~dp0.venv\Scripts\python.exe
if exist "%~dp0venv\Scripts\python.exe" set PYTHON_EXE=%~dp0venv\Scripts\python.exe

echo Verwende Python: %PYTHON_EXE%
echo.

"%PYTHON_EXE%" run_battery_analysis.py %*

echo.
echo ---------------------------------------------------------
echo Fertig (oder Fehler siehe oben). Fenster schliesst sich mit
echo einem beliebigen Tastendruck.
echo ---------------------------------------------------------
pause >nul