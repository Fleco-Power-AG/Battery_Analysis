@echo off
rem ============================================================================
rem  conda_aktivieren.bat  -  gemeinsamer Baustein fuer alle run_*.bat
rem
rem  Ablage : 06 Setup\
rem  Aufruf : call "%ROOT%06 Setup\conda_aktivieren.bat"
rem           if errorlevel 1 goto :ende
rem
rem  Findet Anaconda/Miniconda auch dann, wenn "conda" im Fenster unbekannt ist
rem  (typisch bei Doppelklick), und aktiviert das Environment.
rem  Bewusst OHNE setlocal, damit die Aktivierung im aufrufenden Skript bleibt.
rem ============================================================================

if not defined ENV_NAME set "ENV_NAME=batterieanalyse"

where conda >nul 2>nul
if not errorlevel 1 goto :env_aktivieren

echo "conda" ist in diesem Fenster unbekannt - suche Installation ...
set "_CONDA_ROOT="
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
    if not defined _CONDA_ROOT if exist "%%~P\Scripts\conda.exe" set "_CONDA_ROOT=%%~P"
)
if not defined _CONDA_ROOT goto :fehler_conda

echo Gefunden: %_CONDA_ROOT%
call "%_CONDA_ROOT%\Scripts\activate.bat" "%_CONDA_ROOT%"
where conda >nul 2>nul
if errorlevel 1 goto :fehler_conda

:env_aktivieren
echo Aktiviere Environment "%ENV_NAME%" ...
call conda activate %ENV_NAME%
if errorlevel 1 goto :fehler_env
exit /b 0

:fehler_conda
echo.
echo [FEHLER] Conda wurde nicht gefunden.
echo          Bitte zuerst "06 Setup\setup.bat" ausfuehren oder das Skript
echo          aus dem "Anaconda Prompt" starten.
exit /b 1

:fehler_env
echo.
echo [FEHLER] Environment "%ENV_NAME%" nicht gefunden.
echo          Bitte zuerst "06 Setup\setup.bat" ausfuehren.
exit /b 1
