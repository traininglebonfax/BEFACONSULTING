@echo off
REM Verification avant demonstration client.
REM   verifier_demo.bat       : controle rapide (~1 min, sans IA sauf un appel test de la cle)
REM   verifier_demo.bat --ia  : repetition complete (~5 min, consomme du credit API)
REM Rappel : mettre a jour ANTHROPIC_API_KEY dans .env AVANT (cle valable 24h).
cd /d "%~dp0"
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
if not exist ".venv\Scripts\python.exe" (
    echo Environnement absent : lancer d'abord lancer_local.bat
    pause
    exit /b 1
)
".venv\Scripts\python.exe" -m scripts.verifier_demo %*
pause
