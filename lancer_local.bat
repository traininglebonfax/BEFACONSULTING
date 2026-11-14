@echo off
REM Lance l'application Business Plan IA en local sur http://127.0.0.1:8000/
REM Rappel : mettre a jour ANTHROPIC_API_KEY dans .env AVANT de lancer (cle valable 24h).
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Creation de l'environnement virtuel...
    python -m venv .venv || goto :erreur
    ".venv\Scripts\python.exe" -m pip install --upgrade pip
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt || goto :erreur
)

if not exist ".env" (
    echo ATTENTION : fichier .env absent - les textes IA ne seront pas generes.
)

echo.
echo Application disponible sur http://127.0.0.1:8000/  (Ctrl+C pour arreter)
echo.
start "" http://127.0.0.1:8000/
".venv\Scripts\python.exe" -m uvicorn backend.app:app --port 8000
goto :eof

:erreur
echo Echec de l'installation. Verifier que Python est installe et dans le PATH.
pause
