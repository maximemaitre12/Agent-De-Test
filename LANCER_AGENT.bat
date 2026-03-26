@echo off
chcp 65001 > nul
title Agent Testeur QA — Elio OnePoint
color 0A

cd /d "%~dp0"

echo.
echo  =========================================================
echo   AGENT TESTEUR QA — Sans IA (Playwright)
echo  =========================================================
echo.
echo  Verifie que les dependances sont installees :
echo    pip install playwright openpyxl
echo    playwright install chromium
echo.

:: Verifier que Python est disponible
python --version > nul 2>&1
if errorlevel 1 (
    echo  ERREUR : Python introuvable. Installe Python 3.10+ et relance.
    pause
    exit /b 1
)

:: Verifier que Playwright est installe
python -c "import playwright" > nul 2>&1
if errorlevel 1 (
    echo  Installation de Playwright...
    pip install playwright openpyxl
    playwright install chromium
)

echo  Lancement de l'agent...
echo.
python agent_testeur.py

echo.
pause
