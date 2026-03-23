@echo off
chcp 65001 > nul
title Agent Testeur QA — Elio OnePoint
color 0A

cd /d "%~dp0"

:: Charger les variables du fichier .env si présent
if exist .env (
    for /f "usebackq tokens=1,* delims==" %%A in (".env") do (
        if not "%%A"=="" if not "%%A:~0,1%"=="#" set "%%A=%%B"
    )
)

:: Vérifier que la clé est définie
if "%ANTHROPIC_API_KEY%"=="" (
    echo.
    echo  ERREUR : ANTHROPIC_API_KEY non definie.
    echo  Copie .env.example en .env et mets ta cle API dedans.
    echo.
    pause
    exit /b 1
)

echo.
echo  Agent Testeur QA pret.
echo  Excel et Edge doivent etre ouverts avant de continuer.
echo.
pause

python agent_testeur.py

echo.
pause
