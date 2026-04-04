@echo off
chcp 65001 >nul
echo ============================================
echo   Bulletins Scolaires — Demarrage
echo ============================================
echo.

:: Afficher l'adresse IP locale pour le mobile
echo Adresse pour votre mobile (meme Wi-Fi) :
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /i "IPv4" ^| findstr /v "127.0.0.1"') do (
    set IP=%%a
    setlocal enabledelayedexpansion
    set IP=!IP: =!
    echo   http://!IP!:7860
    endlocal
)
echo.
echo Appuyez sur Ctrl+C pour arreter le serveur.
echo ============================================
echo.

python app.py
pause
