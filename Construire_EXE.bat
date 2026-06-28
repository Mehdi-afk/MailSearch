@echo off
REM ============================================================
REM   Construction de l'application Recherche Outlook en .exe
REM   Double-cliquez sur ce fichier sur votre PC Windows.
REM ============================================================
chcp 65001 >nul
cd /d "%~dp0"

echo.
echo === 1/3 : Verification de Python ===
python --version
if errorlevel 1 (
    echo.
    echo [ERREUR] Python n'est pas installe ou pas dans le PATH.
    echo Installez-le depuis https://www.python.org/downloads/
    echo en cochant "Add Python to PATH", puis relancez ce fichier.
    pause
    exit /b 1
)

echo.
echo === 2/3 : Installation des modules (pywin32 + pyinstaller) ===
python -m pip install --upgrade pip
python -m pip install pywin32 pyinstaller
if errorlevel 1 (
    echo.
    echo [ERREUR] L'installation des modules a echoue.
    pause
    exit /b 1
)

echo.
echo === 3/3 : Construction du .exe ===
REM --onedir (et non --onefile) + --noupx : nettement moins souvent
REM bloque par Windows Defender / SmartScreen.
python -m PyInstaller --onedir --windowed --noupx --clean --name "RechercheOutlook" recherche_outlook.py
if errorlevel 1 (
    echo.
    echo [ERREUR] La construction a echoue.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   TERMINE !
echo   Votre application se trouve dans :
echo   %~dp0dist\RechercheOutlook\RechercheOutlook.exe
echo.
echo   IMPORTANT - si Windows affiche "Windows a protege votre PC" :
echo   cliquez sur "Informations complementaires" puis "Executer
echo   quand meme". C'est un faux positif (exe non signe).
echo.
echo   CONSEIL : le plus simple est d'utiliser "Lancer_Recherche.bat"
echo   qui demarre l'appli sans .exe et sans aucun blocage.
echo ============================================================
echo.
pause
