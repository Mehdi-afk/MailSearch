@echo off
REM ============================================================
REM   Lance l'application Recherche emails SANS .exe
REM   (evite tout blocage SmartScreen / antivirus)
REM   Double-cliquez simplement sur ce fichier.
REM ============================================================
cd /d "%~dp0"

REM pythonw = lance l'appli sans fenetre noire
start "" pythonw "recherche_outlook.py"
if errorlevel 1 (
    REM Si pythonw n'est pas trouve, on tente python
    start "" python "recherche_outlook.py"
)
exit
