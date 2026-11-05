@echo off
chcp 65001 >nul
cd /d "%~dp0"

where py >nul 2>nul
if errorlevel 1 (
  echo Python wurde nicht gefunden. Bitte Python 3.11+ installieren ^(Company Portal oder python.org^).
  pause
  exit /b 1
)

echo Pruefe Abhaengigkeiten ...
py -m pip install --user --quiet --disable-pip-version-check -r requirements.txt
if errorlevel 1 (
  echo Installation der Abhaengigkeiten fehlgeschlagen.
  pause
  exit /b 1
)

echo.
set /p "QUELLE=SharePoint-Link (oder lokaler Ordnerpfad) einfuegen und Enter druecken: "
echo.
py -P agb_check.py "%QUELLE%" --oeffnen
echo.
pause
