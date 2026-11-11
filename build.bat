@echo off
REM ============================================================
REM  Euro Superstore Boekhouding - Windows installer bouwen
REM  Dubbelklik dit bestand op een Windows 10/11 computer.
REM ============================================================
title Boekhouding - app bouwen
echo.
echo   ============================================
echo    Euro Superstore Boekhouding bouwen
echo   ============================================
echo.

where node >nul 2>nul
if errorlevel 1 (
  echo   [FOUT] Node.js is niet gevonden.
  echo.
  echo   Installeer eerst Node.js ^(versie 18 of hoger^):
  echo   https://nodejs.org/  -^>  klik "LTS" en installeer.
  echo.
  echo   Sluit daarna dit venster en dubbelklik opnieuw op build.bat
  echo.
  pause
  exit /b
)

echo   Node.js gevonden. Onderdelen installeren...
echo   (dit duurt de eerste keer enkele minuten)
echo.
call npm install
if errorlevel 1 (
  echo.
  echo   [FOUT] Installeren mislukt. Controleer je internetverbinding.
  pause
  exit /b
)

echo.
echo   Installer bouwen...
echo.
call npm run dist
if errorlevel 1 (
  echo.
  echo   [FOUT] Bouwen mislukt. Zie de melding hierboven.
  pause
  exit /b
)

echo.
echo   ============================================
echo    KLAAR!
echo   ============================================
echo.
echo   Je vindt de app in de map "dist":
echo     - "Euro Superstore Boekhouding Setup 1.2.0.exe"  (installeren)
echo     - "Boekhouding-portable.exe"  (zonder installeren starten)
echo.
explorer dist
pause
