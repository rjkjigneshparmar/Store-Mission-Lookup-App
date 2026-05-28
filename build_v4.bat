@echo off
title Store Mission Lookup v4 — Build

echo.
echo ============================================
echo   Store Mission Lookup v4 — Build to .exe
echo ============================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found.
    pause & exit /b 1
)

echo [1/3] Installing dependencies...
pip install customtkinter pyinstaller openpyxl --quiet
if errorlevel 1 ( echo ERROR: pip install failed. & pause & exit /b 1 )
echo Done.
echo.

echo [2/3] Building executable...
pyinstaller ^
  --onefile ^
  --windowed ^
  --collect-all customtkinter ^
  --hidden-import openpyxl ^
  --name "StoreMissionLookup_v4" ^
  --clean ^
  app_v4.py

if errorlevel 1 ( echo ERROR: Build failed. & pause & exit /b 1 )
echo.

echo [3/3] Cleaning up...
rmdir /s /q build >nul 2>&1
del /q StoreMissionLookup_v4.spec >nul 2>&1
echo Done.
echo.

echo ============================================
echo   SUCCESS!
echo   App:       dist\StoreMissionLookup_v4.exe
echo   Settings:  %APPDATA%\StoreMissionLookup\settings.json
echo   Database:  %APPDATA%\StoreMissionLookup\data.db (default)
echo ============================================
echo.
pause
