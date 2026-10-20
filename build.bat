@echo off
chcp 65001 >nul
title G360 - Outlook Extractor Builder

echo.
echo ==============================================
echo   G360 Outlook Extractor - Build
echo ==============================================
echo.

where uv >nul 2>&1
if errorlevel 1 (
    echo [ERROR] uv no encontrado
    echo.
    echo Por favor instala uv desde:
    echo https://github.com/astral-sh/uv
    echo.
    pause
    exit /b 1
)

echo [CHECK] uv: OK
echo.

echo [SETUP] Sincronizando dependencias...
uv sync

echo.
echo [BUILD] Construyendo ejecutable portatil...
uv run python -m PyInstaller --noconfirm --clean --onefile --name "G360_Outlook_Extractor" --icon "assets\images\favicon.ico" --add-data "assets;assets" --collect-all flet --hidden-import flet_desktop --hidden-import g360_flet.g360_signature --hidden-import win32timezone --paths src main.py

echo.
if exist "dist\G360_Outlook_Extractor.exe" (
    echo [OK] Build completado exitosamente
    echo [INFO] Ejecutable: dist\G360_Outlook_Extractor.exe
) else (
    echo [ERROR] Build fallo
)

echo.
pause