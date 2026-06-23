@echo off
chcp 65001 >nul
title G360 - Outlook Extractor

echo.
echo ==============================================
echo   G360 Outlook Extractor
echo ==============================================
echo.

where uv >nul 2>&1
if errorlevel 1 (
    echo [SETUP] uv no encontrado
    echo.
    echo Por favor instala uv desde:
    echo https://github.com/astral-sh/uv
    echo.
    echo O ejecuta en PowerShell:
    echo   powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
    echo.
    pause
    exit /b 1
)

echo [CHECK] uv: OK

if not exist ".venv" (
    echo [SETUP] Creando entorno virtual...
    uv venv
)

echo [SETUP] Sincronizando dependencias...
uv sync

echo.
echo [RUN] Ejecutando aplicacion...
uv run python main.py