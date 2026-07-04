@echo off
title La Hogaza Madre - Sistema de Gestion
cd /d "%~dp0"

REM Verificar que Python esta instalado
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo ERROR: Python no esta instalado o no esta en el PATH.
    echo Descargalo desde: https://python.org
    pause
    exit /b 1
)

REM Verificar/instalar dependencias
python -c "import openpyxl" 2>nul
if %errorlevel% neq 0 (
    echo Instalando dependencia: openpyxl...
    pip install openpyxl
    if %errorlevel% neq 0 (
        echo ERROR: No se pudo instalar openpyxl.
        echo Ejecuta manualmente: pip install openpyxl
        pause
        exit /b 1
    )
)

REM Verificar que config_default.json existe
if not exist "config_default.json" (
    echo ERROR: Falta el archivo config_default.json
    echo Asegurate de que este en la misma carpeta que el programa.
    pause
    exit /b 1
)

echo.
echo Iniciando La Hogaza Madre...
echo.
python hogaza_madre.py
if %errorlevel% neq 0 (
    echo.
    echo ERROR: La aplicacion se cerro inesperadamente.
    echo Revisa los mensajes de error en pantalla.
    pause
)
