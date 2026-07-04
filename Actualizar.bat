@echo off
title La Hogaza Madre - Actualizador
cd /d "%~dp0"

REM ============================================================
REM  ACTUALIZAR.bat
REM  Sincroniza la ultima version del programa desde GitHub
REM  USO: Hacer doble clic cuando quieras actualizar
REM ============================================================

echo ============================================
echo  La Hogaza Madre — Actualizador
echo ============================================
echo.

REM Verificar que Git esta instalado
git --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Git no esta instalado.
    echo.
    echo Descargalo desde: https://git-scm.com/download/win
    echo Instalalo con todas las opciones por defecto.
    echo Vuelve a ejecutar este script despues de instalar Git.
    pause
    exit /b 1
)

REM Verificar si el repositorio ya esta clonado
if not exist ".git" (
    echo.
    echo === PRIMERA VEZ EN ESTA PC ===
    echo.
    echo Esta carpeta aun no esta vinculada a GitHub.
    echo.
    echo Para la primera instalacion segui estos pasos:
    echo.
    echo  1. Abri una terminal (cmd) en esta carpeta
    echo  2. Ejecuta: git clone ^<URL_DEL_REPO^> temp
    echo  3. Ejecuta: xcopy /E /Y temp\* .
    echo  4. Ejecuta: rmdir /S /Q temp
    echo  5. Copia tu config_negocio.json y stock si los tenias
    echo.
    echo Despues de eso, este script ya funcionara solo.
    echo.
    pause
    exit /b 0
)

echo Conectando con GitHub...
echo.

REM Guardar cambios locales por si hubo modificaciones accidentales
git stash push -m "cambios_locales_previos_actualizacion" 2>nul

REM Traer ultimos cambios
echo Descargando actualizaciones...
git pull origin main 2>nul
if %errorlevel% neq 0 (
    git pull origin master 2>nul
)

if %errorlevel% equ 0 (
    echo.
    echo ============================================
    echo  ACTUALIZACION COMPLETADA CON EXITO
    echo ============================================
    echo.
    echo Ultimos cambios recibidos:
    git log --oneline -3
) else (
    echo.
    echo ERROR: No se pudieron descargar las actualizaciones.
    echo Posibles causas:
    echo   - Sin conexion a internet
    echo   - El repositorio no existe o cambió la URL
    echo   - No tienes permisos de acceso
    echo.
    echo Solucion: Ejecuta manualmente: git pull
)

echo.
pause
