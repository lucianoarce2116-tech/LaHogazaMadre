@echo off
title La Hogaza Madre - Configurar GitHub (SOLO PC PRINCIPAL)
cd /d "%~dp0"

REM ============================================================
REM  SETUP_GITHUB.bat
REM  Configura el repositorio Git en la PC principal
REM  y lo sube a GitHub por primera vez
REM  USO: Ejecutar UNA SOLA VEZ en la PC principal
REM ============================================================

echo ============================================
echo  Configurar GitHub — PC Principal
echo ============================================
echo.
echo Este script configura el control de versiones
echo y te guia para subir el programa a GitHub.
echo.

REM Verificar Git
git --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Git no esta instalado.
    echo Descargalo desde: https://git-scm.com/download/win
    echo Instalalo y ejecuta este script nuevamente.
    pause
    exit /b 1
)

REM Verificar si ya hay un repositorio
if exist ".git" (
    echo El repositorio ya esta inicializado.
    echo Si queres reconectarlo a otro GitHub, borra la carpeta .git
    git remote -v
    pause
    exit /b 0
)

echo PASO 1: Inicializando repositorio local...
git init
if %errorlevel% neq 0 (
    echo ERROR: No se pudo inicializar el repositorio.
    pause
    exit /b 1
)
echo OK.
echo.

echo PASO 2: Agregando archivos al repositorio...
git add .
echo OK.
echo.

echo PASO 3: Creando el primer commit...
git commit -m "Version inicial - La Hogaza Madre v2.1"
echo OK.
echo.
echo ============================================
echo  REPOSITORIO LOCAL CREADO CON EXITO
echo ============================================
echo.
echo Ahora falta subirlo a GitHub.
echo.
echo Segui estos pasos:
echo.
echo  1. Crea una cuenta en https://github.com/signup
echo.
echo  2. Una vez registrada, crea un repositorio:
echo     - Ve a https://github.com/new
echo     - Nombre del repositorio: LaHogazaMadre
echo     - Dejalo en PUBLICO (gratis)
echo     - NO marques ninguna opcion (README, .gitignore, etc.)
echo     - Haz clic en "Create repository"
echo.
echo  3. En la pagina que aparece, busca la seccion:
echo     "...or push an existing repository from the command line"
echo     y ejecuta los comandos que aparecen.
echo.
echo     Generalmente son estos dos:
echo       git remote add origin https://github.com/TU_USUARIO/LaHogazaMadre.git
echo       git branch -M main
echo       git push -u origin main
echo.
echo  IMPORTANTE: Reemplaza TU_USUARIO por tu nombre de usuario de GitHub.
echo.
echo  Una vez hecho esto, en las otras PCs solo necesitas:
echo   1. Instalar Git
echo   2. Ejecutar Actualizar.bat (te guiara la primera vez)
echo.
pause
