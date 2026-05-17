@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

:: ============================================================
::  GIT REVERT - Herramienta para revertir cambios en Git
:: ============================================================

:MENU
cls
echo ============================================================
echo   GIT REVERT - Revertir cambios en el repositorio
echo ============================================================
echo.
echo   Repositorio: %cd%
echo.

:: Verificar que estamos en un repositorio git
git rev-parse --is-inside-work-tree >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Esta carpeta no es un repositorio Git.
    echo  Navega hasta la carpeta correcta y vuelve a ejecutar.
    pause
    exit /b 1
)

:: Mostrar estado actual
echo  Estado actual del repositorio:
echo  ------------------------------------------------------------
git status --short
echo  ------------------------------------------------------------
echo.
echo  Selecciona una opcion:
echo.
echo  --- Comandos rapidos ---
echo  [1] git restore .          ^<-- descarta cambios tracked (sin stagear)
echo  [2] git restore --staged . ^<-- quita TODO del staging area
echo  [3] git clean -fd          ^<-- elimina archivos y carpetas untracked
echo  [4] git restore . + clean  ^<-- limpieza total (tracked + untracked)
echo.
echo  --- Control de commits ---
echo  [5] Revertir un commit especifico (git revert)
echo  [6] Volver a un commit anterior  (git reset --soft/mixed/hard)
echo  [7] Restaurar un archivo especifico
echo.
echo  --- Info ---
echo  [8] Ver historial de commits recientes
echo  [0] Salir
echo.
set /p OPCION="  Tu eleccion: "

if "%OPCION%"=="1" goto RESTORE_ALL
if "%OPCION%"=="2" goto RESTORE_STAGED
if "%OPCION%"=="3" goto CLEAN_FD
if "%OPCION%"=="4" goto FULL_CLEAN
if "%OPCION%"=="5" goto REVERT_COMMIT
if "%OPCION%"=="6" goto RESET_COMMIT
if "%OPCION%"=="7" goto RESTORE_FILE
if "%OPCION%"=="8" goto SHOW_LOG
if "%OPCION%"=="0" goto FIN
goto MENU

:: ============================================================
:: [1] git restore . — descarta cambios en tracked (no staged)
:: ============================================================
:RESTORE_ALL
cls
echo ============================================================
echo  [1] git restore .
echo  Descarta cambios en archivos tracked que NO estan en stage
echo ============================================================
echo.
echo  Archivos afectados:
git diff --name-only
echo.
echo  Comando que se ejecutara:  git restore .
echo  ATENCION: irreversible, los cambios se perderan.
set /p CONFIRM="  Confirmas? (s/n): "
if /i "%CONFIRM%"=="s" (
    git restore .
    echo.
    echo  [OK] git restore . ejecutado.
) else (
    echo  Cancelado.
)
echo.
pause
goto MENU

:: ============================================================
:: [2] git restore --staged . — quita todo del staging area
:: ============================================================
:RESTORE_STAGED
cls
echo ============================================================
echo  [2] git restore --staged .
echo  Quita todos los archivos del staging (los cambios se conservan)
echo ============================================================
echo.
echo  Archivos en staging:
git diff --cached --name-only
echo.
echo  Comando que se ejecutara:  git restore --staged .
echo  (Los cambios en disco NO se pierden, solo salen del stage)
set /p CONFIRM="  Confirmas? (s/n): "
if /i "%CONFIRM%"=="s" (
    git restore --staged .
    echo.
    echo  [OK] git restore --staged . ejecutado.
) else (
    echo  Cancelado.
)
echo.
pause
goto MENU

:: ============================================================
:: [3] git clean -fd — elimina archivos y carpetas untracked
:: ============================================================
:CLEAN_FD
cls
echo ============================================================
echo  [3] git clean -fd
echo  Elimina archivos (-f) y carpetas (-d) no rastreados
echo ============================================================
echo.
echo  Vista previa de lo que se eliminara (dry-run):
git clean -nfd
echo.
echo  Comando que se ejecutara:  git clean -fd
echo  ATENCION: irreversible, estos archivos NO van a la papelera.
set /p CONFIRM="  Confirmas? Escribe SI para continuar: "
if "%CONFIRM%"=="SI" (
    git clean -fd
    echo.
    echo  [OK] git clean -fd ejecutado.
) else (
    echo  Cancelado.
)
echo.
pause
goto MENU

:: ============================================================
:: [4] Limpieza total: git restore . + git clean -fd
:: ============================================================
:FULL_CLEAN
cls
echo ============================================================
echo  [4] Limpieza total del working tree
echo  git restore --staged .  +  git restore .  +  git clean -fd
echo ============================================================
echo.
echo  Esto hara:
echo   1. git restore --staged .  -> quita todo del staging
echo   2. git restore .           -> descarta cambios en tracked
echo   3. git clean -fd           -> borra archivos/carpetas untracked
echo.
echo  Estado actual:
git status --short
echo.
echo  Vista previa untracked (git clean -nfd):
git clean -nfd
echo.
echo  ATENCION: IRREVERSIBLE. El working tree quedara identico al ultimo commit.
set /p CONFIRM="  Confirmas? Escribe SI para continuar: "
if "%CONFIRM%"=="SI" (
    git restore --staged .
    git restore .
    git clean -fd
    echo.
    echo  [OK] Working tree limpio. Equivalente a un checkout fresco.
) else (
    echo  Cancelado.
)
echo.
pause
goto MENU

:: ============================================================
:: [5] Revertir un commit especifico (crea un nuevo commit)
:: ============================================================
:REVERT_COMMIT
cls
echo ============================================================
echo  [5] Revertir un commit especifico
echo ============================================================
echo.
echo  Commits recientes:
git log --oneline -10
echo.
set /p HASH="  Ingresa el hash del commit a revertir: "
if "%HASH%"=="" goto MENU
echo.
echo  Se creara un nuevo commit que revierte: %HASH%
set /p CONFIRM="  Confirmas? (s/n): "
if /i "%CONFIRM%"=="s" (
    git revert %HASH% --no-edit
    echo.
    echo  [OK] Commit revertido correctamente.
) else (
    echo  Operacion cancelada.
)
echo.
pause
goto MENU

:: ============================================================
:: [6] Volver a un commit anterior (git reset)
:: ============================================================
:RESET_COMMIT
cls
echo ============================================================
echo  [6] Volver a un commit anterior (git reset)
echo ============================================================
echo.
echo  Commits recientes:
git log --oneline -10
echo.
set /p HASH="  Ingresa el hash del commit destino: "
if "%HASH%"=="" goto MENU
echo.
echo  Tipo de reset:
echo  [1] --soft   (mantiene cambios en staging)
echo  [2] --mixed  (mantiene cambios en archivos, quita del staging)
echo  [3] --hard   (ELIMINA todos los cambios - irreversible)
set /p TIPO="  Tu eleccion: "

if "%TIPO%"=="1" (
    git reset --soft %HASH%
    echo  [OK] Reset --soft aplicado.
)
if "%TIPO%"=="2" (
    git reset --mixed %HASH%
    echo  [OK] Reset --mixed aplicado.
)
if "%TIPO%"=="3" (
    echo.
    echo  ATENCION: --hard eliminara todos los cambios despues del commit.
    set /p CONFIRM="  Confirmas? Escribe SI: "
    if "!CONFIRM!"=="SI" (
        git reset --hard %HASH%
        echo  [OK] Reset --hard aplicado.
    ) else (
        echo  Cancelado.
    )
)
echo.
pause
goto MENU

:: ============================================================
:: [7] Restaurar un archivo especifico
:: ============================================================
:RESTORE_FILE
cls
echo ============================================================
echo  [7] Restaurar un archivo especifico
echo ============================================================
echo.
echo  Archivos modificados:
git status --short
echo.
set /p ARCHIVO="  Nombre o ruta del archivo a restaurar: "
if "%ARCHIVO%"=="" goto MENU
echo.
echo  [1] Restaurar al ultimo commit
echo  [2] Restaurar desde un commit especifico
set /p SUB="  Tu eleccion: "

if "%SUB%"=="1" (
    git checkout -- "%ARCHIVO%"
    echo  [OK] Archivo restaurado al ultimo commit.
)
if "%SUB%"=="2" (
    git log --oneline -10
    echo.
    set /p HASH="  Hash del commit: "
    git checkout !HASH! -- "%ARCHIVO%"
    echo  [OK] Archivo restaurado desde el commit !HASH!.
)
echo.
pause
goto MENU

:: ============================================================
:: [8] Ver historial de commits
:: ============================================================
:SHOW_LOG
cls
echo ============================================================
echo  [8] Historial de commits recientes
echo ============================================================
echo.
git log --oneline --graph --decorate -20
echo.
pause
goto MENU

:: ============================================================
:FIN
cls
echo  Hasta luego!
echo.
endlocal
exit /b 0