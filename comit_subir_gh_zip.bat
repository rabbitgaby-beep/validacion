@echo off
chcp 65001 >nul
echo.
echo ================================
echo        GIT COMMIT ^& PUSH
echo ================================
echo.

git status

echo.
set /p COMENTARIO="Ingresa el comentario del commit: "

:: ── Fecha y hora ──────────────────────────────────────────
for /f "tokens=1-3 delims=/" %%a in ("%date%") do (
    set DIA=%%a
    set MES=%%b
    set ANIO=%%c
)
set ANIO=%ANIO:~2,2%

for /f "tokens=1-2 delims=:" %%a in ("%time: =0%") do (
    set HORA=%%a
    set MIN=%%b
)

set FECHA_HORA=%DIA%_%MES%_%ANIO%_%HORA%_%MIN%

:: ── Rutas ─────────────────────────────────────────────────
set CARPETA=K:\programacion_ia\Validacion_Template_1
set NOMBRE_ZIP=Validacion_Template_1_%COMENTARIO%_%FECHA_HORA%.zip
set RUTA_ZIP=%CARPETA%\%NOMBRE_ZIP%

echo.
echo [1/4] Agregando archivos...
git add .

echo [2/4] Haciendo commit...
git commit -m "%COMENTARIO%"

echo [3/4] Haciendo push...
git push origin master

echo [4/4] Creando ZIP backup...
echo.
echo  Destino: %RUTA_ZIP%
echo.

powershell -NoProfile -Command "Compress-Archive -Path '%CARPETA%\*' -DestinationPath '%RUTA_ZIP%' -Force"

if %ERRORLEVEL% == 0 (
    echo.
    echo ================================
    echo   Backup ZIP creado con exito!
    echo   %NOMBRE_ZIP%
    echo ================================
) else (
    echo.
    echo  ERROR: No se pudo crear el ZIP.
)

echo.
echo ================================
echo     Todo completado con exito!
echo ================================
echo.
pause