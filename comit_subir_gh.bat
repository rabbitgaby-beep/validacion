
Copiar

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
 
echo.
echo [1/3] Agregando archivos...
git add .
 
echo [2/3] Haciendo commit...
git commit -m "%COMENTARIO%"
 
echo [3/3] Haciendo push...
git push origin master
 
echo.
echo ================================
echo     Push completado con exito!
echo ================================
echo.
pause