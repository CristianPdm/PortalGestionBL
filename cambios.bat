@echo off
setlocal

REM ── Uso: cambios.bat "descripcion del cambio" ────────────────────
if "%~1"=="" (
    echo.
    echo [ERROR] Falta la descripcion del cambio.
    echo.
    echo Uso:   cambios.bat "descripcion del cambio"
    echo Ejemplo: cambios.bat "Corrijo validacion de CUIT"
    echo.
    pause
    exit /b 1
)

echo.
echo ================================================================
echo   Subiendo cambios a GitHub...
echo   Descripcion: %~1
echo ================================================================
echo.

git add .
git commit -m "%~1"
git push

echo.
if errorlevel 1 (
    echo [ERROR] No se pudieron subir los cambios. Ver mensaje arriba.
) else (
    echo [OK] Cambios subidos correctamente.
)
echo ================================================================
echo.
pause
