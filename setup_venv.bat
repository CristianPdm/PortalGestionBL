@echo off
setlocal
echo.
echo ================================================================
echo   Gestion Stock Fiscal - GS -- Crear entorno virtual (venv)
echo ================================================================
echo.

REM ── Verificar Python ────────────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python no encontrado en el sistema.
    echo Instalar desde https://www.python.org/downloads/ y marcar "Add Python to PATH".
    pause
    exit /b 1
)
echo [OK] Python encontrado:
python --version
echo.

REM ── Crear el venv (solo si no existe ya) ─────────────────────────
if exist "venv\Scripts\python.exe" (
    echo [OK] El venv ya existe en .\venv  - no se vuelve a crear.
) else (
    echo Creando entorno virtual en .\venv ...
    python -m venv venv
    if errorlevel 1 (
        echo [ERROR] No se pudo crear el venv.
        pause
        exit /b 1
    )
    echo [OK] venv creado.
)
echo.

REM ── Instalar dependencias DENTRO del venv ────────────────────────
echo Instalando dependencias en el venv...
venv\Scripts\python.exe -m pip install --upgrade pip >nul
venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] No se pudieron instalar las dependencias.
    pause
    exit /b 1
)
echo [OK] Dependencias instaladas dentro del venv.
echo.

REM ── Inicializar base de datos (solo si no existe) ────────────────
if exist "gestion_stock_bl.db" (
    echo [OK] La base de datos ya existe - no se reinicializa.
) else (
    echo Inicializando base de datos...
    venv\Scripts\python.exe app.py --init-only
)

echo.
echo ================================================================
echo   LISTO
echo.
echo   Para trabajar en esta carpeta de ahora en mas:
echo     1. Activar el venv:   venv\Scripts\activate
echo     2. Iniciar el portal: python app.py
echo        (o directamente, sin activar):  venv\Scripts\python.exe app.py
echo ================================================================
echo.
pause
