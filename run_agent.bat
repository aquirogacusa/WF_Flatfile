@echo off
setlocal

:: ==========================================================
:: CONFIGURACION DE RUTAS
:: PROG_PATH: carpeta LOCAL donde viven este .bat y los .py.
::            No debe ser una ruta de red: si se ejecuta desde
::            la red, Windows mostrara el aviso de SmartScreen
::            en cada ejecucion.
:: DATA_PATH: carpeta de los DATOS (zip/csv), configurada en .env.
:: ==========================================================
set "PROG_PATH=%~dp0"
set "PYTHON_EXE=C:\Users\artur\AppData\Local\Python\bin\python.exe"

echo Ejecutando Agente Wells Fargo...
echo Carpeta de programa: %PROG_PATH%
echo Carpeta de datos:    configurada en .env
echo.

cd /d "%PROG_PATH%"

if exist "%PYTHON_EXE%" (
    "%PYTHON_EXE%" wf_customer_agent.py
) else (
    where py >nul 2>nul
    if errorlevel 1 (
        python wf_customer_agent.py
    ) else (
        py -3 wf_customer_agent.py
    )
)

echo.
echo Proceso finalizado. Codigo de salida: %errorlevel%
pause
