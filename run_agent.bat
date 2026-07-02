@echo off
setlocal

:: ==========================================================
:: CONFIGURACION DE RUTA
:: Cambie la siguiente linea por su ruta de red real
:: Ejemplo: set "FILES_PATH=\\Servidor\Carpeta\WellsFargo"
:: ==========================================================
set "FILES_PATH=C:\AIWorkspace\WF_Flatfile"

echo Ejecutando Agente Wells Fargo...
echo Ruta de archivos: %FILES_PATH%
echo.

python wf_customer_agent.py --path "%FILES_PATH%"

echo.
echo Proceso finalizado.
pause
