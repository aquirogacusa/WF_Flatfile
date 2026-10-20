@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo  Maestro Clientes - Merge Customers_SAP + Z5
echo ============================================
python -c "import pandas, openpyxl, dotenv" 2>nul || pip install -r requirements.txt
python customers_agent.py
echo.
if %errorlevel% neq 0 (
    echo *** Ocurrio un error, revisa los mensajes de arriba ***
) else (
    echo *** Proceso terminado correctamente ***
)
pause
