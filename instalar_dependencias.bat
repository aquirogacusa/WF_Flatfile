@echo off
echo Instalando dependencias para el Agente Wells Fargo...
echo.
python -m pip install --upgrade pip
pip install -r requirements.txt
python -m pip install pywin32 python-dotenv --quiet
echo.
echo Instalacion finalizada.
pause
