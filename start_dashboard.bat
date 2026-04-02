@echo off
echo Iniciando Manufacturing Dashboard...
echo.
pip install flask pyodbc pandas --quiet
python app.py
pause
