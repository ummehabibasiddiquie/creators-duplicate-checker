@echo off
cd /d "%~dp0"
python -c "import flask, openpyxl, pymysql" 2>nul
if errorlevel 1 python -m pip install -r requirements.txt
python app.py --open
echo.
echo The checker has stopped.
pause
