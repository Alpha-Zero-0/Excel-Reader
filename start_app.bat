@echo off
echo Starting Excel Reader Tool...
echo Once started, open your browser and go to: http://localhost:8501
echo.

REM Navigate to the script directory
cd /d "%~dp0"

REM Run the Streamlit application
python -m streamlit run excel_reader.py --server.address localhost --server.port 8501

echo.
echo Excel Reader Tool has stopped.
pause
