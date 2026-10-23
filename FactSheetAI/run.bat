@echo off
cd /d "%~dp0"
echo ========================================================
echo Starting FactSheetAI Server...
echo Open your browser at: http://127.0.0.1:8000/
echo Dashboard is at:      http://127.0.0.1:8000/dashboard
echo ========================================================
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" -m uvicorn factsheetai.api:app --host 127.0.0.1 --port 8000 --reload
) else if exist "..\.venv\Scripts\python.exe" (
    "..\.venv\Scripts\python.exe" -m uvicorn factsheetai.api:app --host 127.0.0.1 --port 8000 --reload
) else (
    python -m uvicorn factsheetai.api:app --host 127.0.0.1 --port 8000 --reload
)
pause
