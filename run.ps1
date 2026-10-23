Set-Location -Path $PSScriptRoot
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "Starting FactSheetAI Server..." -ForegroundColor Green
Write-Host "Open your browser at: http://127.0.0.1:8000/" -ForegroundColor Yellow
Write-Host "Dashboard is at:      http://127.0.0.1:8000/dashboard" -ForegroundColor Yellow
Write-Host "========================================================" -ForegroundColor Cyan

if (Test-Path ".\.venv\Scripts\python.exe") {
    & ".\.venv\Scripts\python.exe" -m uvicorn factsheetai.api:app --host 127.0.0.1 --port 8000 --reload
} elseif (Test-Path ".\FactSheetAI\.venv\Scripts\python.exe") {
    Set-Location -Path ".\FactSheetAI"
    & ".\.venv\Scripts\python.exe" -m uvicorn factsheetai.api:app --host 127.0.0.1 --port 8000 --reload
} else {
    python -m uvicorn factsheetai.api:app --host 127.0.0.1 --port 8000 --reload
}
