$ErrorActionPreference = "Stop"

if (-not (Test-Path ".venv")) {
    python -m venv .venv
}

& .\.venv\Scripts\python.exe -m pip install --upgrade pip
& .\.venv\Scripts\python.exe -m pip install -e ".[dev]"
Write-Host ""
Write-Host "Bootstrap concluido."
Write-Host "Ative com: .\.venv\Scripts\Activate.ps1"
Write-Host "Teste com: tales status"
