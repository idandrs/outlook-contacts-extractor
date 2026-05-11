$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$VenvPath = Join-Path $ProjectRoot ".venv"

if (Get-Command py -ErrorAction SilentlyContinue) {
    & py -3 -m venv $VenvPath
} elseif (Get-Command python -ErrorAction SilentlyContinue) {
    & python -m venv $VenvPath
} else {
    throw "Python was not found. Install Python for Windows and retry."
}

$PythonExe = Join-Path $VenvPath "Scripts\python.exe"
$PipExe = Join-Path $VenvPath "Scripts\pip.exe"

& $PythonExe -m pip install --upgrade pip
& $PipExe install -r (Join-Path $ProjectRoot "requirements.txt")

Write-Host "Bootstrap complete."
Write-Host "Next:"
Write-Host "  Set-Location `"$ProjectRoot`""
Write-Host "  .\.venv\Scripts\python.exe .\scripts\export_contacts.py"
