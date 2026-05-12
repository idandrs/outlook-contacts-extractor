$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$PythonExe = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
$ExportScript = Join-Path $ProjectRoot "scripts\export_addresses_from_mail.py"

if (-not (Test-Path $PythonExe)) {
    throw "The virtual environment was not found. Run .\scripts\bootstrap.ps1 first."
}

Push-Location $ProjectRoot
try {
    & $PythonExe $ExportScript
} finally {
    Pop-Location
}
