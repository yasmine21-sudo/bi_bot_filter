$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$pythonExe = Join-Path $repoRoot ".venv\Scripts\python.exe"
$scriptPath = Join-Path $repoRoot "send_test_email.py"

if (-not (Test-Path $pythonExe)) {
    throw "Virtual environment Python was not found at $pythonExe."
}

Push-Location $repoRoot
try {
    & $pythonExe $scriptPath
    exit $LASTEXITCODE
}
finally {
    Pop-Location
}
