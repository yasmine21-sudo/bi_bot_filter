$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$pythonExe = Join-Path $repoRoot ".venv\Scripts\python.exe"
$mainScript = Join-Path $repoRoot "main.py"

if (-not (Test-Path $pythonExe)) {
    throw "Virtual environment Python was not found at $pythonExe. Create .venv first."
}

if (-not (Test-Path $mainScript)) {
    throw "Main script was not found at $mainScript."
}

Push-Location $repoRoot
try {
    & $pythonExe $mainScript
    exit $LASTEXITCODE
}
finally {
    Pop-Location
}
