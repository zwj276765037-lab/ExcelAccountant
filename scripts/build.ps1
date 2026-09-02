$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
$pythonExe = Join-Path $projectRoot '.venv\Scripts\python.exe'

if (-not (Test-Path -LiteralPath $pythonExe)) {
    throw "Python virtual environment not found: $pythonExe"
}

Push-Location $projectRoot
try {
    & $pythonExe -m pytest
    & $pythonExe -m PyInstaller --noconfirm --clean --windowed --name ExcelAccountant --paths src src\excel_accountant\__main__.py
} finally {
    Pop-Location
}
