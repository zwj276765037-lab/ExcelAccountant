param(
    [switch]$SkipTests
)

$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
$pythonExe = Join-Path $projectRoot '.venv\Scripts\python.exe'

if (-not (Test-Path -LiteralPath $pythonExe)) {
    throw "Python virtual environment not found: $pythonExe"
}

Push-Location $projectRoot
try {
    if (-not $SkipTests) {
        & $pythonExe -m pytest -q
        if ($LASTEXITCODE -ne 0) {
            throw "Tests failed with exit code $LASTEXITCODE"
        }
    }

    $commonArguments = @(
        '--noconfirm',
        '--clean',
        '--windowed',
        '--paths', 'src',
        '--collect-all', 'ortools'
    )

    & $pythonExe -m PyInstaller @commonArguments --onedir --name ExcelAccountant-folder excel_accountant_app.py
    if ($LASTEXITCODE -ne 0) {
        throw "One-folder build failed with exit code $LASTEXITCODE"
    }

    $folderExecutable = Join-Path $projectRoot 'dist\ExcelAccountant-folder\ExcelAccountant-folder.exe'
    & $folderExecutable --version
    if ($LASTEXITCODE -ne 0) {
        throw "One-folder smoke test failed with exit code $LASTEXITCODE"
    }

    & $pythonExe -m PyInstaller @commonArguments --onefile --name ExcelAccountant excel_accountant_app.py
    if ($LASTEXITCODE -ne 0) {
        throw "One-file build failed with exit code $LASTEXITCODE"
    }

    $oneFileExecutable = Join-Path $projectRoot 'dist\ExcelAccountant.exe'
    & $oneFileExecutable --version
    if ($LASTEXITCODE -ne 0) {
        throw "One-file smoke test failed with exit code $LASTEXITCODE"
    }

    Write-Host "Build completed:"
    Write-Host "  $folderExecutable"
    Write-Host "  $oneFileExecutable"
} finally {
    Pop-Location
}
