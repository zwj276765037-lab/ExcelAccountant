param(
    [switch]$SkipTests
)

$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
$pythonExe = Join-Path $projectRoot '.venv\Scripts\python.exe'
$originalPath = $env:PATH

if (-not (Test-Path -LiteralPath $pythonExe)) {
    throw "Python virtual environment not found: $pythonExe"
}

function Test-PackagedSelfTest {
    param([Parameter(Mandatory = $true)][string]$Executable)

    $process = Start-Process -FilePath $Executable -ArgumentList '--self-test' -Wait -PassThru -WindowStyle Hidden
    if ($process.ExitCode -ne 0) {
        throw "Packaged self-test failed with exit code $($process.ExitCode): $Executable"
    }
}

function Test-PackagedWindow {
    param([Parameter(Mandatory = $true)][string]$Executable)

    $process = Start-Process -FilePath $Executable -PassThru
    $observedProcesses = @($process.Id)
    try {
        for ($attempt = 0; $attempt -lt 40; $attempt++) {
            Start-Sleep -Milliseconds 500
            $candidates = Get-CimInstance Win32_Process | Where-Object {
                $_.ProcessId -eq $process.Id -or $_.ParentProcessId -eq $process.Id
            }
            $observedProcesses = @($candidates.ProcessId)
            foreach ($candidate in $candidates) {
                $guiProcess = Get-Process -Id $candidate.ProcessId -ErrorAction SilentlyContinue
                if ($null -eq $guiProcess) { continue }
                if ($guiProcess.MainWindowTitle -eq 'Unhandled exception in script') {
                    throw "Packaged GUI raised an unhandled exception: $Executable"
                }
                if ($guiProcess.MainWindowHandle -ne 0 -and
                    $guiProcess.MainWindowTitle -like 'ExcelAccountant*') {
                    return
                }
            }
            if ($process.HasExited) {
                throw "Packaged GUI exited before creating a window: $Executable"
            }
        }
        throw "Packaged GUI did not create the ExcelAccountant window: $Executable"
    } finally {
        foreach ($processId in $observedProcesses) {
            $running = Get-Process -Id $processId -ErrorAction SilentlyContinue
            if ($null -ne $running) {
                $null = $running.CloseMainWindow()
                if (-not $running.WaitForExit(3000)) {
                    Stop-Process -Id $running.Id -Force
                }
            }
        }
    }
}

Push-Location $projectRoot
try {
    # The Codex workspace runtime can place Poppler's private ICU DLLs on PATH.
    # QtCore imports the Windows ICU forwarder with the same filename, so letting
    # PyInstaller collect Poppler's copy creates a package that starts but cannot
    # import QtCore. Exclude only Poppler's bin directory during dependency scan.
    $env:PATH = (($env:PATH -split [IO.Path]::PathSeparator) | Where-Object {
        $_ -notmatch '(?i)[\\/]poppler[\\/]Library[\\/]bin$'
    }) -join [IO.Path]::PathSeparator

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
        '--collect-binaries', 'ortools'
    )

    & $pythonExe -m PyInstaller @commonArguments --onedir --name ExcelAccountant-folder excel_accountant_app.py
    if ($LASTEXITCODE -ne 0) {
        throw "One-folder build failed with exit code $LASTEXITCODE"
    }

    $folderExecutable = Join-Path $projectRoot 'dist\ExcelAccountant-folder\ExcelAccountant-folder.exe'
    Test-PackagedSelfTest $folderExecutable
    Test-PackagedWindow $folderExecutable

    & $pythonExe -m PyInstaller @commonArguments --onefile --name ExcelAccountant excel_accountant_app.py
    if ($LASTEXITCODE -ne 0) {
        throw "One-file build failed with exit code $LASTEXITCODE"
    }

    $oneFileExecutable = Join-Path $projectRoot 'dist\ExcelAccountant.exe'
    Test-PackagedSelfTest $oneFileExecutable
    Test-PackagedWindow $oneFileExecutable

    Write-Host "Build completed:"
    Write-Host "  $folderExecutable"
    Write-Host "  $oneFileExecutable"
} finally {
    $env:PATH = $originalPath
    Pop-Location
}
