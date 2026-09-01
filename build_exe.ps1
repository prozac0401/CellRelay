param(
    [switch]$SkipInstall
)

$ErrorActionPreference = "Stop"
$projectRoot = $PSScriptRoot
$pythonPath = Join-Path $projectRoot ".venv312\Scripts\python.exe"
$distDir = Join-Path $projectRoot "dist"
$executablePath = Join-Path $distDir "CellRelay.exe"

if (-not (Test-Path -LiteralPath $pythonPath -PathType Leaf)) {
    throw "Python 3.12 build environment not found: $pythonPath"
}

& $pythonPath -m pip --version *> $null
if ($LASTEXITCODE -ne 0) {
    & $pythonPath -m ensurepip --upgrade
    if ($LASTEXITCODE -ne 0) {
        throw "Could not initialize pip in the build environment."
    }
}

if (-not $SkipInstall) {
    & $pythonPath -m pip install --disable-pip-version-check -r `
        (Join-Path $projectRoot "requirements-build.txt")
    if ($LASTEXITCODE -ne 0) {
        throw "Could not install build dependencies."
    }
}

Push-Location $projectRoot
try {
    & $pythonPath -m PyInstaller `
        --noconfirm `
        --clean `
        --windowed `
        --onefile `
        --name CellRelay `
        --collect-all playwright `
        --distpath $distDir `
        --workpath (Join-Path $projectRoot "build\pyinstaller") `
        --specpath (Join-Path $projectRoot "build") `
        (Join-Path $projectRoot "main.py")
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller build failed."
    }

    Write-Host "Standalone executable: $executablePath"
}
finally {
    Pop-Location
}
