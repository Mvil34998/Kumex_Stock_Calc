# Simple one-command build for a standalone Windows EXE (no Python required on target PC)
# Usage (from repo root): powershell -ExecutionPolicy Bypass -File packaging/build.ps1

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path | Split-Path -Parent
$venvPython = Join-Path $repoRoot '.venv\Scripts\python.exe'

if (-not (Test-Path $venvPython)) {
    Write-Host "Creating venv in .venv..."
    & python -m venv (Join-Path $repoRoot '.venv')
}

Write-Host "Ensuring pip is up to date..."
& $venvPython -m pip install --upgrade pip

Write-Host "Installing app deps + PyInstaller..."
& $venvPython -m pip install -r (Join-Path $repoRoot 'requirements.txt') pyinstaller

$distPath = Join-Path $repoRoot 'dist'
$buildPath = Join-Path $repoRoot 'build'
$entry = Join-Path $repoRoot 'src\kumex\Kumex Ladu.py'

$null = New-Item -ItemType Directory -Force -Path $distPath | Out-Null
$null = New-Item -ItemType Directory -Force -Path $buildPath | Out-Null

Write-Host "Cleaning previous build artifacts..."
Remove-Item -Recurse -Force (Join-Path $distPath '*') -ErrorAction SilentlyContinue
Remove-Item -Recurse -Force (Join-Path $buildPath '*') -ErrorAction SilentlyContinue

Write-Host "Building EXE..."
& $venvPython -m PyInstaller `
    --noconfirm `
    --clean `
    --noconsole `
    --onefile `
    --name "Kumex" `
    --distpath $distPath `
    --workpath $buildPath `
    $entry

Write-Host "`nDone. Pick up the installer at: $distPath"
