$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $projectRoot

uvx `
    --from pyinstaller `
    --with pypdf `
    --with pyside6 `
    --with pywin32 `
    pyinstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --name pdfconvert `
    --paths src `
    --distpath dist `
    --workpath build/pyinstaller `
    --specpath build `
    src/pdfconvert/app.py

if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller build failed with exit code $LASTEXITCODE."
}

Write-Host "Build complete: $projectRoot\dist\pdfconvert.exe"
