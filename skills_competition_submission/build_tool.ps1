param(
    [string]$PythonExe = "python"
)

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

Write-Host "==> Ensuring PyInstaller is installed"
& $PythonExe -m pip install --upgrade pyinstaller

Write-Host "==> Cleaning previous build artifacts"
if (Test-Path ".\build") { Remove-Item ".\build" -Recurse -Force }
if (Test-Path ".\dist") { Remove-Item ".\dist" -Recurse -Force }
if (Test-Path ".\word_to_excel.spec") { Remove-Item ".\word_to_excel.spec" -Force }

Write-Host "==> Building EXE"
& $PythonExe -m PyInstaller `
    --noconfirm `
    --onefile `
    --name word_to_excel `
    .\word_to_questionbank_excel.py

Write-Host "==> Build finished: .\dist\word_to_excel.exe"
