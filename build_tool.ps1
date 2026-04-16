param(
    [string]$PythonExe = "python"
)

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

Write-Host "==> 安装/升级打包依赖 pyinstaller"
& $PythonExe -m pip install --upgrade pyinstaller

Write-Host "==> 清理旧构建"
if (Test-Path ".\build") { Remove-Item ".\build" -Recurse -Force }
if (Test-Path ".\dist") { Remove-Item ".\dist" -Recurse -Force }
if (Test-Path ".\word_to_excel.spec") { Remove-Item ".\word_to_excel.spec" -Force }

Write-Host "==> 生成 EXE"
& $PythonExe -m PyInstaller `
    --noconfirm `
    --onefile `
    --name word_to_excel `
    .\word_to_questionbank_excel.py

Write-Host "==> 打包完成: .\dist\word_to_excel.exe"
