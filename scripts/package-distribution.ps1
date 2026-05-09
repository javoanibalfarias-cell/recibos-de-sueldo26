<#
Script Name: package-distribution.ps1
Purpose: Create a distribution ZIP containing the Windows executable and a usage BAT.
Usage: Run from repo root; script will look for dist/recibos-sueldo.exe and create dist/recibos-sueldo.zip
#>

Param()

$exePath = Join-Path (Resolve-Path .).Path 'dist\recibos-sueldo.exe'
if (-not (Test-Path $exePath)) {
  Write-Error "Ejecutable no encontrado en $exePath. Construye primero con npm run build:exe"; exit 1
}

$batPath = Join-Path (Resolve-Path .).Path 'run-example.bat'
@("@echo off","setlocal","echo Running example: press to test the exe","") | Out-File -FilePath $batPath -Encoding ASCII -Force

$zipPath = Join-Path (Resolve-Path .).Path 'dist\recibos-sueldo.zip'
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
Compress-Archive -Path $exePath, $batPath -DestinationPath $zipPath -Force
Write-Host "[INFO] Distribución creada: $zipPath" -ForegroundColor Green
