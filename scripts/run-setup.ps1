<#
Script Name: run-setup.ps1
Purpose: CLI-only automation for the app (no exe packaging).
Usage:
  powershell -ExecutionPolicy Bypass -File scripts/run-setup.ps1 -ExcelPath <path_to_excel.xlsx> [-OutputDir <output_dir>] [-Periodo "Abril 2026"]
This script assumes you are in the repository root and is intended to run on Windows.
#>

# Windows guard
if (-not $IsWindows) {
  Write-Error "Este script solo funciona en Windows."; exit 1
}

param(
  [Parameter(Mandatory=$true)][string]$ExcelPath,
  [Parameter(Mandatory=$false)][string]$OutputDir = "./salida",
  [Parameter(Mandatory=$false)][string]$Periodo
)

$ExcelPathSafe = (Resolve-Path $ExcelPath).Path
if (-not (Test-Path $ExcelPathSafe)) { Write-Error "Archivo Excel no encontrado: $ExcelPathSafe"; exit 1 }

$OutputDirSafe = (Resolve-Path $OutputDir -ErrorAction SilentlyContinue).Path
if (-not $OutputDirSafe) { $OutputDirSafe = (Join-Path (Get-Location).Path (Split-Path $OutputDir -Leaf)); if (-not (Test-Path $OutputDirSafe)) { New-Item -ItemType Directory -Path $OutputDirSafe | Out-Null } }

Write-Host "[INFO] CLI-only flow. Ejecutando dependencias y CLI." -ForegroundColor Cyan

if (-not (Get-Command node -ErrorAction SilentlyContinue)) { Write-Error "Node.js no está instalado."; exit 1 }
if (-not (Get-Command npm -ErrorAction SilentlyContinue)) { Write-Error "NPM no está instalado."; exit 1 }

Write-Host "[INFO] Instalando dependencias..." -ForegroundColor Cyan
npm ci 2>&1 | Tee-Object -Variable npmLog
if ($LASTEXITCODE -ne 0) { Write-Host $npmLog; Write-Error "Fallo al instalar dependencias."; exit 2 }

Write-Host "[INFO] Ejecutando pruebas de humo..." -ForegroundColor Cyan
npm test 2>&1 | Tee-Object -Variable testLog
if ($LASTEXITCODE -ne 0) { Write-Warning "Las pruebas de humo reportaron errores." }

 $cmd = "node index.js --excel `"$ExcelPathSafe`" --output `"$OutputDirSafe`" --export-csv"
if ($Periodo) { $cmd = $cmd + " --periodo `"$Periodo`"" }
Write-Host "[INFO] Ejecutando CLI: $cmd" -ForegroundColor Cyan
Invoke-Expression $cmd

Write-Host "[INFO] Flujo CLI completado." -ForegroundColor Green
