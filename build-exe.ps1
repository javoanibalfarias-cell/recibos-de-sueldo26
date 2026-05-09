<# Windows PowerShell script to build an executable using pkg #>
Param()
Set-StrictMode -Version Latest

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Push-Location $root
try {
  if (-not (Test-Path "package.json")) { Write-Error "package.json no encontrado en $root"; exit 1 }
  # Instalar pkg localmente si no está
  if (-not (Test-Path "./node_modules/.bin/pkg")) {
    Write-Host "Instalando pkg localmente..."
    npm install --save-dev pkg@^5
  }
  # Construir ejecutable para Windows x64
  Write-Host "Construyendo ejecutable Windows..."
  if (Test-Path "./dist") { Remove-Item -Recurse -Force ./dist }
  npx pkg index.js --targets node16-win-x64 --output ./dist/recibos-sueldo.exe
  if (Test-Path "./dist/recibos-sueldo.exe") {
    Write-Host "Ejecutable generado: dist/recibos-sueldo.exe"
  } else {
    Write-Error "La compilación no generó el ejecutable. Verifica los logs de pkg."
  }
} catch {
  Write-Error $_.Exception.Message
} finally {
  Pop-Location
}
