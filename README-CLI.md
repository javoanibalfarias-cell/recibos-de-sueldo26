# CLI: Generador de Recibos de Sueldo
Esta carpeta contiene una aplicación local para generar recibos de sueldo a partir de un Excel completo de liquidación. Incluye una CLI en Node.js y, opcionalmente, un script para generar un ejecutable Windows (.exe) usando pkg.

Uso básico:
- Instalar dependencias: npm install
- Ejecutar la CLI para generar recibos HTML y un CSV resumen:
  node index.js --excel "ruta/al/LiquidaciondeSueldosAbril12.xlsx" --output "salida" --periodo "Abril 2026" --export-csv

Generar ejecutable Windows (.exe):
- Opción A: usar npm script si tienes pkg instalado globalmente:
  npm run build:exe
- Opción B: usar script de PowerShell incluido:
  ./build-exe.ps1
- Una vez generado, iniciar el ejecutable:
  dist/recibos-sueldo.exe --excel "ruta/al/LiquidaciondeSueldosAbril12.xlsx" --output "salida" --periodo "Abril 2026" --export-csv

Notas:
- El ejecutable contiene la lógica de generación de recibos y no requiere Node.js en la máquina de destino.
- Requiere Node.js para el build y for PKG durante el empaquetado.
