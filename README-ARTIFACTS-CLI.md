# Guia de Artefactos CLI (Descargar y Usar)

Este proyecto genera artefactos CLI cuando se ejecuta la integración continua de GitHub Actions. Los artefactos permiten distribuir y ejecutar la versión empaquetada de la aplicación en Windows sin necesitar Node.js en la máquina de destino.

## ¿Qué artefactos existen?
- recibos-sueldo-exe: ejecutable Windows (.exe) empaquetado con la CLI.
- distrib-recibos-sueldo-zip: ZIP de distribución que contiene dist/recibos-sueldo.exe y run-example.bat. Útil para distribución offline o USB.

## ¿Dónde los encuentro?
- En GitHub, abre tu repositorio y ve a la pestaña Actions.
- Selecciona la ejecución más reciente de la rama main (o la que corresponda).
- En la sección Artifacts, encontrarás:
  - recibos-sueldo-exe
  - distrib-recibos-sueldo-zip
- Descarga el artefacto que necesites.

## Cómo usar los artefactos
1) Descargar y preparar
- Descomprime (si corresponde) el ZIP distrib-recibos-sueldo-zip en una carpeta.
- Si usas el ejecutable directamente, localiza dist/recibos-sueldo.exe dentro del artefacto descargado.

2) Ejecución con un Excel de liquidación
- Usando el ejecutable (.exe):
```
```
- Usando el BAT de ejemplo (run-example.bat) incluido en el ZIP:
```
run-example.bat "C:\\ruta\\LiquidaciondeSueldosAbril12.xlsx" "C:\\ruta\\salida" "Abril 2026"
```

Notas:
- Las rutas entre comillas deben incluir espacios si los tuvieran.
- El ejecutable está empaquetado para Windows 64-bit (node16-win-x64). Si tu PC es 32-bit, no funcionará con este build.

3) Verificación de salida
- Recibos HTML por empleado: recibo_<legajo>_<nombre>.html
- Resumen CSV: resumen_recibos.csv
- Si descargas el ZIP, también puedes usar run-example.bat para validar rápido.

## Seguridad y mantenimiento
- Descarga los artefactos desde la fuente oficial del repositorio/CI para evitar software malicioso.
- Verifica que el Excel de liquidación siga la estructura esperada (hojas Empleador, Empleados, Remuneración, Ret. y no Rem.).
- Conserva backups de tus datos antes de generar recibos.

## Soporte
- Si tienes problemas para descargar o usar los artefactos, dime qué ejecución/errores ves y te guío paso a paso.

## Notas finales
- Este README describe el flujo de descarga y uso para el artefacto CLI. No describe cambios futuros; si necesitas soporte para otras plataformas o características adicionales, avísame y lo integramos.

## Capturas (próximas mejoras)
- capturas de ejemplo pueden incluirse aquí en futuras versiones.

## Archivos relacionados
- Ejecutables y ZIP disponibles en la CI: ver README-ARTIFACTS-CLI.md en el repositorio para más detalle.
