@echo off
setlocal
cd /d "%~dp0"
REM Si se ha generado el ejecutable, ejecútalo; de lo contrario abre el HTML de inicio
IF EXIST "%~dp0dist\recibos-sueldo.exe" (
  start "" "%~dp0dist\recibos-sueldo.exe"
) ELSE (
  start "" "%~dp0AplicacionRecibos.html"
)
endlocal
