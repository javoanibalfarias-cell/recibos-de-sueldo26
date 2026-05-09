@echo off
REM Example launcher for recibos-sueldo.exe
REM Usage: run with Excel path, Output dir, and Periodo optional
set EXCEL=%~1
set OUT=%~2
set PER=%~3

if "%EXCEL%"=="" (
  echo Usage: run-example.bat "path_to_excel.xlsx" "output_dir" "Periodo"
  goto :eof
)
dist\recibos-sueldo.exe --excel "%EXCEL%" --output "%OUT%" --periodo "%PER%" --export-csv
