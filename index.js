#!/usr/bin/env node
// Generador de recibos de sueldo (CLI)
// Lee un Excel de liquidación y produce recibos HTML (y opcionalmente PDF)

const path = require('path');
const fs = require('fs');
const { hideBin } = require('yargs/helpers');
const yargs = require('yargs/yargs');

const payrollLib = require('./lib/payroll');

async function main(){
  const argv = yargs(hideBin(process.argv))
    .usage('Uso: $0 --excel <archivo.xlsx> [--output <dir>] [--periodo "Mes 2026"] [--format html|html-pdf] [--export-csv]')
    .option('excel', { type: 'string', describe: 'Ruta al archivo Excel de liquidación', demandOption: true })
    .option('output', { type: 'string', describe: 'Directorio de salida', default: './salida' })
    .option('periodo', { type: 'string', describe: 'Periodo a usar en los recibos', default: 'Mes actual' })
    .option('format', { type: 'string', choices: ['html','html-pdf','pdf'], describe: 'Formato de salida', default: 'html' })
    .option('export-csv', { type: 'boolean', describe: 'Exportar resumen CSV', default: true })
    .help()
    .argv;

  const excelPath = path.resolve(argv.excel);
  const outDir = path.resolve(argv.output);
  if(!fs.existsSync(excelPath)){
    console.error('No se encontró el archivo Excel:', excelPath);
    process.exit(1);
  }
  if(!fs.existsSync(outDir)){
    fs.mkdirSync(outDir, { recursive: true });
  }

  console.log('Cargando Excel:', excelPath);
  const data = await payrollLib.loadPayrollFromExcel(excelPath, argv.periodo);
  // Generar recibos HTML por empleado
  const receipts = [];
  data.empleados.forEach((e, idx) => {
    const html = payrollLib.buildReceiptHtml(e, data.empleador, data.periodo);
    const fileName = `recibo_${e.legajo}_${slugify(e.nombre)}.html`;
    const filePath = path.join(outDir, fileName);
    fs.writeFileSync(filePath, html, 'utf8');
    receipts.push({ empleado: e, path: filePath });
  });

  // Export CSV resumen si se solicita
  if(argv['export-csv']){
    const csvPath = path.join(outDir, 'resumen_recibos.csv');
    const csv = payrollLib.buildSummaryCsv(data.empleados);
    fs.writeFileSync(csvPath, csv, 'utf8');
    console.log('CSV resumen generado:', csvPath);
  }

  // Opcional: generar PDFs desde HTML si se desea, usando Puppeteer si está disponible
  if(argv.format === 'html-pdf' || argv.format === 'pdf'){
    // Intentar generar PDFs desde las HTML generadas
    try{
      const puppeteer = require('puppeteer');
      for(const r of receipts){
        const pdfPath = path.join(outDir, path.basename(r.path, '.html') + '.pdf');
        const browser = await puppeteer.launch({ headless: true, args: ['--no-sandbox'] });
        const page = await browser.newPage();
        await page.goto('file://' + r.path, { waitUntil: 'networkidle0' });
        await page.pdf({ path: pdfPath, format: 'A4' });
        await browser.close();
        console.log('PDF generado:', pdfPath);
      }
    }catch(err){
      console.warn('No se pudo generar PDFs ( Puppeteer no disponible ).', err?.message || err);
    }
  }

  console.log('Procesamiento finalizado. Salida en:', outDir);
}

function slugify(s){ return String(s||'').toLowerCase().replace(/[^a-z0-9]+/g,'_').replace(/^_|_$/g,''); }

main().catch(err => {
  console.error('Error:', err);
  process.exit(1);
});
