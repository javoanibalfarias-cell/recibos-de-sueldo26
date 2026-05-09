const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

function fmt(n){ try { return new Intl.NumberFormat('es-AR',{style:'currency',currency:'ARS'}).format(Number(n||0)); } catch { return String(n||0); } }
function num(v){ if(v==null||v===''||String(v).startsWith('#')) return 0; if(typeof v==='number') return v; const s=String(v).replace(/\./g,'').replace(',','.').replace(/[^0-9.-]/g,''); const x=parseFloat(s); return isNaN(x)?0:x; }
function text(v){ return v==null? '': String(v).trim(); }
function maxRow(ws){ return ws && ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']).e.r + 1 : 0; }
function cell(ws, addr){ return ws && ws[addr] ? ws[addr].v : ''; }
function rowCell(ws, col, row){ return cell(ws, String(col) + String(row)); }

function loadPayrollFromExcel(excelPath, periodo){
  const wb = XLSX.readFile(excelPath, {type:'file', cellDates:true});
  const getSheet = name => wb.Sheets[name];
  const empleador = {
    razonSocial: text(cell(getSheet('Empleador'), 'D6')), cuit: text(cell(getSheet('Empleador'), 'D7')),
    domicilio: text(cell(getSheet('Empleador'), 'D8')), localidad: text(cell(getSheet('Empleador'), 'D9')),
    provincia: text(cell(getSheet('Empleador'), 'D10')), cp: text(cell(getSheet('Empleador'), 'D11')),
    pago: text(cell(getSheet('Empleador'), 'D12'))
  };

  const empWS = getSheet('Empleados');
  const remWS = getSheet('Remuneración');
  const retWS = getSheet('Ret. y no Rem.');
  if(!empWS || !remWS || !retWS){ throw new Error('Estructura de Excel no coincide con lo esperado.'); }

  // Leer retenciones por legajo
  const retByLeg = {};
  const retLast = maxRow(retWS);
  for(let r=7; r<=retLast; r++){ const leg = text(rowCell(retWS,'B',r)); if(!leg) continue; retByLeg[leg] = readRetenciones(retWS, r); }

  // Leer remuneraciones por legajo
  const remByLeg = {};
  const remLast = maxRow(remWS);
  for(let r=8; r<=remLast; r++){ const leg = text(rowCell(remWS,'B',r)); if(!leg) continue; remByLeg[leg] = readRemuneraciones(remWS, r); }

  const empleados = [];
  const last = maxRow(empWS);
  for(let r=8; r<=last; r++){
    const leg = text(rowCell(empWS,'C',r)); const nombre = text(rowCell(empWS,'E',r));
    if(!leg || !nombre) continue;
    const remuneracion = remByLeg[leg] || readRemuneraciones(remWS, r);
    const retencion = retByLeg[leg] || readRetenciones(retWS, r-1);
    const bruto = remuneracion.remunerativos.reduce((a,c)=>a+c.importe,0);
    const noRem = remuneracion.noRemunerativos.reduce((a,c)=>a+c.importe,0) + retencion.noRemunerativos.reduce((a,c)=>a+c.importe,0);
    const desc = retencion.descuentos.reduce((a,c)=>a+c.importe,0);
    const neto = retencion.neto || Math.round(bruto + noRem - desc);
    empleados.push({
      legajo: leg, ingreso: rowCell(empWS,'D',r), nombre, cuil: text(rowCell(empWS,'F',r)),
      sueldoBasico: num(rowCell(empWS,'G',r)), cuenta: text(rowCell(empWS,'M',r)),
      calificacion: text(rowCell(empWS,'N',r)), tarea: text(rowCell(empWS,'O',r)), afjp: text(rowCell(empWS,'P',r)),
      remunerativos: remuneracion.remunerativos, noRemunerativos: [...remuneracion.noRemunerativos, ...retencion.noRemunerativos], descuentos: retencion.descuentos,
      bruto, totalNoRem: noRem, totalDescuentos: desc, neto
    });
  }

  return { empleador, empleados, periodo: periodo || 'Period' };
}

function readRemuneraciones(ws, r){
  const conceptos = [];
  const add = (col, label, importe, qtyCol) => { const amount = num(importe); if(amount>0) conceptos.push({codigo:col, concepto:label, cantidad: qtyCol ? text(rowCell(ws, qtyCol, r)) : '', importe: amount}); };
  add('E','Sueldo básico / sueldo exacto', rowCell(ws,'E',r));
  add('G','Básico días normales', rowCell(ws,'G',r),'F'); add('I','Ausentismo', rowCell(ws,'I',r),'H'); add('K','Días trabajados / ajuste', rowCell(ws,'K',r),'J');
  add('M','Presentismo', rowCell(ws,'M',r),'L'); add('O','Antigüedad', rowCell(ws,'O',r),'N'); add('Q','Adicional remunerativo', rowCell(ws,'Q',r),'P');
  add('S','Adicional fijo / porcentaje', rowCell(ws,'S',r),'R'); add('U','Falla de caja', rowCell(ws,'U',r),'T'); add('W','Adicional remunerativo 2',rowCell(ws,'W',r),'V'); add('Y','Adicional remunerativo 3',rowCell(ws,'Y',r),'X');
  [ 'Z','AA','AB','AC','AD','AE' ].forEach(c => add(c, text(rowCell(ws,c,r)) || c, rowCell(ws,c,r)));
  add('AH','Horas extras 50%', rowCell(ws,'AH',r),'AF'); add('AK','Horas extras 100%', rowCell(ws,'AK',r),'AI'); add('AL','SAC proporcional', rowCell(ws,'AL',r)); add('AN','Vacaciones', rowCell(ws,'AN',r),'AM');
  const brutoCached = num(rowCell(ws,'AP',r));
  return { remunerativos: conceptos.filter(c=>c.codigo!=='Z'&&c.codigo!=='AA'&&c.codigo!=='AB'&&c.codigo!=='AC'&&c.codigo!=='AD'&&c.codigo!=='AE'), noRemunerativos: conceptos.filter(c=>['Z','AA','AB','AC','AD','AE'].includes(c.codigo)), brutoCached };
}

function readRetenciones(ws, r){
  const descuentos = []; const noRem = [];
  const addD = (col, label, amount, pctCol) => { const n = num(amount); if(n>0) descuentos.push({codigo:col, concepto:label, porcentaje:pctCol?text(rowCell(ws,pctCol,r)):'', importe:n}); };
  const addN = (col, label, amount) => { const n = num(amount); if(n>0) noRem.push({codigo:col, concepto:label, importe:n}); };
  addD('E','Jubilación', rowCell(ws,'E',r),'D'); addD('G','Ley 19032 / INSSJP', rowCell(ws,'G',r),'F'); addD('I','Obra social', rowCell(ws,'I',r),'H');
  addD('K','Aporte sindical / retención 1', rowCell(ws,'K',r),'J'); addD('M','Retención 2', rowCell(ws,'M',r),'L'); addD('O','Retención 3', rowCell(ws,'O',r),'N');
  addD('Q','Retención 4', rowCell(ws,'Q',r),'P'); addD('S','Seguro / FAECYS', rowCell(ws,'S',r),'R'); addD('U','Retención 5', rowCell(ws,'U',r),'T'); addD('W','Retención 6', rowCell(ws,'W',r),'V'); addD('Y','Retención 7', rowCell(ws,'Y',r),'X');
  addD('Z','Anticipos', rowCell(ws,'Z',r)); addD('AA','Anticipo SAC', rowCell(ws,'AA',r)); addD('Obra Social', rowCell(ws,'AB',r)); // fallback
  ['AD','AE','AF','AG','AH','AI','AJ','AK'].forEach(c=> addN(c, getHeaderFromRow(ws,c,6,c), rowCell(ws,c,r)));
  const subtotal = num(rowCell(ws,'AC',r)); const netoExacto = num(rowCell(ws,'AN',r)); const neto = num(rowCell(ws,'AO',r)) || netoExacto;
  return { descuentos, noRemunerativos:noRem, subtotal, neto };
}

function getHeaderFromRow(ws,col,row,fallback){ return text(rowCell(ws,col,row)) || fallback; }

function buildReceiptHtml(e, empleador, periodo){
  // Genera un HTML simple y profesional para impresión
  const header = `Recibo de Haberes - Periodo: ${periodo}`;
  const habRows = [...e.remunerativos, ...e.noRemunerativos].map(h => `<tr><td>${escape(h.concepto||'')}</td><td>${escape(h.cantidad||'')}</td><td>${escape(h.tipo||'')}</td><td style="text-align:right;">${fmt(h.importe)}</td></tr>`).join('');
  const descRows = (e.descuentos||[]).map(d => `<tr><td>${escape(d.concepto||'')}</td><td>${escape(d.porcentaje||'')}</td><td style="text-align:right;">${fmt(d.importe||0)}</td></tr>`).join('');
  const totalHab = fmt(e.bruto);
  const totalDesc = fmt(e.totalDescuentos);
  const neto = fmt(e.neto);
  const html = `<!doctype html><html><head><meta charset="utf-8"><title>Recibo - ${escape(e.nombre)}</title><style>body{font-family:Arial,Helvetica,sans-serif; padding:20px;} h2{text-align:center; text-transform:uppercase; font-size:14px} .sep{border-top:2px solid #333; margin:8px 0} table{width:100%; border-collapse:collapse; font-size:12px} th,td{border:1px solid #999; padding:6px} th{background:#f0f0f0} .right{text-align:right}</style></head><body><h2>Recibo de Haberes</h2><p><strong>Empleador:</strong> ${escape(empleador.razonSocial||'')}&nbsp;&nbsp;<strong>Período:</strong> ${escape(periodo||'')}</p><div class="sep"></div><p><strong>Empleado:</strong> ${escape(e.nombre)} &nbsp;&nbsp; <strong>Legajo:</strong> ${escape(e.legajo)} &nbsp;&nbsp; <strong>CUIL:</strong> ${escape(e.cuil||'')}</p><div class="sep"></div><h3>Haberes</h3><table><thead><tr><th>Concepto</th><th>Cant./Base</th><th>Tipo</th><th>Importe</th></tr></thead><tbody>${habRows || '<tr><td colspan="4">Sin haberes detectados</td></tr>'}</tbody></table><div class="sep"></div><h3>Descuentos</h3><table><thead><tr><th>Descuento</th><th>%</th><th>Importe</th></tr></thead><tbody>${descRows || '<tr><td colspan="3">Sin descuentos detectados</td></tr>'}</tbody></table><div class="sep"></div><p><strong>Total Haberes:</strong> ${totalHab} &nbsp;&nbsp; <strong>Total Descuentos:</strong> ${totalDesc} &nbsp;&nbsp; <strong>Neto:</strong> ${neto}</p></body></html>`;
  return html;
}

function buildSummaryCsv(empleados){
  const header = ['Legajo','Empleado','CUIL','Bruto','No remunerativo','Descuentos','Neto'];
  const rows = empleados.map(e => [e.legajo, e.nombre, e.cuil, e.bruto, e.totalNoRem, e.totalDescuentos, e.neto]);
  const lines = [header].concat(rows).map(r => r.map(v => '"'+String(v||'').replace(/"/g,'""')+'"').join(';'));
  return lines.join('\r\n');
}

function escape(s){ return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

module.exports = {
  loadPayrollFromExcel, buildReceiptHtml, buildSummaryCsv
};
