/* ══════════════════════════════════════════════════════
   FinDash — script.js
   Soporta dos modos de lectura:
     A) Multi-hoja: hojas "INGRESOS" y "EGRESOS" separadas
        (formato SMTO: columnas específicas por hoja)
     B) Hoja única: columnas Fecha, Tipo, Categoría, Monto
   El año se detecta automáticamente desde los datos.
══════════════════════════════════════════════════════ */

'use strict';

// ─── ESTADO GLOBAL ────────────────────────────────────
let allRows        = [];  // Todos los registros del año detectado
let filteredRows   = [];  // Registros filtrados (por mes)
let barChartInst   = null;
let donutChartInst = null;
let tipoBarInst    = null;
let detectedYear   = new Date().getFullYear();

// ─── PALETA ──────────────────────────────────────────
const PALETTE = [
  '#6366f1','#8b5cf6','#ec4899','#f43f5e',
  '#f97316','#f59e0b','#10b981','#14b8a6',
  '#06b6d4','#3b82f6','#a3e635','#84cc16',
  '#e879f9','#fb7185','#fbbf24','#34d399',
];

const MONTHS_ES   = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
const MONTHS_LONG = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

// ══════════════════════════════════════════════════════
// 1. DRAG & DROP / FILE INPUT
// ══════════════════════════════════════════════════════

function triggerFileInput() { document.getElementById('fileInput').click(); }

function handleDragOver(e) {
  e.preventDefault();
  document.getElementById('dropZone').classList.add('drag-over');
}
function handleDragLeave(e) {
  e.preventDefault();
  document.getElementById('dropZone').classList.remove('drag-over');
}
function handleDrop(e) {
  e.preventDefault();
  document.getElementById('dropZone').classList.remove('drag-over');
  const file = e.dataTransfer.files[0];
  if (file) processFile(file);
}
function handleFileChange(e) {
  const file = e.target.files[0];
  if (file) processFile(file);
}

// ══════════════════════════════════════════════════════
// 2. PROCESAMIENTO DEL ARCHIVO
// ══════════════════════════════════════════════════════

function processFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  if (!['xlsx','xls'].includes(ext)) {
    showError('El archivo debe ser .xlsx o .xls. Por favor verifica el formato.');
    return;
  }
  hideError();
  showLoading();

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const wb   = XLSX.read(data, { type: 'array', cellDates: true });

      let rows = [];

      // ── MODO A: hojas INGRESOS + EGRESOS ──
      const hasIngresos = wb.SheetNames.some(n => n.toUpperCase().includes('INGRESO'));
      const hasEgresos  = wb.SheetNames.some(n => n.toUpperCase().includes('EGRESO'));

      if (hasIngresos && hasEgresos) {
        rows = parseMultiSheet(wb);
      } else {
        // ── MODO B: hoja única genérica ──
        const ws  = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, { defval: null, raw: false });
        if (!raw || raw.length === 0) throw new Error('La hoja de cálculo está vacía.');
        rows = normalizeSingleSheet(raw);
      }

      if (rows.length === 0) {
        throw new Error(
          `No se encontraron transacciones para el año ${detectedYear}. ` +
          'Verifica el archivo o los encabezados.'
        );
      }

      allRows = rows;
      buildDashboard(rows);
    } catch (err) {
      hideLoading();
      showError(err.message || 'Error al leer el archivo. Verifica que sea un Excel válido.');
    }
  };
  reader.onerror = () => { hideLoading(); showError('No se pudo leer el archivo. Intenta de nuevo.'); };
  reader.readAsArrayBuffer(file);
}

// ══════════════════════════════════════════════════════
// 3A. PARSER MULTI-HOJA (formato SMTO)
//     INGRESOS: RFC(0) NOMBRE(1) TIPO(2) FACTURA(3)
//               IMPORTE(4) IVA(5) TOTAL(6) FECHA_PAGO(7)
//     EGRESOS:  RFC(0) PROVEEDOR(1) TIPO(2) POLIZA(3)
//               FACTURA(4) FECHA_FAC(5) CONCEPTO(6)
//               IMPORTE(7) IVA(8) RET(9) TOTAL(10)
// ══════════════════════════════════════════════════════

function findSheetName(wb, keyword) {
  const names = wb.SheetNames;
  const kw    = keyword.toUpperCase();
  // 1) Coincidencia exacta (ej: 'INGRESOS' === 'INGRESOS')
  const exact = names.find(n => n.trim().toUpperCase() === kw || n.trim().toUpperCase() === kw + 'S');
  if (exact) return exact;
  // 2) Nombre más corto que contiene la palabra (evita pivot tables largas)
  const partials = names.filter(n => n.trim().toUpperCase().includes(kw));
  if (partials.length === 0) return null;
  return partials.sort((a,b) => a.length - b.length)[0];
}

function parseMultiSheet(wb) {
  // Preferir hojas con nombre exacto 'INGRESOS' / 'EGRESOS'
  // Si no existen, tomar la que más se parezca (nombre más corto)
  const nameIng = findSheetName(wb, 'INGRESO');
  const nameEgr = findSheetName(wb, 'EGRESO');

  if (!nameIng || !nameEgr) {
    throw new Error('No se encontraron hojas de INGRESOS y EGRESOS en el archivo.');
  }

  const rowsIng = parseSheetIngresos(wb.Sheets[nameIng]);
  const rowsEgr = parseSheetEgresos(wb.Sheets[nameEgr]);
  const combined = [...rowsIng, ...rowsEgr];

  if (combined.length === 0) return [];

  // Auto-detectar año más frecuente en los datos
  const yearCount = {};
  for (const r of combined) {
    yearCount[r.year] = (yearCount[r.year] || 0) + 1;
  }
  detectedYear = parseInt(Object.entries(yearCount).sort((a,b) => b[1]-a[1])[0][0], 10);

  return combined.filter(r => r.year === detectedYear);
}

/**
 * Encuentra el índice de la fila de encabezados buscando una keyword.
 * Retorna el índice 0-based dentro del array raw, o -1 si no se encuentra.
 */
function findHeaderRow(rawArrays, keywords, maxScan = 20) {
  for (let i = 0; i < Math.min(rawArrays.length, maxScan); i++) {
    const row = rawArrays[i];
    if (!row) continue;
    const rowStr = row.map(c => c != null ? String(c).toUpperCase() : '').join('|');
    if (keywords.some(kw => rowStr.includes(kw.toUpperCase()))) return i;
  }
  return -1;
}

/**
 * Busca una clave en un objeto (insensible a mayúsculas, espacios y tildes).
 */
function findKey(obj, hints) {
  const keys = Object.keys(obj);
  for (const hint of hints) {
    const h = hint.toUpperCase().trim();
    const k = keys.find(k => k.toUpperCase().trim() === h);
    if (k) return k;
  }
  // fallback: partial match
  for (const hint of hints) {
    const h = hint.toUpperCase().trim();
    const k = keys.find(k => k.toUpperCase().trim().includes(h));
    if (k) return k;
  }
  return null;
}

function parseSheetIngresos(ws) {
  // Paso 1: Obtener arrays crudos para localizar el encabezado
  const rawArrays = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  const headerIdx = findHeaderRow(rawArrays, ['NOMBRE DEL CLIENTE', 'RFC CLIENTE', 'FECHA DE PAGO']);
  if (headerIdx === -1) return [];

  // Paso 2: Releer desde la fila de encabezado usando nombres de columna
  const objects = XLSX.utils.sheet_to_json(ws, { range: headerIdx, defval: null });

  const rows = [];
  for (const obj of objects) {
    // Saltar filas completamente vacías
    if (Object.values(obj).every(v => v == null || v === '')) continue;

    // ── Fecha ──
    const fechaKey = findKey(obj, ['FECHA DE PAGO', 'FECHA PAGO', 'FECHA']);
    if (!fechaKey) continue;
    const fecha = parseDate(obj[fechaKey]);
    if (!fecha) continue;

    // ── Monto ──
    const montoKey = findKey(obj, ['TOTAL', 'IMPORTE DEL CHEQUE O TRANSFERENCIA ELECTRONICA', 'IMPORTE']);
    if (!montoKey) continue;
    const total = parseMonto(obj[montoKey]);
    if (isNaN(total) || total === 0) continue;

    // ── Tipo ──
    const tipoKey = findKey(obj, ['TIPO']);
    const tipo = (tipoKey && obj[tipoKey]) ? String(obj[tipoKey]).trim() : 'Sin tipo';

    // ── Cliente ──
    const nombreKey = findKey(obj, ['NOMBRE DEL CLIENTE', 'NOMBRE CLIENTE', 'NOMBRE']);
    const nombre = (nombreKey && obj[nombreKey]) ? String(obj[nombreKey]).trim() : '—';

    rows.push({
      fecha,
      year:          fecha.getFullYear(),
      mes:           fecha.getMonth(),
      tipo_registro: 'Ingreso',
      tipo,
      categoria:     tipo,
      subcategoria:  nombre,
      monto:         Math.abs(total),
    });
  }
  return rows;
}

function parseSheetEgresos(ws) {
  // Paso 1: Localizar encabezado
  const rawArrays = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  const headerIdx = findHeaderRow(rawArrays, ['PROVEEDOR', 'RFC PROVEEDOR', 'FECHA FAC']);
  if (headerIdx === -1) return [];

  // Paso 2: Releer con nombres de columna desde el encabezado
  const objects = XLSX.utils.sheet_to_json(ws, { range: headerIdx, defval: null });

  const rows = [];
  for (const obj of objects) {
    if (Object.values(obj).every(v => v == null || v === '')) continue;

    // ── Fecha ──
    const fechaKey = findKey(obj, ['FECHA FAC.', 'FECHA FAC', 'FECHA FACTURA', 'FECHA']);
    if (!fechaKey) continue;
    const fecha = parseDate(obj[fechaKey]);
    if (!fecha) continue;

    // ── Monto ──
    const montoKey = findKey(obj, ['TOTAL', 'IMPORTE DEL CHEQUE O TRANSFERENCIA', 'IMPORTE']);
    if (!montoKey) continue;
    const total = parseMonto(obj[montoKey]);
    if (isNaN(total) || total === 0) continue;

    // ── Tipo ──
    const tipoKey = findKey(obj, ['TIPO']);
    const tipo = (tipoKey && obj[tipoKey]) ? String(obj[tipoKey]).trim() : 'Sin tipo';

    // ── Proveedor ──
    const provKey = findKey(obj, ['PROVEEDOR']);
    const prov = (provKey && obj[provKey]) ? String(obj[provKey]).trim() : '—';

    rows.push({
      fecha,
      year:          fecha.getFullYear(),
      mes:           fecha.getMonth(),
      tipo_registro: 'Egreso',
      tipo,
      categoria:     tipo,
      subcategoria:  prov,
      monto:         Math.abs(total),
    });
  }
  return rows;
}

// ══════════════════════════════════════════════════════
// 3B. PARSER HOJA ÚNICA GENÉRICA
// ══════════════════════════════════════════════════════

const COL_MAP = {
  fecha:      ['fecha','date','fec','dia','day','periodo','fecha de pago','fecha fac'],
  tipo:       ['tipo','type','clase','movimiento','naturaleza'],
  categoria:  ['categoria','categoría','category','cat','rubro','concepto','descripcion'],
  subcategoria:['subcategoria','subcategoría','subcategory','subcat','detalle'],
  monto:      ['monto','amount','valor','importe','total','cantidad','sum'],
};

function findCol(headers, aliases) {
  const hLow = headers.map(h => String(h).toLowerCase().trim());
  for (const alias of aliases) {
    const idx = hLow.indexOf(alias);
    if (idx !== -1) return headers[idx];
  }
  return null;
}

function normalizeSingleSheet(raw) {
  const headers    = Object.keys(raw[0]);
  const colFecha   = findCol(headers, COL_MAP.fecha);
  const colTipo    = findCol(headers, COL_MAP.tipo);
  const colCat     = findCol(headers, COL_MAP.categoria);
  const colSubcat  = findCol(headers, COL_MAP.subcategoria);
  const colMonto   = findCol(headers, COL_MAP.monto);

  if (!colFecha) throw new Error('No se encontró columna de Fecha.');
  if (!colMonto) throw new Error('No se encontró columna de Monto.');

  const rows = [];
  for (const row of raw) {
    try {
      const fecha = parseDate(row[colFecha]);
      if (!fecha) continue;
      const monto = parseMonto(row[colMonto]);
      if (isNaN(monto) || monto === 0) continue;

      let tipoReg = 'Egreso';
      if (colTipo && row[colTipo]) {
        const rt = String(row[colTipo]).toLowerCase().trim();
        if (rt.includes('ingreso')||rt.includes('income')||rt.includes('entrada')||rt==='+'||rt.includes('crédito')) tipoReg = 'Ingreso';
        else if (rt.includes('egreso')||rt.includes('gasto')||rt.includes('expense')||rt.includes('salida')||rt==='-') tipoReg = 'Egreso';
        else tipoReg = monto >= 0 ? 'Ingreso' : 'Egreso';
      } else {
        tipoReg = monto >= 0 ? 'Ingreso' : 'Egreso';
      }

      const cat    = colCat    && row[colCat]    ? String(row[colCat]).trim()    : 'Sin categoría';
      const subcat = colSubcat && row[colSubcat] ? String(row[colSubcat]).trim() : '—';
      const tipoVal= colTipo   && row[colTipo]   ? String(row[colTipo]).trim()   : tipoReg;

      rows.push({
        fecha,
        year:          fecha.getFullYear(),
        mes:           fecha.getMonth(),
        tipo_registro: tipoReg,
        tipo:          capitalizar(tipoVal),
        categoria:     capitalizar(cat),
        subcategoria:  capitalizar(subcat),
        monto:         Math.abs(monto),
      });
    } catch (_) { /* fila con error: ignorar */ }
  }

  if (rows.length === 0) return [];

  // Auto-detectar año
  const yearCount = {};
  for (const r of rows) yearCount[r.year] = (yearCount[r.year] || 0) + 1;
  detectedYear = parseInt(Object.entries(yearCount).sort((a,b) => b[1]-a[1])[0][0], 10);
  return rows.filter(r => r.year === detectedYear);
}

// ══════════════════════════════════════════════════════
// 4. CONSTRUCCIÓN DEL DASHBOARD
// ══════════════════════════════════════════════════════

function buildDashboard(rows) {
  filteredRows = rows;
  renderKPIs(rows);
  renderBarChart(rows);
  renderDonutChart(rows);
  renderTipoCharts(rows);
  renderCategoryTable(rows);
  renderTxTable(rows);
  buildMonthFilter(rows);

  hideLoading();
  document.getElementById('uploadSection').classList.add('hidden');
  document.getElementById('dashboardSection').classList.remove('hidden');
  document.getElementById('dashSubtitle').textContent = `Año ${detectedYear}`;
  document.getElementById('monthFilterWrapper').classList.remove('hidden');
  document.getElementById('exportCsvBtn').classList.remove('hidden');
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ── KPIs ──
function renderKPIs(rows) {
  const ing = rows.filter(r => r.tipo_registro === 'Ingreso');
  const egr = rows.filter(r => r.tipo_registro === 'Egreso');
  const totalIng = ing.reduce((s,r) => s + r.monto, 0);
  const totalEgr = egr.reduce((s,r) => s + r.monto, 0);
  const balance  = totalIng - totalEgr;
  const tasa     = totalIng > 0 ? (balance / totalIng * 100) : 0;

  document.getElementById('kpiIncome').textContent        = formatMoney(totalIng);
  document.getElementById('kpiIncomeDetail').textContent  = `${ing.length} transacciones`;
  document.getElementById('kpiExpense').textContent       = formatMoney(totalEgr);
  document.getElementById('kpiExpenseDetail').textContent = `${egr.length} transacciones`;
  document.getElementById('kpiBalance').textContent       = formatMoney(balance);
  document.getElementById('kpiBalanceDetail').textContent = balance >= 0 ? '✓ Balance positivo' : '⚠ Balance negativo';
  document.getElementById('kpiBalance').style.color       = balance >= 0 ? 'var(--income)' : 'var(--expense)';
  document.getElementById('kpiRate').textContent          = `${tasa.toFixed(1)}%`;
}

// ── Barras: Ingresos vs Egresos por mes ──
function renderBarChart(rows) {
  const meses = Array.from({length:12}, ()=>({ing:0, egr:0}));
  const conDatos = new Set();
  for (const r of rows) {
    meses[r.mes][r.tipo_registro==='Ingreso'?'ing':'egr'] += r.monto;
    conDatos.add(r.mes);
  }
  const sorted   = [...conDatos].sort((a,b)=>a-b);
  const labels   = sorted.map(m => MONTHS_ES[m]);
  const ingData  = sorted.map(m => meses[m].ing);
  const egrData  = sorted.map(m => meses[m].egr);

  const isDark   = document.documentElement.getAttribute('data-theme') !== 'light';
  const grid     = isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.06)';
  const tick     = isDark ? '#64748b' : '#94a3b8';
  const ctx      = document.getElementById('barChart').getContext('2d');

  if (barChartInst) barChartInst.destroy();
  barChartInst = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label:'Ingresos', data:ingData, backgroundColor:'rgba(16,185,129,0.75)', borderRadius:6, borderSkipped:false },
        { label:'Egresos',  data:egrData, backgroundColor:'rgba(244,63,94,0.75)',  borderRadius:6, borderSkipped:false },
      ],
    },
    options: {
      responsive:true, maintainAspectRatio:false,
      interaction:{mode:'index',intersect:false},
      plugins:{
        legend:{display:false},
        tooltip: tooltipDefaults(isDark, v => formatMoney(v)),
      },
      scales:{
        x:{grid:{display:false}, ticks:{color:tick, font:{family:'Inter',size:12}}},
        y:{grid:{color:grid}, border:{display:false},
           ticks:{color:tick, font:{family:'Inter',size:12}, callback:v=>formatMoneyShort(v)}},
      },
    },
  });
}

// ── Dona: Egresos por tipo ──
function renderDonutChart(rows) {
  const egr      = rows.filter(r => r.tipo_registro === 'Egreso');
  const byTipo   = agrupar(egr, 'tipo');
  const sorted   = Object.entries(byTipo).sort((a,b)=>b[1]-a[1]).slice(0,10);
  const total    = egr.reduce((s,r)=>s+r.monto,0);

  document.getElementById('donutTotal').textContent = formatMoney(total);

  const isDark = document.documentElement.getAttribute('data-theme') !== 'light';
  const ctx    = document.getElementById('donutChart').getContext('2d');
  if (donutChartInst) donutChartInst.destroy();
  donutChartInst = new Chart(ctx, {
    type:'doughnut',
    data:{
      labels: sorted.map(([t])=>t),
      datasets:[{
        data:   sorted.map(([,v])=>v),
        backgroundColor: sorted.map((_,i)=>PALETTE[i%PALETTE.length]),
        borderColor: isDark?'#1e293b':'#fff', borderWidth:3, hoverOffset:6,
      }],
    },
    options:{
      responsive:true, maintainAspectRatio:false, cutout:'72%',
      plugins:{
        legend:{position:'bottom', labels:{color:isDark?'#94a3b8':'#475569', font:{family:'Inter',size:11}, padding:12, boxWidth:10, usePointStyle:true}},
        tooltip: tooltipDefaults(isDark, (v,ctx2)=>{
          const pct = total>0 ? (v/total*100).toFixed(1) : '0';
          return ` ${ctx2.label}: ${formatMoney(v)} (${pct}%)`;
        }, true),
      },
    },
  });
}

// ── Barras horizontales: Ingresos Y Egresos por tipo ──
function renderTipoCharts(rows) {
  renderTipoBar(
    rows.filter(r => r.tipo_registro === 'Ingreso'),
    'tipoIngChart',
    'rgba(16,185,129,0.8)',
    'tipoIngBadge'
  );
  renderTipoBar(
    rows.filter(r => r.tipo_registro === 'Egreso'),
    'tipoEgrChart',
    'rgba(244,63,94,0.8)',
    'tipoEgrBadge'
  );
}

function renderTipoBar(rows, canvasId, color, badgeId) {
  const byTipo = agrupar(rows, 'tipo');
  const sorted = Object.entries(byTipo).sort((a,b)=>b[1]-a[1]);
  const labels = sorted.map(([t])=>t);
  const data   = sorted.map(([,v])=>v);

  document.getElementById(badgeId).textContent = `${sorted.length} tipos`;

  // Altura dinámica: 36px por barra + 40px padding, mínimo 220px
  const dynamicHeight = Math.max(220, sorted.length * 36 + 40);
  const wrapper = document.getElementById(canvasId).parentElement;
  wrapper.style.height = dynamicHeight + 'px';

  const isDark = document.documentElement.getAttribute('data-theme') !== 'light';
  const grid   = isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.06)';
  const tick   = isDark ? '#64748b' : '#94a3b8';

  const existingChart = Chart.getChart(canvasId);
  if (existingChart) existingChart.destroy();

  const ctx = document.getElementById(canvasId).getContext('2d');
  new Chart(ctx, {
    type:'bar',
    data:{
      labels,
      datasets:[{
        data,
        backgroundColor: color,
        borderRadius:5,
        borderSkipped:false,
      }],
    },
    options:{
      indexAxis:'y',
      responsive:true, maintainAspectRatio:false,
      plugins:{
        legend:{display:false},
        tooltip: tooltipDefaults(isDark, v => formatMoney(v)),
      },
      scales:{
        x:{grid:{color:grid}, border:{display:false},
           ticks:{color:tick, font:{family:'Inter',size:11}, callback:v=>formatMoneyShort(v)}},
        y:{grid:{display:false}, ticks:{color:isDark?'#94a3b8':'#475569', font:{family:'Inter',size:11}}},
      },
    },
  });
}

// ── Tabla de categorías/tipos ──
function renderCategoryTable(rows) {
  const egr      = rows.filter(r => r.tipo_registro === 'Egreso');
  const ing      = rows.filter(r => r.tipo_registro === 'Ingreso');
  const totalEgr = egr.reduce((s,r)=>s+r.monto,0);

  const byCatE = agrupar(egr,'tipo');
  const byCatI = agrupar(ing,'tipo');
  const cntE   = contar(egr,'tipo');
  const cntI   = contar(ing,'tipo');

  const allCats = new Set([...Object.keys(byCatE), ...Object.keys(byCatI)]);
  const entries = [];
  for (const cat of allCats) {
    if (byCatE[cat]) entries.push({cat, tipo:'Egreso',  total:byCatE[cat], count:cntE[cat]||0});
    if (byCatI[cat]) entries.push({cat, tipo:'Ingreso', total:byCatI[cat], count:cntI[cat]||0});
  }
  entries.sort((a,b)=>b.total-a.total);

  const tbody = document.getElementById('categoryTableBody');
  tbody.innerHTML = '';
  for (const e of entries) {
    const pct = (e.tipo==='Egreso' && totalEgr>0) ? (e.total/totalEgr*100) : null;
    const tr  = document.createElement('tr');
    tr.innerHTML = `
      <td><strong style="color:var(--text-1)">${escHtml(e.cat)}</strong></td>
      <td><span class="type-badge ${e.tipo==='Ingreso'?'type-income':'type-expense'}">${e.tipo}</span></td>
      <td class="text-right" style="color:var(--text-1);font-weight:600">${formatMoney(e.total)}</td>
      <td class="text-right">
        ${pct!==null
          ? `<div class="progress-cell"><span>${pct.toFixed(1)}%</span>
             <div class="progress-bar"><div class="progress-fill" style="width:${Math.min(pct,100)}%"></div></div></div>`
          : '<span style="color:var(--text-3)">—</span>'}
      </td>
      <td class="text-right">${e.count}</td>`;
    tbody.appendChild(tr);
  }
  document.getElementById('tableBadge').textContent = `${entries.length} categorías`;
}

// ── Tabla de transacciones ──
function renderTxTable(rows) {
  const sorted = [...rows].sort((a,b)=>b.fecha-a.fecha).slice(0,100);
  const tbody  = document.getElementById('txTableBody');
  tbody.innerHTML = '';
  for (const r of sorted) {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${formatDate(r.fecha)}</td>
      <td><span class="type-badge ${r.tipo_registro==='Ingreso'?'type-income':'type-expense'}">${r.tipo_registro}</span></td>
      <td><strong style="color:var(--text-1)">${escHtml(r.tipo)}</strong></td>
      <td style="color:var(--text-3)">${escHtml(r.subcategoria)}</td>
      <td class="text-right" style="color:${r.tipo_registro==='Ingreso'?'var(--income)':'var(--expense)'};font-weight:600">
        ${r.tipo_registro==='Ingreso'?'+':'-'}${formatMoney(r.monto)}
      </td>`;
    tbody.appendChild(tr);
  }
  document.getElementById('txBadge').textContent = `${sorted.length} registros`;
}

// ══════════════════════════════════════════════════════
// 5. FILTRO POR MES
// ══════════════════════════════════════════════════════

function buildMonthFilter(rows) {
  const meses = [...new Set(rows.map(r=>r.mes))].sort((a,b)=>a-b);
  const sel   = document.getElementById('monthFilter');
  while (sel.options.length > 1) sel.remove(1);
  for (const m of meses) {
    const opt = document.createElement('option');
    opt.value = m; opt.textContent = MONTHS_LONG[m];
    sel.appendChild(opt);
  }
  sel.onchange = () => applyMonthFilter(sel.value);
}

function applyMonthFilter(val) {
  filteredRows = val==='all' ? allRows : allRows.filter(r=>r.mes===parseInt(val,10));
  renderKPIs(filteredRows);
  renderBarChart(filteredRows);
  renderDonutChart(filteredRows);
  renderTipoCharts(filteredRows);
  renderCategoryTable(filteredRows);
  renderTxTable(filteredRows);
  const label = val==='all' ? `Año ${detectedYear}` : `${MONTHS_LONG[parseInt(val,10)]} ${detectedYear}`;
  document.getElementById('dashSubtitle').textContent = label;
}

// ══════════════════════════════════════════════════════
// 6. EXPORTAR CSV
// ══════════════════════════════════════════════════════

function exportCSV() {
  const header = ['Fecha','Tipo','Categoría/Tipo','Proveedor/Cliente','Monto'];
  const lines  = [header.join(',')];
  for (const r of filteredRows) {
    lines.push([
      formatDate(r.fecha), r.tipo_registro,
      `"${r.tipo}"`, `"${r.subcategoria}"`, r.monto.toFixed(2),
    ].join(','));
  }
  const blob = new Blob([lines.join('\n')], {type:'text/csv;charset=utf-8;'});
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href=url; a.download=`findash_${detectedYear}.csv`;
  document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
}

// ══════════════════════════════════════════════════════
// 7. TEMA DARK / LIGHT
// ══════════════════════════════════════════════════════

function toggleTheme() {
  const html   = document.documentElement;
  const isDark = html.getAttribute('data-theme')==='dark';
  html.setAttribute('data-theme', isDark?'light':'dark');
  document.getElementById('iconSun').classList.toggle('hidden',!isDark);
  document.getElementById('iconMoon').classList.toggle('hidden',isDark);
  if (allRows.length > 0) {
    renderBarChart(filteredRows);
    renderDonutChart(filteredRows);
    renderTipoCharts(filteredRows);
  }
}

// ══════════════════════════════════════════════════════
// 8. RESET
// ══════════════════════════════════════════════════════

function resetApp() {
  allRows=[]; filteredRows=[];
  document.getElementById('dashboardSection').classList.add('hidden');
  document.getElementById('monthFilterWrapper').classList.add('hidden');
  document.getElementById('exportCsvBtn').classList.add('hidden');
  document.getElementById('uploadSection').classList.remove('hidden');
  document.getElementById('monthFilter').value='all';
  document.getElementById('fileInput').value='';
  [barChartInst,donutChartInst].forEach(c=>{if(c){c.destroy();}});
  barChartInst=donutChartInst=null;
  // Destruir charts de tipo
  ['tipoIngChart','tipoEgrChart'].forEach(id=>{
    const c = Chart.getChart(id);
    if(c) c.destroy();
  });
  hideError();
  window.scrollTo({top:0,behavior:'smooth'});
}

// ══════════════════════════════════════════════════════
// 9. HELPERS UI
// ══════════════════════════════════════════════════════

function showLoading() {
  document.getElementById('uploadSection').classList.add('hidden');
  document.getElementById('loadingSection').classList.remove('hidden');
}
function hideLoading() { document.getElementById('loadingSection').classList.add('hidden'); }
function showError(msg) {
  document.getElementById('errorText').textContent = msg;
  document.getElementById('errorMsg').classList.remove('hidden');
  document.getElementById('uploadSection').classList.remove('hidden');
  hideLoading();
}
function hideError() { document.getElementById('errorMsg').classList.add('hidden'); }

// ══════════════════════════════════════════════════════
// 10. HELPERS DE DATOS
// ══════════════════════════════════════════════════════

function agrupar(rows, key) {
  return rows.reduce((acc,r)=>{ acc[r[key]]=(acc[r[key]]||0)+r.monto; return acc; },{});
}
function contar(rows, key) {
  return rows.reduce((acc,r)=>{ acc[r[key]]=(acc[r[key]]||0)+1; return acc; },{});
}

// ── Parsear fecha ──
function parseDate(val) {
  if (val == null || val === '') return null;

  // Ya es Date (SheetJS con cellDates:true)
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;

  // Número → serial de Excel (ej: 42739 = 4-Ene-2017)
  if (typeof val === 'number') {
    // Excel serial: días desde 1-Ene-1900, con bug de año bisiesto 1900
    // Equivalente JS: (serial - 25569) días desde 1-Ene-1970 en UTC
    if (val < 1 || val > 2958465) return null; // rango sensato (1900–9999)
    const ms  = Math.round((val - 25569) * 86400 * 1000);
    const utc = new Date(ms);
    // Convertir a fecha local (evitar off-by-one por timezone)
    return new Date(utc.getUTCFullYear(), utc.getUTCMonth(), utc.getUTCDate());
  }

  // String
  const str = String(val).trim();
  if (!str) return null;

  // ISO y otros formatos parseables por el motor JS
  const direct = new Date(str);
  if (!isNaN(direct.getTime())) {
    return new Date(direct.getFullYear(), direct.getMonth(), direct.getDate());
  }

  // DD/MM/YYYY o DD-MM-YYYY
  const ddmm = str.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})$/);
  if (ddmm) {
    const d = new Date(+ddmm[3], +ddmm[2]-1, +ddmm[1]);
    if (!isNaN(d.getTime())) return d;
  }

  // YYYY/MM/DD
  const yyyymm = str.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})$/);
  if (yyyymm) {
    const d = new Date(+yyyymm[1], +yyyymm[2]-1, +yyyymm[3]);
    if (!isNaN(d.getTime())) return d;
  }

  return null;
}

function parseMonto(val) {
  if (val===null||val===undefined||val==='') return NaN;
  if (typeof val==='number') return val;
  const c = String(val).replace(/[\$\s,]/g,'').replace(/[()]/g,m=>m==='('?'-':'');
  return parseFloat(c);
}

function capitalizar(s) {
  if (!s||s==='—') return s;
  return s.charAt(0).toUpperCase()+s.slice(1);
}

function formatMoney(n) {
  return new Intl.NumberFormat('es-MX',{style:'currency',currency:'MXN',minimumFractionDigits:0,maximumFractionDigits:0}).format(n);
}
function formatMoneyShort(n) {
  if(Math.abs(n)>=1_000_000)return`$${(n/1_000_000).toFixed(1)}M`;
  if(Math.abs(n)>=1_000)return`$${(n/1_000).toFixed(0)}K`;
  return`$${n}`;
}
function formatDate(d) {
  return d.toLocaleDateString('es-MX',{day:'2-digit',month:'short',year:'numeric'});
}
function escHtml(s) {
  if(!s)return'';
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// Opciones base para tooltips de Chart.js
function tooltipDefaults(isDark, labelFn, isDonut=false) {
  return {
    backgroundColor: isDark?'#1e293b':'#fff',
    titleColor:      isDark?'#f1f5f9':'#0f172a',
    bodyColor:       isDark?'#94a3b8':'#475569',
    borderColor:     isDark?'rgba(255,255,255,0.1)':'rgba(0,0,0,0.1)',
    borderWidth:1, padding:12, cornerRadius:10,
    callbacks: {
      label: isDonut
        ? ctx => labelFn(ctx.parsed, ctx)
        : ctx => ` ${ctx.dataset.label||''}: ${labelFn(ctx.parsed.y ?? ctx.parsed)}`,
    },
  };
}
