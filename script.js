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
let allRows        = [];  // Todos los registros (todos los años)
let yearRows       = [];  // Registros del año seleccionado
let filteredRows   = [];  // Registros filtrados (año + mes)
let allYears       = [];  // Lista de años disponibles
let barChartInst   = null;
let donutChartInst = null;
let tipoBarInst    = null;
let detectedYear   = new Date().getFullYear();

// ─── ESTADO TABLA DE TRANSACCIONES ────────────────────
let txCurrentRows  = [];  // filas actuales antes de filtros de columna
let txColFilters   = {};  // colKey → Set<string> | vacío = sin filtro
let txOpenCol      = null; // columna cuyo dropdown está abierto

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
      // Sin cellDates:true — trabajamos con strings via raw:false
      const wb = XLSX.read(data, { type: 'array' });

      let rows = [];

      // ── MODO A: hojas INGRESOS + EGRESOS (formato SMTO) ──
      const hasIngresos = wb.SheetNames.some(n => n.toUpperCase().includes('INGRESO'));
      const hasEgresos  = wb.SheetNames.some(n => n.toUpperCase().includes('EGRESO'));

      if (hasIngresos && hasEgresos) {
        rows = parseMultiSheet(wb);
      } else {
        // ── MODO B: hoja única genérica ──
        rows = parseSingleSheet(wb.Sheets[wb.SheetNames[0]]);
      }

      if (rows.length === 0) {
        throw new Error(
          `No se encontraron transacciones válidas en el archivo. ` +
          `Hojas: [${wb.SheetNames.join(', ')}]. ` +
          'Verifica que tenga hojas INGRESOS y EGRESOS con columnas de Fecha, Tipo y Total.'
        );
      }

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
// 3A. PARSER MULTI-HOJA
//     Estrategia: sheet_to_json con raw:false + dateNF.
//     Todo llega como string → sin problemas de tipos.
// ══════════════════════════════════════════════════════

/** Lee un worksheet como array de arrays de strings (sin tipos). */
function wsToStrings(ws) {
  return XLSX.utils.sheet_to_json(ws, {
    header:  1,
    raw:     false,       // TODO como string formateado
    dateNF:  'yyyy-mm-dd', // Fechas → "2017-01-04"
    defval:  '',          // Celdas vacías → ""
  });
}

/** Elige hoja por nombre: exacta primero, luego la más corta con la palabra. */
function findSheetName(wb, keyword) {
  const kw = keyword.toUpperCase();
  const exact = wb.SheetNames.find(n => {
    const u = n.trim().toUpperCase();
    return u === kw || u === kw + 'S';
  });
  if (exact) return exact;
  const partials = wb.SheetNames.filter(n => n.trim().toUpperCase().includes(kw));
  return partials.length ? partials.sort((a,b)=>a.length-b.length)[0] : null;
}

/**
 * Busca la fila de encabezados y mapea nombres → índice de columna.
 * @returns {{ headerIdx, colMap }} o null
 */
function detectHeader(rows, dateHints, maxScan) {
  maxScan = Math.min(maxScan || 25, rows.length);
  for (let i = 0; i < maxScan; i++) {
    const row = rows[i];
    if (!row || !row.some(c => c)) continue;
    // Verificar si esta fila contiene una columna de fecha
    const hasDate = row.some(cell => {
      const cu = String(cell).toUpperCase().trim();
      return dateHints.some(h => cu.includes(h.toUpperCase()));
    });
    if (!hasDate) continue;
    // Construir mapa de nombre→índice
    const colMap = {};
    row.forEach((cell, idx) => {
      const k = String(cell).toUpperCase().trim();
      if (k) colMap[k] = idx;
    });
    return { headerIdx: i, colMap };
  }
  return null;
}

/**
 * Busca el índice de columna usando hints (exacto → parcial).
 */
function findColIdx(colMap, hints) {
  // Exacta
  for (const h of hints) {
    const k = h.toUpperCase().trim();
    if (colMap[k] !== undefined) return colMap[k];
  }
  // Parcial: el nombre de la columna contiene el hint
  for (const h of hints) {
    const k = h.toUpperCase().trim();
    const key = Object.keys(colMap).find(ck => ck.includes(k));
    if (key !== undefined) return colMap[key];
  }
  return -1;
}

/**
 * Parser principal: strings → filas normalizadas.
 * Recibe el worksheet, el tipo (Ingreso/Egreso) y hints de columnas.
 */
function parseSheetStrings(ws, tipoReg, dateHints, totalHints, tipoHints, nameHints,
                           importeHints, ivaHints, retHints) {
  if (!ws) return [];
  const rows = wsToStrings(ws);
  if (!rows.length) return [];

  const hdr = detectHeader(rows, dateHints);
  if (!hdr) return [];

  const { headerIdx, colMap } = hdr;
  const cFecha   = findColIdx(colMap, dateHints);
  const cTotal   = findColIdx(colMap, totalHints);
  const cTipo    = findColIdx(colMap, tipoHints);
  const cNombre  = findColIdx(colMap, nameHints);
  const cImporte = importeHints ? findColIdx(colMap, importeHints) : -1;
  const cIva     = ivaHints     ? findColIdx(colMap, ivaHints)     : -1;
  const cRet     = retHints     ? findColIdx(colMap, retHints)     : -1;

  if (cFecha === -1 || cTotal === -1) return [];

  const result = [];
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || !row.some(c => c)) continue;

    const rawFecha = row[cFecha] || '';
    const rawTotal = row[cTotal] || '';
    if (!rawFecha && !rawTotal) continue;

    const fecha = parseDate(rawFecha);
    if (!fecha) continue;

    const total = parseMonto(rawTotal);
    if (isNaN(total) || total === 0) continue;

    const tipo    = cTipo   >= 0 && row[cTipo]   ? String(row[cTipo]).trim()   : 'Sin tipo';
    const nombre  = cNombre >= 0 && row[cNombre] ? String(row[cNombre]).trim() : '—';
    const importe = cImporte >= 0 ? (parseMonto(row[cImporte] || '0') || 0) : Math.abs(total);
    const iva     = cIva     >= 0 ? (parseMonto(row[cIva]     || '0') || 0) : 0;
    const ret     = cRet     >= 0 ? (parseMonto(row[cRet]     || '0') || 0) : 0;

    result.push({
      fecha,
      year:          fecha.getFullYear(),
      mes:           fecha.getMonth(),
      tipo_registro: tipoReg,
      tipo,
      categoria:     tipo,
      subcategoria:  nombre,
      monto:         Math.abs(total),
      importe:       Math.abs(importe),
      iva:           Math.abs(iva),
      ret:           Math.abs(ret),
    });
  }
  return result;
}

function parseMultiSheet(wb) {
  const nameIng = findSheetName(wb, 'INGRESO');
  const nameEgr = findSheetName(wb, 'EGRESO');
  if (!nameIng || !nameEgr) throw new Error('No se encontraron hojas INGRESOS y EGRESOS.');

  const rowsIng = parseSheetStrings(
    wb.Sheets[nameIng], 'Ingreso',
    ['FECHA DE PAGO', 'FECHA PAGO', 'FECHA'],           // date
    ['TOTAL'],                                           // total
    ['TIPO'],                                            // tipo
    ['NOMBRE DEL CLIENTE', 'NOMBRE CLIENTE', 'NOMBRE'], // nombre
    ['IMPORTE'],                                         // importe
    ['IVA'],                                             // iva
    null                                                 // sin ret en INGRESOS
  );

  const rowsEgr = parseSheetStrings(
    wb.Sheets[nameEgr], 'Egreso',
    ['FECHA FAC', 'FECHA FACTURA', 'FECHA'],            // date
    ['TOTAL'],                                           // total
    ['TIPO'],                                            // tipo
    ['PROVEEDOR'],                                       // nombre
    ['IMPORTE'],                                         // importe
    ['IVA'],                                             // iva
    ['RET', 'RET/ ISR', 'RET/ISR', 'RETENCION']        // ret
  );

  const combined = [...rowsIng, ...rowsEgr];
  if (combined.length === 0) return [];

  // Auto-detectar el año más frecuente en los datos
  const yearCount = {};
  for (const r of combined) yearCount[r.year] = (yearCount[r.year] || 0) + 1;
  detectedYear = parseInt(Object.entries(yearCount).sort((a,b) => b[1]-a[1])[0][0], 10);

  // Retornar TODOS los años (el filtrado por año ocurre en buildDashboard)
  return combined;
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

  // Auto-detectar año dominante
  const yearCount = {};
  for (const r of rows) yearCount[r.year] = (yearCount[r.year] || 0) + 1;
  detectedYear = parseInt(Object.entries(yearCount).sort((a,b) => b[1]-a[1])[0][0], 10);

  // Retornar TODOS los años
  return rows;
}

/** Wrapper: recibe un worksheet, normaliza con la función genérica. */
function parseSingleSheet(ws) {
  const raw = XLSX.utils.sheet_to_json(ws, {
    raw: false, dateNF: 'yyyy-mm-dd', defval: '',
  });
  if (!raw || raw.length === 0) throw new Error('La hoja no contiene datos.');
  return normalizeSingleSheet(raw);
}

// ══════════════════════════════════════════════════════
// 4. CONSTRUCCIÓN DEL DASHBOARD
// ══════════════════════════════════════════════════════

function buildDashboard(rows) {
  allRows  = rows;
  allYears = [...new Set(rows.map(r => r.year))].sort((a,b) => a - b);

  // Por defecto mostrar TODOS los datos (para que el total coincida con el Excel)
  yearRows     = rows;
  filteredRows = rows;

  renderKPIs(filteredRows);
  renderBarChart(filteredRows);
  renderDonutChart(filteredRows);
  renderTipoCharts(filteredRows);
  renderCategoryTable(filteredRows);
  renderTxTable(filteredRows);
  buildYearFilter(rows);
  buildMonthFilter(rows);

  hideLoading();
  document.getElementById('uploadSection').classList.add('hidden');
  document.getElementById('dashboardSection').classList.remove('hidden');
  document.getElementById('dashSubtitle').textContent = 'Datos completos';
  document.getElementById('yearFilterWrapper').classList.remove('hidden');
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
  const egr    = rows.filter(r => r.tipo_registro === 'Egreso');
  const byTipo = agrupar(egr, 'tipo');
  const total  = egr.reduce((s,r) => s + r.monto, 0);

  // Top 8 categorías + "Otros" para no saturar la leyenda
  const all    = Object.entries(byTipo).sort((a,b) => b[1] - a[1]);
  const top    = all.slice(0, 8);
  const othersSum = all.slice(8).reduce((s,[,v]) => s + v, 0);
  const slices = othersSum > 0 ? [...top, ['Otros', othersSum]] : top;

  document.getElementById('donutTotal').textContent = formatMoney(total);

  const isDark = document.documentElement.getAttribute('data-theme') !== 'light';
  const ctx    = document.getElementById('donutChart').getContext('2d');
  if (donutChartInst) donutChartInst.destroy();
  donutChartInst = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: slices.map(([t]) => t),
      datasets: [{
        data:            slices.map(([,v]) => v),
        backgroundColor: slices.map((_,i) => PALETTE[i % PALETTE.length]),
        borderColor:     isDark ? '#1e293b' : '#fff',
        borderWidth: 2,
        hoverOffset: 8,
      }],
    },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: '68%',
      layout: { padding: { top: 8, bottom: 8 } },
      plugins: {
        legend: {
          position: 'right',
          labels: {
            color:         isDark ? '#94a3b8' : '#475569',
            font:          { family:'Inter', size:11 },
            padding:       10,
            boxWidth:      10,
            usePointStyle: true,
            // Truncar etiquetas largas
            generateLabels: chart => {
              const ds = chart.data.datasets[0];
              return chart.data.labels.map((lbl, i) => ({
                text:        lbl.length > 18 ? lbl.slice(0,16)+'…' : lbl,
                fillStyle:   ds.backgroundColor[i],
                strokeStyle: ds.borderColor,
                lineWidth:   ds.borderWidth,
                hidden:      false,
                index:       i,
              }));
            },
          },
        },
        tooltip: {
          backgroundColor: isDark?'#1e293b':'#fff',
          titleColor:      isDark?'#f1f5f9':'#0f172a',
          bodyColor:       isDark?'#94a3b8':'#475569',
          borderColor:     isDark?'rgba(255,255,255,0.1)':'rgba(0,0,0,0.1)',
          borderWidth:1, padding:12, cornerRadius:10,
          callbacks: {
            label: ctx => {
              const v   = ctx.dataset.data[ctx.dataIndex];
              const pct = total > 0 ? (v / total * 100).toFixed(1) : '0';
              return ` ${ctx.label}: ${formatMoney(v)} (${pct}%)`;
            },
          },
        },
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

// ══════════════════════════════════════════════════════
// TABLA AVANZADA DE TRANSACCIONES
// Columnas filtrable + totales dinámicos
// ══════════════════════════════════════════════════════

const TX_COLS = [
  { key:'fecha_str',    label:'Fecha',            align:'left',  filterable:true,  numeric:false },
  { key:'tipo_registro',label:'Movimiento',       align:'left',  filterable:true,  numeric:false },
  { key:'tipo',         label:'Tipo',             align:'left',  filterable:true,  numeric:false },
  { key:'subcategoria', label:'Proveedor / Cliente',align:'left',filterable:true,  numeric:false },
  { key:'importe',      label:'Importe',          align:'right', filterable:false, numeric:true  },
  { key:'iva',          label:'IVA',              align:'right', filterable:false, numeric:true  },
  { key:'ret',          label:'Ret',              align:'right', filterable:false, numeric:true  },
  { key:'monto',        label:'Total',            align:'right', filterable:false, numeric:true  },
];

function txGetVal(row, key) {
  if (key === 'fecha_str') return formatDate(row.fecha);
  return String(row[key] ?? '');
}

function renderTxTable(rows) {
  txCurrentRows = [...rows].sort((a,b) => b.fecha - a.fecha);
  txColFilters  = {};          // reset column filters when year/month changes
  txOpenCol     = null;
  buildTxHeader();
  refreshTxTable();
}

function refreshTxTable() {
  const visible = getTxFiltered();
  renderTxBody(visible);
  renderTxFooter(visible);
  const badge = document.getElementById('txBadge');
  badge.textContent = visible.length === txCurrentRows.length
    ? `${visible.length} registros`
    : `${visible.length} de ${txCurrentRows.length} registros`;
  // Update active-filter dots on headers
  TX_COLS.filter(c => c.filterable).forEach(c => {
    const btn = document.querySelector(`.tx-filter-btn[data-col="${c.key}"]`);
    if (btn) btn.classList.toggle('active', !!(txColFilters[c.key] && txColFilters[c.key].size > 0));
  });
}

function getTxFiltered() {
  return txCurrentRows.filter(row => {
    for (const [key, allowed] of Object.entries(txColFilters)) {
      if (!allowed || allowed.size === 0) continue;
      if (!allowed.has(txGetVal(row, key))) return false;
    }
    return true;
  });
}

function buildTxHeader() {
  const thead = document.querySelector('#txTable thead tr');
  if (!thead) return;
  thead.innerHTML = '';
  TX_COLS.forEach(col => {
    const th = document.createElement('th');
    th.className = col.align === 'right' ? 'text-right' : '';
    if (col.filterable) {
      th.innerHTML = `
        <span class="th-label">${col.label}</span>
        <button class="tx-filter-btn" data-col="${col.key}" title="Filtrar por ${col.label}">
          <svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
            <polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3"/>
          </svg>
        </button>`;
      th.querySelector('.tx-filter-btn').addEventListener('click', e => {
        e.stopPropagation();
        toggleTxDropdown(col.key, e.currentTarget);
      });
    } else {
      th.innerHTML = `<span class="th-label">${col.label}</span>`;
    }
    thead.appendChild(th);
  });
}

function renderTxBody(rows) {
  const tbody = document.getElementById('txTableBody');
  tbody.innerHTML = '';
  rows.forEach(r => {
    const isInc = r.tipo_registro === 'Ingreso';
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="td-date">${escHtml(formatDate(r.fecha))}</td>
      <td><span class="type-badge ${isInc?'type-income':'type-expense'}">${r.tipo_registro}</span></td>
      <td><strong class="td-tipo">${escHtml(r.tipo)}</strong></td>
      <td class="td-nombre">${escHtml(r.subcategoria)}</td>
      <td class="text-right td-num">${r.importe > 0 ? formatMoney(r.importe) : '—'}</td>
      <td class="text-right td-num">${r.iva     > 0 ? formatMoney(r.iva)     : '—'}</td>
      <td class="text-right td-num">${r.ret     > 0 ? formatMoney(r.ret)     : '—'}</td>
      <td class="text-right td-total" style="color:${isInc?'var(--income)':'var(--expense)'}">
        ${isInc?'+':'-'}${formatMoney(r.monto)}
      </td>`;
    tbody.appendChild(tr);
  });
}

function renderTxFooter(rows) {
  const sumImporte = rows.reduce((s,r)=>s+r.importe,0);
  const sumIva     = rows.reduce((s,r)=>s+r.iva,    0);
  const sumRet     = rows.reduce((s,r)=>s+r.ret,    0);
  const sumTotal   = rows.reduce((s,r)=>s+r.monto,  0);
  const el = id => document.getElementById(id);
  if (el('footImporte')) el('footImporte').textContent = formatMoney(sumImporte);
  if (el('footIva'))     el('footIva').textContent     = formatMoney(sumIva);
  if (el('footRet'))     el('footRet').textContent     = formatMoney(sumRet);
  if (el('footTotal'))   el('footTotal').textContent   = formatMoney(sumTotal);
  if (el('footCount'))   el('footCount').textContent   = `${rows.length} registros`;
}

// ── DROPDOWN DE FILTRO ──
function toggleTxDropdown(colKey, btn) {
  const existing = document.getElementById('txDropdownPanel');
  if (txOpenCol === colKey && existing) { closeTxDropdown(); return; }
  closeTxDropdown();
  txOpenCol = colKey;
  openTxDropdown(colKey, btn);
}

function openTxDropdown(colKey, anchorBtn) {
  // Unique values in current ALL rows (before column filter for other cols)
  const uniqueVals = [...new Set(txCurrentRows.map(r => txGetVal(r, colKey)))].sort();
  const activeSet  = txColFilters[colKey] || null;

  const panel = document.createElement('div');
  panel.id = 'txDropdownPanel';
  panel.className = 'tx-dropdown-panel';
  panel.innerHTML = `
    <div class="tx-dp-head">
      <input id="txDpSearch" class="tx-dp-search" placeholder="Buscar…" autocomplete="off"/>
      <div class="tx-dp-actions">
        <button id="txDpAll">Todos</button>
        <button id="txDpNone">Ninguno</button>
      </div>
    </div>
    <div id="txDpList" class="tx-dp-list"></div>
    <div class="tx-dp-foot">
      <button id="txDpApply" class="tx-dp-btn-apply">Aplicar</button>
      <button id="txDpClear" class="tx-dp-btn-clear">Limpiar filtro</button>
    </div>`;
  document.body.appendChild(panel);

  // Populate checkboxes
  const list = panel.querySelector('#txDpList');
  function renderOptions(filter) {
    list.innerHTML = '';
    uniqueVals
      .filter(v => !filter || v.toLowerCase().includes(filter))
      .forEach(val => {
        const checked = !activeSet || activeSet.has(val);
        const item = document.createElement('label');
        item.className = 'tx-dp-item';
        item.innerHTML = `
          <input type="checkbox" value="${escHtml(val)}" ${checked?'checked':''}>
          <span>${escHtml(val)}</span>`;
        list.appendChild(item);
      });
  }
  renderOptions('');

  panel.querySelector('#txDpSearch').addEventListener('input', e => renderOptions(e.target.value.toLowerCase()));
  panel.querySelector('#txDpAll').addEventListener('click',  () => list.querySelectorAll('input').forEach(i=>i.checked=true));
  panel.querySelector('#txDpNone').addEventListener('click', () => list.querySelectorAll('input').forEach(i=>i.checked=false));
  panel.querySelector('#txDpApply').addEventListener('click', () => {
    const checked = [...list.querySelectorAll('input:checked')].map(i=>i.value);
    if (checked.length === uniqueVals.length) {
      delete txColFilters[colKey];
    } else {
      txColFilters[colKey] = new Set(checked);
    }
    closeTxDropdown();
    refreshTxTable();
  });
  panel.querySelector('#txDpClear').addEventListener('click', () => {
    delete txColFilters[colKey];
    closeTxDropdown();
    refreshTxTable();
  });

  // Position below button
  const rect = anchorBtn.getBoundingClientRect();
  panel.style.top  = `${rect.bottom + window.scrollY + 4}px`;
  panel.style.left = `${Math.min(rect.left + window.scrollX, window.innerWidth - 280)}px`;

  // Close on outside click
  setTimeout(() => document.addEventListener('click', outsideTxClick), 0);
}

function outsideTxClick(e) {
  const panel = document.getElementById('txDropdownPanel');
  if (panel && !panel.contains(e.target)) closeTxDropdown();
}

function closeTxDropdown() {
  const panel = document.getElementById('txDropdownPanel');
  if (panel) panel.remove();
  txOpenCol = null;
  document.removeEventListener('click', outsideTxClick);
}

// ══════════════════════════════════════════════════════
// 5A. FILTRO POR AÑO
// ══════════════════════════════════════════════════════

function buildYearFilter(rows) {
  const years = [...new Set(rows.map(r => r.year))].sort((a,b) => a - b);
  const sel   = document.getElementById('yearFilter');
  while (sel.options.length > 1) sel.remove(1);
  for (const y of years) {
    const opt = document.createElement('option');
    opt.value = y; opt.textContent = y;
    sel.appendChild(opt);
  }
  sel.value = 'all';  // mostrar todo por defecto
  sel.onchange = () => applyYearFilter(sel.value);
}

function applyYearFilter(val) {
  yearRows     = val === 'all' ? allRows : allRows.filter(r => r.year === parseInt(val, 10));
  filteredRows = yearRows;
  buildMonthFilter(yearRows);
  document.getElementById('monthFilter').value = 'all';
  renderKPIs(filteredRows);
  renderBarChart(filteredRows);
  renderDonutChart(filteredRows);
  renderTipoCharts(filteredRows);
  renderCategoryTable(filteredRows);
  renderTxTable(filteredRows);
  const label = val === 'all' ? 'Todos los años' : `Año ${val}`;
  document.getElementById('dashSubtitle').textContent = label;
}

// ══════════════════════════════════════════════════════
// 5B. FILTRO POR MES
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
  filteredRows = val==='all' ? yearRows : yearRows.filter(r=>r.mes===parseInt(val,10));
  renderKPIs(filteredRows);
  renderBarChart(filteredRows);
  renderDonutChart(filteredRows);
  renderTipoCharts(filteredRows);
  renderCategoryTable(filteredRows);
  renderTxTable(filteredRows);
  const yearSel = document.getElementById('yearFilter').value;
  const yearLabel = yearSel === 'all' ? 'Todos los años' : `Año ${yearSel}`;
  const label = val==='all' ? yearLabel : `${MONTHS_LONG[parseInt(val,10)]} — ${yearLabel}`;
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
  allRows=[]; yearRows=[]; filteredRows=[]; allYears=[];
  txCurrentRows=[]; txColFilters={}; txOpenCol=null;
  closeTxDropdown();
  document.getElementById('dashboardSection').classList.add('hidden');
  document.getElementById('yearFilterWrapper').classList.add('hidden');
  document.getElementById('monthFilterWrapper').classList.add('hidden');
  document.getElementById('exportCsvBtn').classList.add('hidden');
  document.getElementById('uploadSection').classList.remove('hidden');
  document.getElementById('yearFilter').value='all';
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
// horizontal=true para barras horizontales (indexAxis:'y')
function tooltipDefaults(isDark, labelFn, isDonut=false, horizontal=false) {
  return {
    backgroundColor: isDark?'#1e293b':'#fff',
    titleColor:      isDark?'#f1f5f9':'#0f172a',
    bodyColor:       isDark?'#94a3b8':'#475569',
    borderColor:     isDark?'rgba(255,255,255,0.1)':'rgba(0,0,0,0.1)',
    borderWidth:1, padding:12, cornerRadius:10,
    callbacks: {
      label: isDonut
        ? ctx => labelFn(ctx.parsed, ctx)
        // Acceso directo al array de datos — más confiable que ctx.parsed en ambos ejes
        : ctx => ` ${ctx.dataset.label||''}: ${labelFn(ctx.dataset.data[ctx.dataIndex])}`,
    },
  };
}
