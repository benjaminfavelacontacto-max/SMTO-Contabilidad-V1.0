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

// ─── ESTADO TABLA DESGLOSE POR TIPO ───────────────────
let catAllRows     = [];  // todas las filas del desglose (rebuilt en render)
let catColFilters  = {};  // colKey → Set<string>
let catOpenCol     = null;

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

    // Fecha puede ser null — esas filas se incluyen como "Sin fecha"
    const fecha = parseDate(rawFecha);

    const total = parseMonto(rawTotal);
    if (isNaN(total) || total === 0) continue;

    const tipo    = cTipo   >= 0 && row[cTipo]   ? String(row[cTipo]).trim()   : 'Sin tipo';
    const nombre  = cNombre >= 0 && row[cNombre] ? String(row[cNombre]).trim() : '—';
    const importe = cImporte >= 0 ? (parseMonto(row[cImporte] || '0') || 0) : Math.abs(total);
    const iva     = cIva     >= 0 ? (parseMonto(row[cIva]     || '0') || 0) : 0;
    const ret     = cRet     >= 0 ? (parseMonto(row[cRet]     || '0') || 0) : 0;

    result.push({
      fecha,                                          // Date | null
      year:          fecha ? fecha.getFullYear() : 'Sin fecha',
      mes:           fecha ? fecha.getMonth()    : null,
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

  // Auto-detectar el año más frecuente (solo filas con fecha válida)
  const yearCount = {};
  for (const r of combined) {
    if (r.year !== 'Sin fecha') yearCount[r.year] = (yearCount[r.year] || 0) + 1;
  }
  const topYear = Object.entries(yearCount).sort((a,b) => b[1]-a[1])[0];
  detectedYear = topYear ? parseInt(topYear[0], 10) : new Date().getFullYear();

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
      const fecha = parseDate(row[colFecha]);   // puede ser null → 'Sin fecha'
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
        year:          fecha ? fecha.getFullYear() : 'Sin fecha',
        mes:           fecha ? fecha.getMonth()    : null,
        tipo_registro: tipoReg,
        tipo:          capitalizar(tipoVal),
        categoria:     capitalizar(cat),
        subcategoria:  capitalizar(subcat),
        monto:         Math.abs(monto),
      });
    } catch (_) { /* fila con error: ignorar */ }
  }

  if (rows.length === 0) return [];

  // Auto-detectar año dominante (solo filas con fecha válida)
  const yearCount = {};
  for (const r of rows) {
    if (r.year !== 'Sin fecha') yearCount[r.year] = (yearCount[r.year] || 0) + 1;
  }
  const topYear = Object.entries(yearCount).sort((a,b) => b[1]-a[1])[0];
  detectedYear = topYear ? parseInt(topYear[0], 10) : new Date().getFullYear();

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
    if (r.mes === null) continue;  // filas sin fecha no se grafican por mes
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

  // Top 8 + "Otros"
  const all       = Object.entries(byTipo).sort((a,b) => b[1] - a[1]);
  const top       = all.slice(0, 8);
  const othersSum = all.slice(8).reduce((s,[,v]) => s + v, 0);
  const slices    = othersSum > 0 ? [...top, ['Otros', othersSum]] : top;

  const isDark     = document.documentElement.getAttribute('data-theme') !== 'light';
  const lblColor   = '#ffffff';                          // blanco para contraste
  const subColor   = isDark ? '#94a3b8' : '#64748b';
  const centerMain = isDark ? '#f1f5f9' : '#0f172a';

  // Plugin: dibuja TOTAL + valor en el centro del anillo (canvas-space)
  const centerPlugin = {
    id: 'donutCenter',
    beforeDraw(chart) {
      const { ctx: c, chartArea } = chart;
      if (!chartArea) return;
      const cx = chartArea.left + chartArea.width  / 2;
      const cy = chartArea.top  + chartArea.height / 2;
      c.save();
      c.textAlign    = 'center';
      c.textBaseline = 'middle';
      // Sub-label
      c.font      = '500 11px Inter, system-ui, sans-serif';
      c.fillStyle = subColor;
      c.fillText('TOTAL', cx, cy - 15);
      // Valor principal
      c.font      = '800 18px Inter, system-ui, sans-serif';
      c.fillStyle = centerMain;
      c.fillText(formatMoneyShort(total), cx, cy + 10);
      c.restore();
    },
  };

  const ctxEl = document.getElementById('donutChart').getContext('2d');
  if (donutChartInst) donutChartInst.destroy();
  donutChartInst = new Chart(ctxEl, {
    type: 'doughnut',
    plugins: [centerPlugin],
    data: {
      labels: slices.map(([t]) => t),
      datasets: [{
        data:            slices.map(([,v]) => v),
        backgroundColor: slices.map((_,i) => PALETTE[i % PALETTE.length]),
        borderColor:     isDark ? '#1e293b' : '#f8fafc',
        borderWidth:     2,
        hoverOffset:     10,
      }],
    },
    options: {
      responsive:        true,
      maintainAspectRatio: false,
      cutout:            '68%',
      layout:            { padding: { top: 4, bottom: 4, left: 4, right: 4 } },
      plugins: {
        legend: {
          position: 'bottom',
          align:    'center',
          labels: {
            color:         lblColor,
            font:          { family: 'Inter', size: 11, weight: '500' },
            padding:       16,
            boxWidth:      10,
            usePointStyle: true,
            pointStyleWidth: 8,
          },
        },
        tooltip: {
          backgroundColor: isDark ? '#1e293b' : '#fff',
          titleColor:      isDark ? '#f1f5f9' : '#0f172a',
          bodyColor:       isDark ? '#94a3b8' : '#475569',
          borderColor:     isDark ? 'rgba(255,255,255,0.12)' : 'rgba(0,0,0,0.08)',
          borderWidth: 1, padding: 12, cornerRadius: 10,
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

// ══════════════════════════════════════════════════════
// TABLA DESGLOSE POR TIPO — con filtros tipo Excel
// ══════════════════════════════════════════════════════

const CAT_COLS = [
  { key:'cat',   label:'Tipo',          align:'left',  filterable:true  },
  { key:'tipo',  label:'Movimiento',    align:'left',  filterable:true  },
  { key:'total', label:'Total',         align:'right', filterable:false },
  { key:'pct',   label:'% del Gasto',   align:'right', filterable:false },
  { key:'count', label:'Transacciones', align:'right', filterable:false },
];

function catGetVal(row, key) { return String(row[key] ?? ''); }

function renderCategoryTable(rows) {
  const egr      = rows.filter(r => r.tipo_registro === 'Egreso');
  const ing      = rows.filter(r => r.tipo_registro === 'Ingreso');
  const totalEgr = egr.reduce((s, r) => s + r.monto, 0);

  const byCatE = agrupar(egr, 'tipo');
  const byCatI = agrupar(ing, 'tipo');
  const cntE   = contar(egr, 'tipo');
  const cntI   = contar(ing, 'tipo');

  const allCats = new Set([...Object.keys(byCatE), ...Object.keys(byCatI)]);
  const entries = [];
  for (const cat of allCats) {
    if (byCatE[cat]) {
      const t = byCatE[cat];
      entries.push({ cat, tipo:'Egreso',  total:t, count:cntE[cat]||0, pct: totalEgr > 0 ? t/totalEgr*100 : 0 });
    }
    if (byCatI[cat]) {
      entries.push({ cat, tipo:'Ingreso', total:byCatI[cat], count:cntI[cat]||0, pct: 0 });
    }
  }
  entries.sort((a, b) => b.total - a.total);

  catAllRows    = entries;
  catColFilters = {};
  catOpenCol    = null;
  buildCatHeader();
  refreshCatTable();
}

function buildCatHeader() {
  const thead = document.querySelector('#categoryTable thead tr');
  if (!thead) return;
  thead.innerHTML = '';
  CAT_COLS.forEach(col => {
    const th = document.createElement('th');
    th.className = col.align === 'right' ? 'text-right cat-th-right' : 'cat-th-left';
    if (col.filterable) {
      th.innerHTML = `
        <span class="th-label">${col.label}</span>
        <button class="tx-filter-btn cat-filter-btn" data-catcol="${col.key}" title="Filtrar por ${col.label}">
          <svg width="11" height="11" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
            <polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3"/>
          </svg>
        </button>`;
      th.querySelector('.cat-filter-btn').addEventListener('click', e => {
        e.stopPropagation();
        toggleCatDropdown(col.key, e.currentTarget);
      });
    } else {
      th.innerHTML = `<span class="th-label">${col.label}</span>`;
    }
    thead.appendChild(th);
  });
}

function getCatFiltered() {
  return catAllRows.filter(row => {
    for (const [key, allowed] of Object.entries(catColFilters)) {
      if (!allowed || allowed.size === 0) continue;
      if (!allowed.has(catGetVal(row, key))) return false;
    }
    return true;
  });
}

function refreshCatTable() {
  const visible = getCatFiltered();
  renderCatBody(visible);
  document.getElementById('tableBadge').textContent =
    visible.length === catAllRows.length
      ? `${visible.length} categorías`
      : `${visible.length} de ${catAllRows.length} categorías`;

  // Actualizar estado de botones filtro
  CAT_COLS.filter(c => c.filterable).forEach(c => {
    const btn = document.querySelector(`.cat-filter-btn[data-catcol="${c.key}"]`);
    if (!btn) return;
    const isActive = !!(catColFilters[c.key] && catColFilters[c.key].size > 0);
    btn.classList.toggle('active', isActive);
    let dot = btn.querySelector('.tx-filter-dot');
    if (isActive && !dot) { dot = document.createElement('span'); dot.className='tx-filter-dot'; btn.appendChild(dot); }
    if (!isActive && dot) dot.remove();
  });
}

function renderCatBody(entries) {
  const tbody = document.getElementById('categoryTableBody');
  const frag  = document.createDocumentFragment();
  for (const e of entries) {
    const isInc = e.tipo === 'Ingreso';
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="cat-td-nombre"><strong>${escHtml(e.cat)}</strong></td>
      <td class="cat-td-mov"><span class="type-badge ${isInc?'type-income':'type-expense'}">${e.tipo}</span></td>
      <td class="text-right cat-td-total">${formatMoney(e.total)}</td>
      <td class="text-right cat-td-pct">
        ${e.tipo==='Egreso' && e.pct > 0
          ? `<div class="progress-cell">
               <span class="pct-label">${e.pct.toFixed(1)}%</span>
               <div class="progress-bar"><div class="progress-fill" style="width:${Math.min(e.pct,100)}%"></div></div>
             </div>`
          : '<span class="td-nil">—</span>'}
      </td>
      <td class="text-right cat-td-count">${e.count.toLocaleString('es-MX')}</td>`;
    frag.appendChild(tr);
  }
  tbody.innerHTML = '';
  tbody.appendChild(frag);
}

// ── DROPDOWN PARA TABLA DESGLOSE ──
function toggleCatDropdown(colKey, btn) {
  const existing = document.getElementById('txDropdownPanel');
  if (catOpenCol === colKey && existing) { closeCatDropdown(); return; }
  closeCatDropdown();
  catOpenCol = colKey;
  openCatDropdown(colKey, btn);
}

function openCatDropdown(colKey, anchorBtn) {
  const col      = CAT_COLS.find(c => c.key === colKey);
  const allVals  = [...new Set(catAllRows.map(r => catGetVal(r, colKey)))].sort();
  const activeSet = catColFilters[colKey] || null;

  const panel = document.createElement('div');
  panel.id        = 'txDropdownPanel';   // mismo id → close function funciona
  panel.className = 'tx-dropdown-panel';
  panel.innerHTML = `
    <div class="tx-dp-title">${escHtml(col?.label || colKey)}</div>
    <div class="tx-dp-head">
      <div class="tx-dp-search-wrap">
        <svg class="tx-dp-search-icon" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
          <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
        </svg>
        <input id="txDpSearch" class="tx-dp-search" placeholder="Buscar…" autocomplete="off"/>
      </div>
    </div>
    <div id="txDpList" class="tx-dp-list"></div>
    <div class="tx-dp-foot">
      <button id="txDpApply" class="tx-dp-btn-apply">Aplicar</button>
      <button id="txDpClear" class="tx-dp-btn-clear">Limpiar</button>
    </div>`;
  document.body.appendChild(panel);

  const list = panel.querySelector('#txDpList');

  function renderOptions(filterText) {
    list.innerHTML = '';
    const filtered = allVals.filter(v => !filterText || v.toLowerCase().includes(filterText));
    // Select All
    const saRow = document.createElement('label');
    saRow.className = 'tx-dp-item tx-dp-select-all-item';
    const allChecked  = filtered.every(v => !activeSet || activeSet.has(v));
    const someChecked = filtered.some(v => !activeSet || activeSet.has(v));
    saRow.innerHTML = `<input type="checkbox" id="txDpCbAll" ${allChecked?'checked':''}>
      <span><strong>Seleccionar todo</strong></span><span class="tx-dp-count">${filtered.length}</span>`;
    const cbAll = saRow.querySelector('#txDpCbAll');
    if (!allChecked && someChecked) cbAll.indeterminate = true;
    cbAll.addEventListener('change', () => list.querySelectorAll('.tx-dp-value-cb').forEach(i=>i.checked=cbAll.checked));
    list.appendChild(saRow);
    const sep = document.createElement('div'); sep.className='tx-dp-sep'; list.appendChild(sep);

    filtered.forEach(val => {
      const checked = !activeSet || activeSet.has(val);
      const item = document.createElement('label');
      item.className = 'tx-dp-item';
      item.innerHTML = `<input type="checkbox" class="tx-dp-value-cb" value="${escHtml(val)}" ${checked?'checked':''}>
        <span title="${escHtml(val)}">${escHtml(val)}</span>`;
      list.appendChild(item);
    });
    list.addEventListener('change', e => {
      if (e.target.classList.contains('tx-dp-value-cb')) {
        const cbs = [...list.querySelectorAll('.tx-dp-value-cb')];
        const sel = cbs.filter(i=>i.checked).length;
        const cbA = list.querySelector('#txDpCbAll');
        if (cbA) { cbA.checked = sel===cbs.length; cbA.indeterminate = sel>0 && sel<cbs.length; }
      }
    });
  }
  renderOptions('');
  panel.querySelector('#txDpSearch').addEventListener('input', e => renderOptions(e.target.value.toLowerCase()));
  panel.querySelector('#txDpApply').addEventListener('click', () => {
    const checked = [...list.querySelectorAll('.tx-dp-value-cb:checked')].map(i=>i.value);
    if (checked.length === allVals.length || checked.length === 0) delete catColFilters[colKey];
    else catColFilters[colKey] = new Set(checked);
    closeCatDropdown(); refreshCatTable();
  });
  panel.querySelector('#txDpClear').addEventListener('click', () => {
    delete catColFilters[colKey]; closeCatDropdown(); refreshCatTable();
  });

  const rect   = anchorBtn.getBoundingClientRect();
  const panelW = 280;
  panel.style.top  = `${rect.bottom + 4}px`;
  panel.style.left = `${Math.min(rect.left, window.innerWidth - panelW - 8)}px`;
  requestAnimationFrame(() => {
    const ph = panel.offsetHeight;
    if (rect.bottom + ph > window.innerHeight - 8) panel.style.top = `${rect.top - ph - 4}px`;
  });
  setTimeout(() => document.addEventListener('click', outsideCatClick), 0);
}

function outsideCatClick(e) {
  const panel = document.getElementById('txDropdownPanel');
  if (panel && !panel.contains(e.target)) closeCatDropdown();
}
function closeCatDropdown() {
  const panel = document.getElementById('txDropdownPanel');
  if (panel) panel.remove();
  catOpenCol = null;
  document.removeEventListener('click', outsideCatClick);
}

// ══════════════════════════════════════════════════════
// TABLA AVANZADA DE TRANSACCIONES
// Filtros tipo Excel + totales dinámicos + sin límite de filas
// ══════════════════════════════════════════════════════

const TX_COLS = [
  { key:'fecha_str',    label:'Fecha',              align:'left',  filterable:true,  numeric:false },
  { key:'tipo_registro',label:'Movimiento',         align:'left',  filterable:true,  numeric:false },
  { key:'tipo',         label:'Tipo',               align:'left',  filterable:true,  numeric:false },
  { key:'subcategoria', label:'Proveedor / Cliente', align:'left', filterable:true,  numeric:false },
  { key:'importe',      label:'Importe',            align:'right', filterable:false, numeric:true  },
  { key:'iva',          label:'IVA',                align:'right', filterable:false, numeric:true  },
  { key:'ret',          label:'Ret',                align:'right', filterable:false, numeric:true  },
  { key:'monto',        label:'Total',              align:'right', filterable:false, numeric:true  },
];

function txGetVal(row, key) {
  if (key === 'fecha_str') return formatDate(row.fecha);
  if (key === 'year')      return String(row.year ?? 'Sin fecha');
  return String(row[key] ?? '');
}

/** Construye / actualiza el botón de filtro de Año en el header de la tabla */
function buildTxYearControl() {
  const btn = document.getElementById('txYearFilterBtn');
  if (!btn) return;
  btn.onclick = e => { e.stopPropagation(); toggleTxDropdown('year', btn); };
  syncYearBtnState();
}

function syncYearBtnState() {
  const btn = document.getElementById('txYearFilterBtn');
  if (!btn) return;
  const isActive = !!(txColFilters['year'] && txColFilters['year'].size > 0);
  btn.classList.toggle('active', isActive);
  const lbl = btn.querySelector('.tx-yr-label');
  if (lbl) lbl.textContent = isActive ? `Año (${txColFilters['year'].size})` : 'Año';
}

/** Cuenta cuántos filtros de columna están activos */
function txActiveFilterCount() {
  return Object.values(txColFilters).filter(s => s && s.size > 0).length;
}

function renderTxTable(rows) {
  txCurrentRows = [...rows].sort((a, b) => {
    if (!a.fecha && !b.fecha) return 0;
    if (!a.fecha) return 1;   // sin fecha → al final
    if (!b.fecha) return -1;
    return b.fecha - a.fecha;
  });
  txColFilters  = {};
  txOpenCol     = null;
  buildTxHeader();
  buildTxYearControl();
  refreshTxTable();
}

function refreshTxTable() {
  const visible = getTxFiltered();
  renderTxBody(visible);
  renderTxFooter(visible);

  // Badge de registros
  const badge = document.getElementById('txBadge');
  badge.textContent = visible.length === txCurrentRows.length
    ? `${visible.length.toLocaleString('es-MX')} registros`
    : `${visible.length.toLocaleString('es-MX')} de ${txCurrentRows.length.toLocaleString('es-MX')} registros`;

  // Resaltar botones de filtro activos + contador
  TX_COLS.filter(c => c.filterable).forEach(c => {
    const btn = document.querySelector(`.tx-filter-btn[data-col="${c.key}"]`);
    if (!btn) return;
    const isActive = !!(txColFilters[c.key] && txColFilters[c.key].size > 0);
    btn.classList.toggle('active', isActive);
    // Actualizar el dot indicador dentro del botón
    let dot = btn.querySelector('.tx-filter-dot');
    if (isActive) {
      if (!dot) {
        dot = document.createElement('span');
        dot.className = 'tx-filter-dot';
        btn.appendChild(dot);
      }
    } else {
      if (dot) dot.remove();
    }
  });

  // Botón "Limpiar todos" — aparece solo si hay filtros activos
  const clearAllBtn = document.getElementById('txClearAllBtn');
  if (clearAllBtn) {
    const count = txActiveFilterCount();
    clearAllBtn.style.display = count > 0 ? 'inline-flex' : 'none';
    clearAllBtn.textContent   = count > 1 ? `Limpiar ${count} filtros` : 'Limpiar filtro';
  }
  // Sincronizar botón de año
  syncYearBtnState();
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
          <svg width="11" height="11" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
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
  const frag  = document.createDocumentFragment();

  rows.forEach(r => {
    const isInc = r.tipo_registro === 'Ingreso';
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="td-date">${escHtml(formatDate(r.fecha))}</td>
      <td><span class="type-badge ${isInc ? 'type-income' : 'type-expense'}">${r.tipo_registro}</span></td>
      <td><strong class="td-tipo">${escHtml(r.tipo)}</strong></td>
      <td class="td-nombre">${escHtml(r.subcategoria)}</td>
      <td class="text-right td-num">${r.importe > 0 ? formatMoney(r.importe) : '<span class="td-nil">—</span>'}</td>
      <td class="text-right td-num">${r.iva     > 0 ? formatMoney(r.iva)     : '<span class="td-nil">—</span>'}</td>
      <td class="text-right td-num">${r.ret     > 0 ? formatMoney(r.ret)     : '<span class="td-nil">—</span>'}</td>
      <td class="text-right td-total ${isInc ? 'td-total-inc' : 'td-total-egr'}">
        ${isInc ? '+' : '-'}${formatMoney(r.monto)}
      </td>`;
    frag.appendChild(tr);
  });

  tbody.innerHTML = '';
  tbody.appendChild(frag);
}

function renderTxFooter(rows) {
  const sumImporte = rows.reduce((s, r) => s + r.importe, 0);
  const sumIva     = rows.reduce((s, r) => s + r.iva,     0);
  const sumRet     = rows.reduce((s, r) => s + r.ret,     0);
  const sumTotal   = rows.reduce((s, r) => s + r.monto,   0);

  const get = id => document.getElementById(id);
  if (get('footImporte')) get('footImporte').textContent = formatMoney(sumImporte);
  if (get('footIva'))     get('footIva').textContent     = formatMoney(sumIva);
  if (get('footRet'))     get('footRet').textContent     = formatMoney(sumRet);
  if (get('footTotal'))   get('footTotal').textContent   = formatMoney(sumTotal);
  if (get('footCount'))   get('footCount').textContent   = `${rows.length.toLocaleString('es-MX')} registros`;
}

// ── DROPDOWN TIPO EXCEL ──────────────────────────────

function toggleTxDropdown(colKey, btn) {
  const existing = document.getElementById('txDropdownPanel');
  if (txOpenCol === colKey && existing) { closeTxDropdown(); return; }
  closeTxDropdown();
  txOpenCol = colKey;
  openTxDropdown(colKey, btn);
}

function openTxDropdown(colKey, anchorBtn) {
  const colLabel   = TX_COLS.find(c => c.key === colKey)?.label || colKey;
  // Para el key 'year': orden numérico ascendente con 'Sin fecha' al final
  const rawVals  = [...new Set(txCurrentRows.map(r => txGetVal(r, colKey)))];
  const allVals  = colKey === 'year'
    ? rawVals.sort((a, b) => {
        if (a === 'Sin fecha') return 1;
        if (b === 'Sin fecha') return -1;
        return Number(a) - Number(b);
      })
    : rawVals.sort();
  const activeSet  = txColFilters[colKey] || null;   // null → todos activos
  const totalCount = allVals.length;

  // ── Panel ──
  const panel = document.createElement('div');
  panel.id        = 'txDropdownPanel';
  panel.className = 'tx-dropdown-panel';
  panel.innerHTML = `
    <div class="tx-dp-title">${escHtml(colLabel)}</div>
    <div class="tx-dp-head">
      <div class="tx-dp-search-wrap">
        <svg class="tx-dp-search-icon" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
          <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
        </svg>
        <input id="txDpSearch" class="tx-dp-search" placeholder="Buscar en ${totalCount} valores…" autocomplete="off"/>
      </div>
    </div>
    <div id="txDpList" class="tx-dp-list"></div>
    <div class="tx-dp-foot">
      <button id="txDpApply" class="tx-dp-btn-apply">Aplicar</button>
      <button id="txDpClear" class="tx-dp-btn-clear">Limpiar</button>
    </div>`;
  document.body.appendChild(panel);

  const list = panel.querySelector('#txDpList');

  // ── Renderiza checkboxes con "Seleccionar todo" al inicio ──
  function renderOptions(filterText) {
    list.innerHTML = '';
    const filtered = allVals.filter(v => !filterText || v.toLowerCase().includes(filterText));

    // Fila "Seleccionar todo" (solo sin búsqueda activa o con todos visibles)
    const selectAllRow = document.createElement('label');
    selectAllRow.className = 'tx-dp-item tx-dp-select-all-item';
    const allChecked   = filtered.every(v => !activeSet || activeSet.has(v));
    const someChecked  = filtered.some(v => !activeSet || activeSet.has(v));
    selectAllRow.innerHTML = `
      <input type="checkbox" id="txDpCbAll" ${allChecked ? 'checked' : ''}>
      <span><strong>Seleccionar todo</strong></span>
      <span class="tx-dp-count">${filtered.length}</span>`;
    const cbAll = selectAllRow.querySelector('#txDpCbAll');
    if (!allChecked && someChecked) cbAll.indeterminate = true;
    cbAll.addEventListener('change', () => {
      list.querySelectorAll('.tx-dp-value-cb').forEach(i => i.checked = cbAll.checked);
    });
    list.appendChild(selectAllRow);

    // Separador
    const sep = document.createElement('div');
    sep.className = 'tx-dp-sep';
    list.appendChild(sep);

    // Items individuales
    filtered.forEach(val => {
      const checked = !activeSet || activeSet.has(val);
      const item = document.createElement('label');
      item.className = 'tx-dp-item';
      item.innerHTML = `
        <input type="checkbox" class="tx-dp-value-cb" value="${escHtml(val)}" ${checked ? 'checked' : ''}>
        <span title="${escHtml(val)}">${escHtml(val)}</span>`;
      list.appendChild(item);
    });

    // Actualizar estado del "Seleccionar todo" cuando cambia un item
    list.addEventListener('change', e => {
      if (e.target.classList.contains('tx-dp-value-cb')) {
        const cbs      = [...list.querySelectorAll('.tx-dp-value-cb')];
        const total    = cbs.length;
        const selected = cbs.filter(i => i.checked).length;
        const cbA      = list.querySelector('#txDpCbAll');
        if (cbA) {
          cbA.checked       = selected === total;
          cbA.indeterminate = selected > 0 && selected < total;
        }
      }
    });
  }

  renderOptions('');

  // ── Búsqueda ──
  panel.querySelector('#txDpSearch').addEventListener('input', e => {
    renderOptions(e.target.value.toLowerCase());
  });

  // ── Aplicar ──
  panel.querySelector('#txDpApply').addEventListener('click', () => {
    const checked = [...list.querySelectorAll('.tx-dp-value-cb:checked')].map(i => i.value);
    if (checked.length === allVals.length) {
      delete txColFilters[colKey];
    } else if (checked.length === 0) {
      // Ninguno seleccionado: sin cambio (o limpiar)
      delete txColFilters[colKey];
    } else {
      txColFilters[colKey] = new Set(checked);
    }
    closeTxDropdown();
    refreshTxTable();
  });

  // ── Limpiar este filtro ──
  panel.querySelector('#txDpClear').addEventListener('click', () => {
    delete txColFilters[colKey];
    closeTxDropdown();
    refreshTxTable();
  });

  // ── Posicionamiento viewport-relativo (position:fixed → sin scrollY) ──
  const rect   = anchorBtn.getBoundingClientRect();
  const panelW = 280;
  const left   = Math.min(rect.left, window.innerWidth - panelW - 8);
  panel.style.top  = `${rect.bottom + 4}px`;
  panel.style.left = `${left}px`;

  // Ajustar si se sale por abajo
  requestAnimationFrame(() => {
    const ph = panel.offsetHeight;
    if (rect.bottom + ph > window.innerHeight - 8) {
      panel.style.top = `${rect.top - ph - 4}px`;
    }
  });

  // Cerrar al hacer clic fuera
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

/** Limpia todos los filtros de columna activos (incluye año) */
function clearAllTxFilters() {
  txColFilters = {};
  refreshTxTable();
  syncYearBtnState();
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

  // Serial de Excel guardado como string numérico (ej: "42739" → 4-Ene-2017)
  // Rango: 1 (1-Ene-1900) a 2958465 (31-Dic-9999)
  if (/^\d{4,6}$/.test(str)) {
    const serial = parseInt(str, 10);
    if (serial > 1 && serial < 2958465) {
      const ms  = Math.round((serial - 25569) * 86400 * 1000);
      const utc = new Date(ms);
      return new Date(utc.getUTCFullYear(), utc.getUTCMonth(), utc.getUTCDate());
    }
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
  if (!d) return 'Sin fecha';
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
        // Para barras horizontales (sin label de dataset) se usa el label de categoría
        : ctx => {
            const val  = ctx.dataset.data[ctx.dataIndex];
            const lbl  = ctx.dataset.label || ctx.label || '';
            return ` ${lbl ? lbl + ': ' : ''}${labelFn(val)}`;
          },
    },
  };
}
