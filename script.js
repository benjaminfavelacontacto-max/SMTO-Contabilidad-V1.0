/* ══════════════════════════════════════════════════════
   FinDash — script.js
   Lógica principal: lectura de Excel, procesamiento,
   generación de gráficas y tabla de datos.
══════════════════════════════════════════════════════ */

'use strict';

// ─── ESTADO GLOBAL ────────────────────────────────────
let allRows       = [];   // Todos los registros del año actual
let filteredRows  = [];   // Registros filtrados (por mes)
let barChartInst  = null; // Instancia de Chart.js (barras)
let donutChartInst= null; // Instancia de Chart.js (dona)
const CURRENT_YEAR = new Date().getFullYear();

// ─── PALETA DE COLORES PARA CATEGORÍAS ───────────────
const PALETTE = [
  '#6366f1','#8b5cf6','#ec4899','#f43f5e',
  '#f97316','#f59e0b','#10b981','#14b8a6',
  '#06b6d4','#3b82f6','#a3e635','#84cc16',
  '#e879f9','#fb7185','#fbbf24','#34d399',
];

// Meses en español
const MONTHS_ES = [
  'Ene','Feb','Mar','Abr','May','Jun',
  'Jul','Ago','Sep','Oct','Nov','Dic'
];
const MONTHS_LONG = [
  'Enero','Febrero','Marzo','Abril','Mayo','Junio',
  'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'
];

// ══════════════════════════════════════════════════════
// 1. DRAG & DROP / FILE INPUT
// ══════════════════════════════════════════════════════

function triggerFileInput() {
  document.getElementById('fileInput').click();
}

function handleDragOver(e) {
  e.preventDefault();
  e.stopPropagation();
  document.getElementById('dropZone').classList.add('drag-over');
}

function handleDragLeave(e) {
  e.preventDefault();
  document.getElementById('dropZone').classList.remove('drag-over');
}

function handleDrop(e) {
  e.preventDefault();
  e.stopPropagation();
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
  // Validar extensión
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

      // Tomar la primera hoja
      const sheetName = wb.SheetNames[0];
      const ws        = wb.Sheets[sheetName];

      // Convertir a JSON (primera fila como headers)
      const raw = XLSX.utils.sheet_to_json(ws, { defval: null, raw: false });

      if (!raw || raw.length === 0) {
        throw new Error('La hoja de cálculo está vacía o no tiene datos válidos.');
      }

      // Normalizar columnas y limpiar datos
      const rows = normalizeRows(raw);

      if (rows.length === 0) {
        throw new Error(
          `No se encontraron transacciones del año ${CURRENT_YEAR}. ` +
          'Verifica que el Excel tenga columnas de Fecha, Tipo y Monto.'
        );
      }

      allRows = rows;
      buildDashboard(rows);
    } catch (err) {
      hideLoading();
      showError(err.message || 'Error al leer el archivo. Verifica que sea un Excel válido.');
    }
  };

  reader.onerror = () => {
    hideLoading();
    showError('No se pudo leer el archivo. Intenta de nuevo.');
  };

  reader.readAsArrayBuffer(file);
}

// ══════════════════════════════════════════════════════
// 3. NORMALIZACIÓN DE DATOS
// ══════════════════════════════════════════════════════

/**
 * Mapeo flexible de nombres de columnas.
 * Detecta variantes en mayúsculas/minúsculas y otros idiomas.
 */
const COL_MAP = {
  fecha:        ['fecha','date','fec','f.','dia','day','periodo','period'],
  tipo:         ['tipo','type','clase','class','categoria_tipo','movimiento','mov','naturaleza'],
  categoria:    ['categoria','categoría','category','cat','rubro','concepto','concept','descripcion','description'],
  subcategoria: ['subcategoria','subcategoría','subcategory','subcat','subrubro','detalle'],
  monto:        ['monto','amount','valor','value','importe','total','cantidad','sum','precio'],
};

function findCol(headers, aliases) {
  const hLow = headers.map(h => String(h).toLowerCase().trim());
  for (const alias of aliases) {
    const idx = hLow.indexOf(alias);
    if (idx !== -1) return headers[idx];
  }
  return null;
}

function normalizeRows(raw) {
  if (!raw.length) return [];
  const headers = Object.keys(raw[0]);

  // Detectar columnas
  const colFecha    = findCol(headers, COL_MAP.fecha);
  const colTipo     = findCol(headers, COL_MAP.tipo);
  const colCat      = findCol(headers, COL_MAP.categoria);
  const colSubcat   = findCol(headers, COL_MAP.subcategoria);
  const colMonto    = findCol(headers, COL_MAP.monto);

  // Columnas mínimas requeridas
  if (!colFecha) throw new Error('No se encontró columna de Fecha. Verifica los encabezados del Excel.');
  if (!colMonto) throw new Error('No se encontró columna de Monto. Verifica los encabezados del Excel.');

  const rows = [];

  for (const row of raw) {
    try {
      // ── FECHA ──
      const rawFecha = row[colFecha];
      const fecha = parseDate(rawFecha);
      if (!fecha) continue;
      if (fecha.getFullYear() !== CURRENT_YEAR) continue;

      // ── MONTO ──
      const rawMonto = row[colMonto];
      const monto = parseMonto(rawMonto);
      if (isNaN(monto) || monto === 0) continue;

      // ── TIPO ──
      let tipo = 'Egreso'; // default
      if (colTipo && row[colTipo]) {
        const rawTipo = String(row[colTipo]).toLowerCase().trim();
        if (
          rawTipo.includes('ingreso') || rawTipo.includes('income') ||
          rawTipo.includes('entrada') || rawTipo.includes('credit') ||
          rawTipo.includes('crédito') || rawTipo === 'in' || rawTipo === '+'
        ) {
          tipo = 'Ingreso';
        } else if (
          rawTipo.includes('egreso') || rawTipo.includes('gasto') ||
          rawTipo.includes('expense') || rawTipo.includes('salida') ||
          rawTipo.includes('debit') || rawTipo.includes('débito') ||
          rawTipo === 'out' || rawTipo === '-'
        ) {
          tipo = 'Egreso';
        } else {
          // Si el monto es negativo → egreso, positivo → ingreso
          tipo = monto >= 0 ? 'Ingreso' : 'Egreso';
        }
      } else {
        // Sin columna tipo: inferir por signo del monto
        tipo = monto >= 0 ? 'Ingreso' : 'Egreso';
      }

      // ── CATEGORÍA ──
      const cat    = colCat    && row[colCat]    ? String(row[colCat]).trim()    : 'Sin categoría';
      const subcat = colSubcat && row[colSubcat] ? String(row[colSubcat]).trim() : '—';

      rows.push({
        fecha,
        mes:   fecha.getMonth(),       // 0-11
        tipo,
        categoria:    capitalizar(cat),
        subcategoria: capitalizar(subcat),
        monto:        Math.abs(monto), // Siempre positivo
      });
    } catch (_) {
      // Fila con error: ignorar y continuar
    }
  }

  return rows;
}

// ── Parsear fecha (Excel serial o string) ──
function parseDate(val) {
  if (!val) return null;

  // Ya es Date (SheetJS con cellDates:true)
  if (val instanceof Date) {
    return isNaN(val.getTime()) ? null : val;
  }

  // Número → serial de Excel
  if (typeof val === 'number') {
    const d = XLSX.SSF.parse_date_code(val);
    if (d) return new Date(d.y, d.m - 1, d.d);
    return null;
  }

  // String
  const str = String(val).trim();
  if (!str) return null;

  // Intentar parse directo
  const direct = new Date(str);
  if (!isNaN(direct.getTime())) return direct;

  // Formato DD/MM/YYYY
  const ddmmyyyy = str.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})$/);
  if (ddmmyyyy) {
    const d = new Date(+ddmmyyyy[3], +ddmmyyyy[2] - 1, +ddmmyyyy[1]);
    if (!isNaN(d.getTime())) return d;
  }

  // Formato YYYY/MM/DD
  const yyyymmdd = str.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})$/);
  if (yyyymmdd) {
    const d = new Date(+yyyymmdd[1], +yyyymmdd[2] - 1, +yyyymmdd[3]);
    if (!isNaN(d.getTime())) return d;
  }

  return null;
}

// ── Parsear monto (quitar $, comas, espacios) ──
function parseMonto(val) {
  if (val === null || val === undefined || val === '') return NaN;
  if (typeof val === 'number') return val;
  const cleaned = String(val).replace(/[\$\s,]/g, '').replace(/[()]/g, m => m === '(' ? '-' : '');
  return parseFloat(cleaned);
}

// ── Capitalizar primera letra ──
function capitalizar(str) {
  if (!str || str === '—') return str;
  return str.charAt(0).toUpperCase() + str.slice(1);
}

// ══════════════════════════════════════════════════════
// 4. CONSTRUCCIÓN DEL DASHBOARD
// ══════════════════════════════════════════════════════

function buildDashboard(rows) {
  filteredRows = rows;

  // KPIs
  renderKPIs(rows);

  // Gráficas
  renderBarChart(rows);
  renderDonutChart(rows);

  // Tablas
  renderCategoryTable(rows);
  renderTxTable(rows);

  // Filtro por mes
  buildMonthFilter(rows);

  // Mostrar dashboard
  hideLoading();
  document.getElementById('uploadSection').classList.add('hidden');
  document.getElementById('dashboardSection').classList.remove('hidden');
  document.getElementById('dashSubtitle').textContent = `Año ${CURRENT_YEAR}`;
  document.getElementById('monthFilterWrapper').classList.remove('hidden');
  document.getElementById('exportCsvBtn').classList.remove('hidden');

  // Scroll suave al top
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ── KPIs ──
function renderKPIs(rows) {
  const ingresos = rows.filter(r => r.tipo === 'Ingreso');
  const egresos  = rows.filter(r => r.tipo === 'Egreso');

  const totalIngresos = ingresos.reduce((s, r) => s + r.monto, 0);
  const totalEgresos  = egresos.reduce((s, r)  => s + r.monto, 0);
  const balance       = totalIngresos - totalEgresos;
  const tasa          = totalIngresos > 0 ? ((balance / totalIngresos) * 100) : 0;

  document.getElementById('kpiIncome').textContent        = formatMoney(totalIngresos);
  document.getElementById('kpiIncomeDetail').textContent  = `${ingresos.length} transacciones`;
  document.getElementById('kpiExpense').textContent       = formatMoney(totalEgresos);
  document.getElementById('kpiExpenseDetail').textContent = `${egresos.length} transacciones`;
  document.getElementById('kpiBalance').textContent       = formatMoney(balance);
  document.getElementById('kpiBalanceDetail').textContent = balance >= 0 ? '✓ Balance positivo' : '⚠ Balance negativo';
  document.getElementById('kpiRate').textContent          = `${tasa.toFixed(1)}%`;

  // Color del balance
  const balanceCard = document.getElementById('kpiBalance').closest('.kpi-card');
  balanceCard.style.setProperty('--bal-color', balance >= 0 ? 'var(--income)' : 'var(--expense)');
  document.getElementById('kpiBalance').style.color = balance >= 0 ? 'var(--income)' : 'var(--expense)';
}

// ── Gráfica de barras ──
function renderBarChart(rows) {
  const meses    = Array.from({ length: 12 }, () => ({ ingreso: 0, egreso: 0 }));
  const mesConDatos = new Set();

  for (const r of rows) {
    meses[r.mes][r.tipo === 'Ingreso' ? 'ingreso' : 'egreso'] += r.monto;
    mesConDatos.add(r.mes);
  }

  // Solo mostrar meses con datos
  const labels = [...mesConDatos].sort((a, b) => a - b).map(m => MONTHS_ES[m]);
  const incomeData  = [...mesConDatos].sort((a,b)=>a-b).map(m => meses[m].ingreso);
  const expenseData = [...mesConDatos].sort((a,b)=>a-b).map(m => meses[m].egreso);

  const isDark = document.documentElement.getAttribute('data-theme') !== 'light';
  const gridColor = isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.06)';
  const textColor = isDark ? '#64748b' : '#94a3b8';

  const ctx = document.getElementById('barChart').getContext('2d');

  if (barChartInst) barChartInst.destroy();

  barChartInst = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        {
          label: 'Ingresos',
          data: incomeData,
          backgroundColor: 'rgba(16,185,129,0.75)',
          borderColor:     'rgba(16,185,129,1)',
          borderWidth: 0,
          borderRadius: 6,
          borderSkipped: false,
        },
        {
          label: 'Egresos',
          data: expenseData,
          backgroundColor: 'rgba(244,63,94,0.75)',
          borderColor:     'rgba(244,63,94,1)',
          borderWidth: 0,
          borderRadius: 6,
          borderSkipped: false,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: isDark ? '#1e293b' : '#fff',
          titleColor: isDark ? '#f1f5f9' : '#0f172a',
          bodyColor:  isDark ? '#94a3b8' : '#475569',
          borderColor: isDark ? 'rgba(255,255,255,0.1)' : 'rgba(0,0,0,0.1)',
          borderWidth: 1,
          padding: 12,
          cornerRadius: 10,
          callbacks: {
            label: ctx => ` ${ctx.dataset.label}: ${formatMoney(ctx.parsed.y)}`,
          },
        },
      },
      scales: {
        x: {
          grid: { display: false },
          ticks: { color: textColor, font: { family: 'Inter', size: 12 } },
        },
        y: {
          grid: { color: gridColor },
          ticks: {
            color: textColor,
            font: { family: 'Inter', size: 12 },
            callback: v => formatMoneyShort(v),
          },
          border: { display: false },
        },
      },
    },
  });
}

// ── Gráfica dona ──
function renderDonutChart(rows) {
  const egresos = rows.filter(r => r.tipo === 'Egreso');
  const bycat   = agruparPorCategoria(egresos);
  const sorted  = Object.entries(bycat).sort((a,b) => b[1] - a[1]).slice(0, 10);

  const labels  = sorted.map(([cat]) => cat);
  const data    = sorted.map(([,val]) => val);
  const colors  = labels.map((_, i) => PALETTE[i % PALETTE.length]);

  const totalEgresos = egresos.reduce((s, r) => s + r.monto, 0);
  document.getElementById('donutTotal').textContent = formatMoney(totalEgresos);

  const isDark = document.documentElement.getAttribute('data-theme') !== 'light';
  const ctx = document.getElementById('donutChart').getContext('2d');

  if (donutChartInst) donutChartInst.destroy();

  donutChartInst = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels,
      datasets: [{
        data,
        backgroundColor: colors,
        borderColor:     isDark ? '#1e293b' : '#fff',
        borderWidth: 3,
        hoverOffset: 6,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: '72%',
      plugins: {
        legend: {
          position: 'bottom',
          labels: {
            color: isDark ? '#94a3b8' : '#475569',
            font: { family: 'Inter', size: 11 },
            padding: 12,
            boxWidth: 10,
            boxHeight: 10,
            usePointStyle: true,
            pointStyleWidth: 10,
          },
        },
        tooltip: {
          backgroundColor: isDark ? '#1e293b' : '#fff',
          titleColor: isDark ? '#f1f5f9' : '#0f172a',
          bodyColor:  isDark ? '#94a3b8' : '#475569',
          borderColor: isDark ? 'rgba(255,255,255,0.1)' : 'rgba(0,0,0,0.1)',
          borderWidth: 1,
          padding: 12,
          cornerRadius: 10,
          callbacks: {
            label: ctx => {
              const pct = totalEgresos > 0
                ? ((ctx.parsed / totalEgresos) * 100).toFixed(1)
                : '0.0';
              return ` ${ctx.label}: ${formatMoney(ctx.parsed)} (${pct}%)`;
            },
          },
        },
      },
    },
  });
}

// ── Tabla de categorías ──
function renderCategoryTable(rows) {
  const egresos      = rows.filter(r => r.tipo === 'Egreso');
  const ingresos     = rows.filter(r => r.tipo === 'Ingreso');
  const totalEgresos = egresos.reduce((s,r) => s + r.monto, 0);

  const byCatE = agruparPorCategoria(egresos);
  const byCatI = agruparPorCategoria(ingresos);
  const contE  = contarPorCategoria(egresos);
  const contI  = contarPorCategoria(ingresos);

  // Combinar todas las categorías
  const allCats = new Set([...Object.keys(byCatE), ...Object.keys(byCatI)]);
  const entries = [];

  for (const cat of allCats) {
    if (byCatE[cat]) entries.push({ cat, tipo: 'Egreso',  total: byCatE[cat], count: contE[cat] || 0 });
    if (byCatI[cat]) entries.push({ cat, tipo: 'Ingreso', total: byCatI[cat], count: contI[cat] || 0 });
  }

  // Ordenar por total descendente
  entries.sort((a,b) => b.total - a.total);

  const tbody = document.getElementById('categoryTableBody');
  tbody.innerHTML = '';

  for (const e of entries) {
    const pct = (e.tipo === 'Egreso' && totalEgresos > 0)
      ? ((e.total / totalEgresos) * 100)
      : null;

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong style="color:var(--text-1)">${escHtml(e.cat)}</strong></td>
      <td><span class="type-badge ${e.tipo === 'Ingreso' ? 'type-income' : 'type-expense'}">${e.tipo}</span></td>
      <td class="text-right" style="color:var(--text-1);font-weight:600">${formatMoney(e.total)}</td>
      <td class="text-right">
        ${pct !== null
          ? `<div class="progress-cell">
               <span>${pct.toFixed(1)}%</span>
               <div class="progress-bar">
                 <div class="progress-fill" style="width:${Math.min(pct,100)}%"></div>
               </div>
             </div>`
          : '<span style="color:var(--text-3)">—</span>'
        }
      </td>
      <td class="text-right">${e.count}</td>
    `;
    tbody.appendChild(tr);
  }

  document.getElementById('tableBadge').textContent = `${entries.length} categorías`;
}

// ── Tabla de últimas transacciones (máx. 100) ──
function renderTxTable(rows) {
  const sorted = [...rows]
    .sort((a,b) => b.fecha - a.fecha)
    .slice(0, 100);

  const tbody = document.getElementById('txTableBody');
  tbody.innerHTML = '';

  for (const r of sorted) {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${formatDate(r.fecha)}</td>
      <td><span class="type-badge ${r.tipo === 'Ingreso' ? 'type-income' : 'type-expense'}">${r.tipo}</span></td>
      <td>${escHtml(r.categoria)}</td>
      <td style="color:var(--text-3)">${escHtml(r.subcategoria)}</td>
      <td class="text-right" style="color:${r.tipo==='Ingreso'?'var(--income)':'var(--expense)'};font-weight:600">
        ${r.tipo==='Ingreso'?'+':'-'}${formatMoney(r.monto)}
      </td>
    `;
    tbody.appendChild(tr);
  }

  document.getElementById('txBadge').textContent = `${sorted.length} registros`;
}

// ══════════════════════════════════════════════════════
// 5. FILTRO POR MES
// ══════════════════════════════════════════════════════

function buildMonthFilter(rows) {
  const meses = [...new Set(rows.map(r => r.mes))].sort((a,b)=>a-b);
  const sel   = document.getElementById('monthFilter');

  // Limpiar opciones previas (excepto "Todos")
  while (sel.options.length > 1) sel.remove(1);

  for (const m of meses) {
    const opt = document.createElement('option');
    opt.value = m;
    opt.textContent = MONTHS_LONG[m];
    sel.appendChild(opt);
  }

  // Listener
  sel.onchange = () => applyMonthFilter(sel.value);
}

function applyMonthFilter(val) {
  filteredRows = val === 'all'
    ? allRows
    : allRows.filter(r => r.mes === parseInt(val, 10));

  renderKPIs(filteredRows);
  renderBarChart(filteredRows);
  renderDonutChart(filteredRows);
  renderCategoryTable(filteredRows);
  renderTxTable(filteredRows);

  const label = val === 'all' ? `Año ${CURRENT_YEAR}` : `${MONTHS_LONG[parseInt(val,10)]} ${CURRENT_YEAR}`;
  document.getElementById('dashSubtitle').textContent = label;
}

// ══════════════════════════════════════════════════════
// 6. EXPORTAR CSV
// ══════════════════════════════════════════════════════

function exportCSV() {
  const rows = filteredRows;
  const header = ['Fecha','Tipo','Categoría','Subcategoría','Monto'];
  const lines  = [header.join(',')];

  for (const r of rows) {
    lines.push([
      formatDate(r.fecha),
      r.tipo,
      `"${r.categoria}"`,
      `"${r.subcategoria}"`,
      r.monto.toFixed(2),
    ].join(','));
  }

  const blob = new Blob([lines.join('\n')], { type: 'text/csv;charset=utf-8;' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = `findash_${CURRENT_YEAR}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ══════════════════════════════════════════════════════
// 7. TEMA DARK / LIGHT
// ══════════════════════════════════════════════════════

function toggleTheme() {
  const html = document.documentElement;
  const isDark = html.getAttribute('data-theme') === 'dark';
  html.setAttribute('data-theme', isDark ? 'light' : 'dark');

  document.getElementById('iconSun').classList.toggle('hidden', !isDark);
  document.getElementById('iconMoon').classList.toggle('hidden', isDark);

  // Re-renderizar gráficas con colores actualizados
  if (allRows.length > 0) {
    renderBarChart(filteredRows);
    renderDonutChart(filteredRows);
  }
}

// ══════════════════════════════════════════════════════
// 8. RESET / NUEVO ARCHIVO
// ══════════════════════════════════════════════════════

function resetApp() {
  allRows = [];
  filteredRows = [];

  // Ocultar dashboard
  document.getElementById('dashboardSection').classList.add('hidden');
  document.getElementById('monthFilterWrapper').classList.add('hidden');
  document.getElementById('exportCsvBtn').classList.add('hidden');

  // Mostrar upload
  document.getElementById('uploadSection').classList.remove('hidden');
  document.getElementById('monthFilter').value = 'all';

  // Reset input file
  document.getElementById('fileInput').value = '';

  // Destruir gráficas
  if (barChartInst)   { barChartInst.destroy();   barChartInst   = null; }
  if (donutChartInst) { donutChartInst.destroy(); donutChartInst = null; }

  hideError();
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ══════════════════════════════════════════════════════
// 9. HELPERS UI
// ══════════════════════════════════════════════════════

function showLoading() {
  document.getElementById('uploadSection').classList.add('hidden');
  document.getElementById('loadingSection').classList.remove('hidden');
}
function hideLoading() {
  document.getElementById('loadingSection').classList.add('hidden');
}
function showError(msg) {
  const el = document.getElementById('errorMsg');
  document.getElementById('errorText').textContent = msg;
  el.classList.remove('hidden');
  document.getElementById('uploadSection').classList.remove('hidden');
  hideLoading();
}
function hideError() {
  document.getElementById('errorMsg').classList.add('hidden');
}

// ══════════════════════════════════════════════════════
// 10. HELPERS DE DATOS
// ══════════════════════════════════════════════════════

function agruparPorCategoria(rows) {
  return rows.reduce((acc, r) => {
    acc[r.categoria] = (acc[r.categoria] || 0) + r.monto;
    return acc;
  }, {});
}

function contarPorCategoria(rows) {
  return rows.reduce((acc, r) => {
    acc[r.categoria] = (acc[r.categoria] || 0) + 1;
    return acc;
  }, {});
}

// ── Formatear dinero ──
function formatMoney(n) {
  return new Intl.NumberFormat('es-MX', {
    style: 'currency',
    currency: 'MXN',
    minimumFractionDigits: 0,
    maximumFractionDigits: 0,
  }).format(n);
}

// ── Formatear dinero corto para ejes (10K, 1M) ──
function formatMoneyShort(n) {
  if (Math.abs(n) >= 1_000_000) return `$${(n/1_000_000).toFixed(1)}M`;
  if (Math.abs(n) >= 1_000)     return `$${(n/1_000).toFixed(0)}K`;
  return `$${n}`;
}

// ── Formatear fecha ──
function formatDate(d) {
  return d.toLocaleDateString('es-MX', {
    day:   '2-digit',
    month: 'short',
    year:  'numeric',
  });
}

// ── Escape HTML ──
function escHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;');
}
