/* ============================================================
   RE-REGISTRATION DASHBOARD — MAIN JAVASCRIPT
   ============================================================ */

'use strict';

// ────────────────────────────────────────────────────────────
// STATE
// ────────────────────────────────────────────────────────────
let rawData       = [];   // all rows parsed from Excel
let filteredData  = [];   // rows after filters
let allColumns    = [];   // detected column headers
let columnMap     = {};   // normalized key → original header
let charts        = {};   // Chart.js instances
let currentPage   = 1;
const PAGE_SIZE   = 25;

// ────────────────────────────────────────────────────────────
// COLUMN KEY DETECTION (fuzzy matching)
// ────────────────────────────────────────────────────────────
const COL_PATTERNS = {
  batch:       [/batch/i, /cohort/i],
  program:     [/program/i, /programme/i, /course/i, /degree/i, /dept/i, /department/i, /stream/i, /specialization/i],
  studentId:   [/student.?id/i, /enroll/i, /roll/i, /reg.?no/i, /registration.?no/i],
  studentName: [/student.?name/i, /name/i, /full.?name/i],
  reRegStatus: [/re.?reg.*status/i, /re-reg/i, /rereg/i, /re_reg/i, /re.reg/i, /re.?registration.?status/i],
  reRegDone:   [/re.?reg.*done/i, /re.?reg.*complete/i, /registered/i],
  iaStatus:    [/ia.*status/i, /internal.*status/i, /ia/i, /internal.?assessment/i],
  iaSubmit:    [/ia.*submit/i, /ia.*done/i, /internal.*submit/i],
  iaPct:       [/ia.*percent/i, /ia.*pct/i, /ia.*\%/i, /internal.*percent/i],
  sem:         [/sem(ester)?/i, /semester/i],
  gender:      [/gender/i, /sex/i],
  mobile:      [/mobile/i, /phone/i, /contact/i],
  email:       [/email/i, /mail/i],
  date:        [/date/i, /timestamp/i],
  center:      [/center/i, /centre/i, /campus/i, /city/i, /location/i],
  iac:         [/IAC/i, /iac/i]
};

function detectColumn(headers, key) {
  const patterns = COL_PATTERNS[key] || [];
  for (const h of headers) {
    for (const p of patterns) {
      if (p.test(h)) return h;
    }
  }
  return null;
}

function normalizeVal(val) {
  if (val === null || val === undefined) return '';
  return String(val).trim();
}

// ────────────────────────────────────────────────────────────
// FILE HANDLING
// ────────────────────────────────────────────────────────────
document.getElementById('fileInput').addEventListener('change', handleFile);

// Drag & Drop
const dropArea = document.getElementById('dropArea');
const uploadCard = document.querySelector('.upload-card');
['dragenter','dragover'].forEach(e => {
  uploadCard.addEventListener(e, ev => { ev.preventDefault(); uploadCard.classList.add('drag-over'); });
});
['dragleave','drop'].forEach(e => {
  uploadCard.addEventListener(e, ev => { ev.preventDefault(); uploadCard.classList.remove('drag-over'); });
});
uploadCard.addEventListener('drop', ev => {
  const file = ev.dataTransfer.files[0];
  if (file) processFile(file);
});

function handleFile(e) {
  const file = e.target.files[0];
  if (file) processFile(file);
}

function processFile(file) {
  showLoading(true);
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const data   = new Uint8Array(e.target.result);
      const wb     = XLSX.read(data, { type: 'array', cellDates: true });
      parseWorkbook(wb, file.name);
    } catch(err) {
      showLoading(false);
      alert('Error reading file: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

// ────────────────────────────────────────────────────────────
// WORKBOOK PARSING
// ────────────────────────────────────────────────────────────
function parseWorkbook(wb, fileName) {
  rawData = [];

  // Try each sheet — pick the one with the most rows
  let bestSheet = null, bestRows = 0;
  wb.SheetNames.forEach(name => {
    const sheet = wb.Sheets[name];
    const json  = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    if (json.length > bestRows) { bestRows = json.length; bestSheet = sheet; }
  });

  if (!bestSheet) { alert('No data found in the Excel file.'); showLoading(false); return; }

  const json = XLSX.utils.sheet_to_json(bestSheet, { defval: '' });
  if (!json.length) { alert('Sheet appears empty.'); showLoading(false); return; }

  allColumns  = Object.keys(json[0]);
  rawData     = json;

  // Build columnMap
  columnMap = {};
  Object.keys(COL_PATTERNS).forEach(key => {
    columnMap[key] = detectColumn(allColumns, key);
  });

  // Update last-updated
  const lu = document.getElementById('lastUpdated');
  lu.classList.remove('hidden');
  lu.querySelector('span').textContent = 'Loaded: ' + fileName + ' · ' + rawData.length + ' rows';

  // Populate filters
  populateFilters();

  // Initial render
  filteredData = [...rawData];
  renderDashboard();

  showLoading(false);
  document.getElementById('uploadZone').classList.add('hidden');
  document.getElementById('dashboard').classList.remove('hidden');
}

// ────────────────────────────────────────────────────────────
// HELPERS
// ────────────────────────────────────────────────────────────
function getVal(row, key) {
  const col = columnMap[key];
  return col ? normalizeVal(row[col]) : '';
}

function isReRegDone(row) {
  const status = getVal(row, 'reRegStatus').toLowerCase();
  const done   = getVal(row, 'reRegDone').toLowerCase();
  const combined = status + done;
  return /\b(yes|done|complete|completed|registered|success|1|true|rereg)\b/.test(combined);
}

function isIASubmitted(row) {
  const status = getVal(row, 'iaStatus').toLowerCase();
  const sub    = getVal(row, 'iaSubmit').toLowerCase();
  const combined = status + sub;
  return /\b(yes|done|submit|submitted|complete|1|true)\b/.test(combined);
}

function pct(num, den) {
  if (!den) return 0;
  return Math.round((num / den) * 100 * 10) / 10;
}

function pctClass(p) {
  if (p >= 75) return 'pct-high';
  if (p >= 50) return 'pct-med';
  return 'pct-low';
}

// ────────────────────────────────────────────────────────────
// FILTER POPULATION
// ────────────────────────────────────────────────────────────
function populateFilters() {
  const batchSel   = document.getElementById('filterBatch');
  const programSel = document.getElementById('filterProgram');

  const batches  = [...new Set(rawData.map(r => getVal(r, 'batch')).filter(Boolean))].sort();
  const programs = [...new Set(rawData.map(r => getVal(r, 'program')).filter(Boolean))].sort();

  batches.forEach(b  => { const o = document.createElement('option'); o.value = b; o.textContent = b; batchSel.appendChild(o); });
  programs.forEach(p => { const o = document.createElement('option'); o.value = p; o.textContent = p; programSel.appendChild(o); });
}

function applyFilters() {
  const batch   = document.getElementById('filterBatch').value;
  const program = document.getElementById('filterProgram').value;
  const status  = document.getElementById('filterStatus').value;
  const search  = document.getElementById('searchInput').value.toLowerCase();

  filteredData = rawData.filter(row => {
    if (batch   !== 'ALL' && getVal(row, 'batch')   !== batch)   return false;
    if (program !== 'ALL' && getVal(row, 'program')  !== program) return false;
    if (status === 'DONE'    && !isReRegDone(row))    return false;
    if (status === 'PENDING' &&  isReRegDone(row))    return false;
    if (search) {
      const rowStr = Object.values(row).join(' ').toLowerCase();
      if (!rowStr.includes(search)) return false;
    }
    return true;
  });

  currentPage = 1;
  renderDashboard();
}

function resetFilters() {
  document.getElementById('filterBatch').value   = 'ALL';
  document.getElementById('filterProgram').value = 'ALL';
  document.getElementById('filterStatus').value  = 'ALL';
  document.getElementById('searchInput').value   = '';
  filteredData = [...rawData];
  currentPage  = 1;
  renderDashboard();
}

['filterBatch','filterProgram','filterStatus'].forEach(id =>
  document.getElementById(id).addEventListener('change', applyFilters));
document.getElementById('searchInput').addEventListener('input', applyFilters);

// ────────────────────────────────────────────────────────────
// ANALYTICS COMPUTATION
// ────────────────────────────────────────────────────────────
function computeStats(data) {
  const total      = data.length;
  const reRegDone  = data.filter(isReRegDone).length;
  const reRegPend  = total - reRegDone;
  const reRegPct   = pct(reRegDone, total);

  const iaSubmitted = data.filter(isIASubmitted).length;
  const iaPct_val   = pct(iaSubmitted, total);

  // Batch-wise
  const batchMap = {};
  data.forEach(row => {
    const batch   = getVal(row, 'batch')   || 'Unknown';
    const program = getVal(row, 'program') || 'Unknown';
    const key     = batch;
    if (!batchMap[key]) batchMap[key] = { batch, program: new Set(), total: 0, done: 0, ia: 0 };
    batchMap[key].total++;
    batchMap[key].program.add(program);
    if (isReRegDone(row))    batchMap[key].done++;
    if (isIASubmitted(row))  batchMap[key].ia++;
  });

  const batchStats = Object.values(batchMap).map(b => ({
    batch:    b.batch,
    program:  [...b.program].join(', '),
    total:    b.total,
    done:     b.done,
    pending:  b.total - b.done,
    reRegPct: pct(b.done, b.total),
    ia:       b.ia,
    iaPct:    pct(b.ia, b.total),
  })).sort((a, b) => b.total - a.total);

  // Program-wise
  const progMap = {};
  data.forEach(row => {
    const prog = getVal(row, 'program') || 'Unknown';
    if (!progMap[prog]) progMap[prog] = { prog, total: 0, done: 0, ia: 0 };
    progMap[prog].total++;
    if (isReRegDone(row))   progMap[prog].done++;
    if (isIASubmitted(row)) progMap[prog].ia++;
  });
  const progStats = Object.values(progMap).sort((a, b) => b.total - a.total);

  // Center/City wise
  const centerMap = {};
  data.forEach(row => {
    const center = getVal(row, 'center') || null;
    if (!center) return;
    if (!centerMap[center]) centerMap[center] = { total: 0, done: 0 };
    centerMap[center].total++;
    if (isReRegDone(row)) centerMap[center].done++;
  });

  return { total, reRegDone, reRegPend, reRegPct, iaSubmitted, iaPct: iaPct_val, batchStats, progStats, centerMap };
}

// ────────────────────────────────────────────────────────────
// RENDER DASHBOARD
// ────────────────────────────────────────────────────────────
function renderDashboard() {
  const stats = computeStats(filteredData);
  renderKPIs(stats);
  renderBatchTable(stats.batchStats);
  renderCharts(stats);
  renderProgramGrid(stats.progStats);
  renderDetailTable();
}

// ── KPI CARDS ──────────────────────────────────────────────
function renderKPIs(stats) {
  const grid = document.getElementById('kpiGrid');

  const cards = [
    {
      label: 'Total Students',
      value: stats.total.toLocaleString(),
      icon: 'fas fa-users',
      color: '#4f46e5', bg: '#e0e7ff',
      sub: 'Filtered dataset',
    },
    {
      label: 'Re-Reg Done',
      value: stats.reRegDone.toLocaleString(),
      icon: 'fas fa-check-circle',
      color: '#10b981', bg: '#d1fae5',
      sub: `${stats.reRegPct}% completed`,
      pct: stats.reRegPct,
    },
    {
      label: 'Re-Reg Pending',
      value: stats.reRegPend.toLocaleString(),
      icon: 'fas fa-clock',
      color: '#ef4444', bg: '#fee2e2',
      sub: `${pct(stats.reRegPend, stats.total)}% pending`,
      pct: pct(stats.reRegPend, stats.total),
    },
    {
      label: 'Re-Reg %',
      value: stats.reRegPct + '%',
      icon: 'fas fa-percentage',
      color: '#3b82f6', bg: '#dbeafe',
      sub: 'Completion rate',
      pct: stats.reRegPct,
    },
    {
      label: 'IA Submitted',
      value: stats.iaSubmitted.toLocaleString(),
      icon: 'fas fa-file-alt',
      color: '#8b5cf6', bg: '#ede9fe',
      sub: `${stats.iaPct}% IA done`,
      pct: stats.iaPct,
    },
    {
      label: 'IA %',
      value: stats.iaPct + '%',
      icon: 'fas fa-tasks',
      color: '#f97316', bg: '#ffedd5',
      sub: 'Internal Assessment rate',
      pct: stats.iaPct,
    },
    {
      label: 'Total Batches',
      value: stats.batchStats.length.toLocaleString(),
      icon: 'fas fa-layer-group',
      color: '#14b8a6', bg: '#ccfbf1',
      sub: 'Unique batches',
    },
    {
      label: 'Avg Re-Reg/Batch',
      value: stats.batchStats.length
        ? Math.round(stats.reRegDone / stats.batchStats.length).toLocaleString()
        : '0',
      icon: 'fas fa-chart-line',
      color: '#f59e0b', bg: '#fef3c7',
      sub: 'Average completions',
    },
  ];

  grid.innerHTML = cards.map(c => `
    <div class="kpi-card" style="--kpi-color:${c.color};--kpi-bg:${c.bg}">
      <div class="kpi-icon"><i class="${c.icon}"></i></div>
      <div class="kpi-label">${c.label}</div>
      <div class="kpi-value">${c.value}</div>
      <div class="kpi-sub">${c.sub}</div>
      ${c.pct !== undefined ? `
      <div class="progress-bar-wrap">
        <div class="progress-bar-bg"><div class="progress-bar-fill" style="width:${Math.min(c.pct,100)}%"></div></div>
        <div class="progress-label"><span>0%</span><span>${c.pct}%</span></div>
      </div>` : ''}
    </div>
  `).join('');
}

// ── BATCH SUMMARY TABLE ────────────────────────────────────
function renderBatchTable(batchStats) {
  const tbody = document.getElementById('batchSummaryBody');

  if (!batchStats.length) {
    tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:24px;color:var(--text-muted)">No batch data found</td></tr>';
    return;
  }

  tbody.innerHTML = batchStats.map(b => {
    const rp = b.reRegPct, ip = b.iaPct;
    const statusClass = rp >= 90 ? 'status-done' : rp >= 50 ? 'status-partial' : 'status-pending';
    const statusLabel = rp >= 90 ? '✅ Complete'   : rp >= 50 ? '⚠️ Partial'    : '❌ Behind';
    return `
      <tr>
        <td><strong>${b.batch}</strong></td>
        <td>${b.program}</td>
        <td><strong>${b.total}</strong></td>
        <td style="color:#10b981;font-weight:600">${b.done}</td>
        <td style="color:#ef4444">${b.pending}</td>
        <td><span class="pct-pill ${pctClass(rp)}">${rp}%</span></td>
        <td>${b.ia}</td>
        <td><span class="pct-pill ${pctClass(ip)}">${ip}%</span></td>
        <td><span class="status-pill ${statusClass}">${statusLabel}</span></td>
      </tr>`;
  }).join('');
}

// ── CHARTS ─────────────────────────────────────────────────
function destroyCharts() {
  Object.values(charts).forEach(c => { if (c) c.destroy(); });
  charts = {};
}

function renderCharts(stats) {
  destroyCharts();

  const bs = stats.batchStats;
  const labels = bs.map(b => b.batch.length > 18 ? b.batch.slice(0,16)+'…' : b.batch);

  // ── 1. Re-Reg % Bar Chart ──
  const ctx1 = document.getElementById('reRegBarChart').getContext('2d');
  charts.reRegBar = new Chart(ctx1, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Re-Reg %',
        data: bs.map(b => b.reRegPct),
        backgroundColor: bs.map(b =>
          b.reRegPct >= 75 ? '#10b981cc' : b.reRegPct >= 50 ? '#f59e0bcc' : '#ef4444cc'),
        borderRadius: 6,
        borderSkipped: false,
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => `${ctx.raw}%` } }
      },
      scales: {
        y: { beginAtZero: true, max: 100, ticks: { callback: v => v + '%' }, grid: { color: '#f1f5f9' } },
        x: { ticks: { font: { size: 11 }, maxRotation: 45 }, grid: { display: false } }
      }
    }
  });

  // ── 2. Overall Pie Chart ──
  const ctx2 = document.getElementById('reRegPieChart').getContext('2d');
  charts.reRegPie = new Chart(ctx2, {
    type: 'doughnut',
    data: {
      labels: ['Re-Reg Done', 'Pending'],
      datasets: [{
        data: [stats.reRegDone, stats.reRegPend],
        backgroundColor: ['#10b981', '#ef4444'],
        borderWidth: 3,
        borderColor: '#fff',
        hoverOffset: 8,
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      cutout: '65%',
      plugins: {
        legend: { position: 'bottom', labels: { padding: 16, font: { size: 13 } } },
        tooltip: {
          callbacks: {
            label: ctx => {
              const val = ctx.raw;
              const total = ctx.chart.data.datasets[0].data.reduce((a,b) => a+b, 0);
              return ` ${val} (${pct(val, total)}%)`;
            }
          }
        }
      }
    }
  });

  // ── 3. IA % Bar Chart ──
  const ctx3 = document.getElementById('iaBarChart').getContext('2d');
  charts.iaBar = new Chart(ctx3, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        {
          label: 'IA %',
          data: bs.map(b => b.iaPct),
          backgroundColor: '#8b5cf6cc',
          borderRadius: 6,
          borderSkipped: false,
        },
        {
          label: 'Re-Reg %',
          data: bs.map(b => b.reRegPct),
          backgroundColor: '#3b82f6cc',
          borderRadius: 6,
          borderSkipped: false,
        }
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position: 'top', labels: { font: { size: 12 } } },
        tooltip: { callbacks: { label: ctx => `${ctx.dataset.label}: ${ctx.raw}%` } }
      },
      scales: {
        y: { beginAtZero: true, max: 100, ticks: { callback: v => v + '%' }, grid: { color: '#f1f5f9' } },
        x: { ticks: { font: { size: 11 }, maxRotation: 45 }, grid: { display: false } }
      }
    }
  });

  // ── 4. Stacked Bar: Done vs Pending ──
  const ctx4 = document.getElementById('stackedBarChart').getContext('2d');
  charts.stacked = new Chart(ctx4, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        {
          label: 'Re-Reg Done',
          data: bs.map(b => b.done),
          backgroundColor: '#10b981cc',
          borderRadius: 4,
        },
        {
          label: 'Pending',
          data: bs.map(b => b.pending),
          backgroundColor: '#ef4444cc',
          borderRadius: 4,
        }
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position: 'top' },
      },
      scales: {
        x: { stacked: true, ticks: { font: { size: 11 }, maxRotation: 45 }, grid: { display: false } },
        y: { stacked: true, grid: { color: '#f1f5f9' } }
      }
    }
  });
}

// ── PROGRAM GRID ───────────────────────────────────────────
function renderProgramGrid(progStats) {
  const grid = document.getElementById('programGrid');
  if (!progStats.length) { grid.innerHTML = '<p style="color:var(--text-muted)">No program data detected.</p>'; return; }

  const colors = ['#4f46e5','#10b981','#f59e0b','#ef4444','#8b5cf6','#3b82f6','#14b8a6','#f97316'];

  grid.innerHTML = progStats.map((p, i) => {
    const rp = pct(p.done, p.total);
    const ip = pct(p.ia,   p.total);
    const col = colors[i % colors.length];
    return `
      <div class="kpi-card program-card" style="--kpi-color:${col};--kpi-bg:${col}22">
        <div class="kpi-icon"><i class="fas fa-university"></i></div>
        <div class="program-name">${p.prog}</div>
        <div class="program-stat"><span>Total Students</span><span><strong>${p.total}</strong></span></div>
        <div class="program-stat"><span>Re-Reg Done</span><span style="color:#10b981"><strong>${p.done}</strong></span></div>
        <div class="program-stat"><span>Pending</span><span style="color:#ef4444">${p.total - p.done}</span></div>
        <div class="progress-bar-wrap" style="--kpi-color:${col}">
          <div class="progress-bar-bg"><div class="progress-bar-fill" style="width:${rp}%"></div></div>
          <div class="progress-label"><span>Re-Reg: ${rp}%</span><span>IA: ${ip}%</span></div>
        </div>
      </div>`;
  }).join('');
}

// ── DETAIL TABLE ───────────────────────────────────────────
function renderDetailTable() {
  const head = document.getElementById('detailTableHead');
  const body = document.getElementById('detailTableBody');
  const meta = document.getElementById('tableMeta');

  const start = (currentPage - 1) * PAGE_SIZE;
  const end   = start + PAGE_SIZE;
  const page  = filteredData.slice(start, end);

  meta.textContent = `Showing ${start + 1}–${Math.min(end, filteredData.length)} of ${filteredData.length} records`;

  // Determine columns to show (limit to 15 most informative)
  let showCols = allColumns;
  if (showCols.length > 15) {
    // Prioritize detected columns
    const priority = ['studentId','studentName','batch','program','reRegStatus','reRegDone','iaStatus','iaSubmit','sem','center','gender'].map(k => columnMap[k]).filter(Boolean);
    const rest = allColumns.filter(c => !priority.includes(c)).slice(0, 15 - priority.length);
    showCols = [...new Set([...priority, ...rest])];
  }

  head.innerHTML = '<tr>' + showCols.map(c => `<th>${c}</th>`).join('') + '</tr>';

  body.innerHTML = page.map(row => {
    return '<tr>' + showCols.map(col => {
      let val = normalizeVal(row[col]);
      // Highlight re-reg & IA status
      if (col === columnMap['reRegStatus'] || col === columnMap['reRegDone']) {
        const isDone = /\b(yes|done|complete|completed|registered|success|1|true)\b/i.test(val);
        return `<td><span class="status-pill ${isDone ? 'status-done' : 'status-pending'}">${val || 'N/A'}</span></td>`;
      }
      if (col === columnMap['iaStatus'] || col === columnMap['iaSubmit']) {
        const isDone = /\b(yes|done|submit|submitted|complete|1|true)\b/i.test(val);
        return `<td><span class="status-pill ${isDone ? 'status-done' : 'status-pending'}">${val || 'N/A'}</span></td>`;
      }
      return `<td>${val || '—'}</td>`;
    }).join('') + '</tr>';
  }).join('');

  renderPagination(filteredData.length);
}

function renderPagination(total) {
  const pages   = Math.ceil(total / PAGE_SIZE);
  const pg      = document.getElementById('pagination');
  if (pages <= 1) { pg.innerHTML = ''; return; }

  let html = `<button class="page-btn" onclick="goPage(${currentPage-1})" ${currentPage===1?'disabled':''}>‹ Prev</button>`;

  const range = getPageRange(currentPage, pages);
  range.forEach(p => {
    if (p === '…') html += `<span style="padding:6px 4px;color:var(--text-muted)">…</span>`;
    else html += `<button class="page-btn ${p===currentPage?'active':''}" onclick="goPage(${p})">${p}</button>`;
  });

  html += `<button class="page-btn" onclick="goPage(${currentPage+1})" ${currentPage===pages?'disabled':''}>Next ›</button>`;
  pg.innerHTML = html;
}

function getPageRange(cur, total) {
  if (total <= 7) return Array.from({length:total},(_,i)=>i+1);
  const pages = [];
  pages.push(1);
  if (cur > 3) pages.push('…');
  for (let p = Math.max(2, cur-1); p <= Math.min(total-1, cur+1); p++) pages.push(p);
  if (cur < total - 2) pages.push('…');
  pages.push(total);
  return pages;
}

function goPage(p) {
  const pages = Math.ceil(filteredData.length / PAGE_SIZE);
  if (p < 1 || p > pages) return;
  currentPage = p;
  renderDetailTable();
  document.getElementById('detailTable').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ────────────────────────────────────────────────────────────
// EXPORT CSV
// ────────────────────────────────────────────────────────────
function exportFilteredData() {
  if (!filteredData.length) return;
  const ws = XLSX.utils.json_to_sheet(filteredData);
  const csv = XLSX.utils.sheet_to_csv(ws);
  const blob = new Blob([csv], { type: 'text/csv' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = 'rereg_filtered_data.csv';
  a.click();
  URL.revokeObjectURL(url);
}

// ────────────────────────────────────────────────────────────
// LOADING
// ────────────────────────────────────────────────────────────
function showLoading(show) {
  document.getElementById('loadingScreen').classList.toggle('hidden', !show);
}
