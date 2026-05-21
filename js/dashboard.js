/* ============================================================
   RE-REGISTRATION DASHBOARD — MAIN JAVASCRIPT
   ============================================================ */

'use strict';

// ────────────────────────────────────────────────────────────
// STATE
// ────────────────────────────────────────────────────────────
let rawData       = [];
let filteredData  = [];
let allColumns    = [];
let columnMap     = {};
let charts        = {};
let currentPage   = 1;

const PAGE_SIZE = 25;

// ────────────────────────────────────────────────────────────
// COLUMN DETECTION
// ────────────────────────────────────────────────────────────
const COL_PATTERNS = {
  batch:       [/batch/i, /cohort/i],
  program:     [/program/i, /programme/i, /course/i, /degree/i],
  studentId:   [/student.?id/i, /roll/i, /reg/i],
  studentName: [/student.?name/i, /name/i],
  sem:         [/sem/i, /semester/i],
  gender:      [/gender/i],
  mobile:      [/mobile/i, /phone/i],
  email:       [/email/i],
  center:      [/center/i, /centre/i, /campus/i],
  iaStatus:    [/ia/i, /internal/i]
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
// FILE INPUT
// ────────────────────────────────────────────────────────────
document
  .getElementById('fileInput')
  .addEventListener('change', handleFile);

function handleFile(e) {
  const file = e.target.files[0];

  if (file) processFile(file);
}

function processFile(file) {
  showLoading(true);

  const reader = new FileReader();

  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);

      const workbook = XLSX.read(data, {
        type: 'array',
        cellDates: true
      });

      parseWorkbook(workbook, file.name);

    } catch (err) {
      showLoading(false);
      alert('Error reading file: ' + err.message);
    }
  };

  reader.readAsArrayBuffer(file);
}

// ────────────────────────────────────────────────────────────
// PARSE WORKBOOK
// ────────────────────────────────────────────────────────────
function parseWorkbook(workbook, fileName) {

  rawData = [];

  let bestSheet = null;
  let maxRows   = 0;

  workbook.SheetNames.forEach(name => {

    const sheet = workbook.Sheets[name];

    const json = XLSX.utils.sheet_to_json(sheet, {
      defval: ''
    });

    if (json.length > maxRows) {
      maxRows   = json.length;
      bestSheet = sheet;
    }
  });

  if (!bestSheet) {
    alert('No data found');
    showLoading(false);
    return;
  }

  const json = XLSX.utils.sheet_to_json(bestSheet, {
    defval: ''
  });

  if (!json.length) {
    alert('Sheet empty');
    showLoading(false);
    return;
  }

  rawData    = json;
  allColumns = Object.keys(json[0]);

  // Detect columns
  columnMap = {};

  Object.keys(COL_PATTERNS).forEach(key => {
    columnMap[key] = detectColumn(allColumns, key);
  });

  // Last Updated
  const lu = document.getElementById('lastUpdated');

  lu.classList.remove('hidden');

  lu.querySelector('span').textContent =
    `Loaded: ${fileName} · ${rawData.length} rows`;

  populateFilters();

  filteredData = [...rawData];

  renderDashboard();

  showLoading(false);

  document
    .getElementById('uploadZone')
    .classList.add('hidden');

  document
    .getElementById('dashboard')
    .classList.remove('hidden');
}

// ────────────────────────────────────────────────────────────
// HELPERS
// ────────────────────────────────────────────────────────────
function getVal(row, key) {
  const col = columnMap[key];

  return col ? normalizeVal(row[col]) : '';
}

/* ============================================================
   IMPORTANT CHANGE
   Re-Reg is calculated ONLY from LAST COLUMN
   ============================================================ */

function isReRegDone(row) {

  const values = Object.values(row);

  if (!values.length) return false;

  const lastColumnValue = String(
    values[values.length - 1] || ''
  )
    .trim()
    .replace(/\s+/g, ' ')
    .toLowerCase();

  return [
    'done',
    'completed',
    'complete',
    'success',
    'registered',
    'yes',
    'true',
    '1'
  ].includes(lastColumnValue);
}

// IA Status
function isIASubmitted(row) {

  const ia = getVal(row, 'iaStatus')
    .toLowerCase();

  return /\b(done|submitted|complete|yes|true|1)\b/
    .test(ia);
}

function pct(num, den) {

  if (!den) return 0;

  return Math.round((num / den) * 1000) / 10;
}

function pctClass(p) {

  if (p >= 75) return 'pct-high';

  if (p >= 50) return 'pct-med';

  return 'pct-low';
}

// ────────────────────────────────────────────────────────────
// FILTERS
// ────────────────────────────────────────────────────────────
function populateFilters() {

  const batchSel   =
    document.getElementById('filterBatch');

  const programSel =
    document.getElementById('filterProgram');

  batchSel.innerHTML =
    '<option value="ALL">All Batches</option>';

  programSel.innerHTML =
    '<option value="ALL">All Programs</option>';

  const batches = [
    ...new Set(
      rawData
        .map(r => getVal(r, 'batch'))
        .filter(Boolean)
    )
  ].sort();

  const programs = [
    ...new Set(
      rawData
        .map(r => getVal(r, 'program'))
        .filter(Boolean)
    )
  ].sort();

  batches.forEach(b => {

    const o = document.createElement('option');

    o.value = b;
    o.textContent = b;

    batchSel.appendChild(o);
  });

  programs.forEach(p => {

    const o = document.createElement('option');

    o.value = p;
    o.textContent = p;

    programSel.appendChild(o);
  });
}

function applyFilters() {

  const batch =
    document.getElementById('filterBatch').value;

  const program =
    document.getElementById('filterProgram').value;

  const status =
    document.getElementById('filterStatus').value;

  const search =
    document.getElementById('searchInput')
      .value
      .toLowerCase();

  filteredData = rawData.filter(row => {

    if (
      batch !== 'ALL' &&
      getVal(row, 'batch') !== batch
    ) return false;

    if (
      program !== 'ALL' &&
      getVal(row, 'program') !== program
    ) return false;

    if (
      status === 'DONE' &&
      !isReRegDone(row)
    ) return false;

    if (
      status === 'PENDING' &&
      isReRegDone(row)
    ) return false;

    if (search) {

      const rowStr = Object.values(row)
        .join(' ')
        .toLowerCase();

      if (!rowStr.includes(search))
        return false;
    }

    return true;
  });

  currentPage = 1;

  renderDashboard();
}

function resetFilters() {

  document.getElementById('filterBatch').value =
    'ALL';

  document.getElementById('filterProgram').value =
    'ALL';

  document.getElementById('filterStatus').value =
    'ALL';

  document.getElementById('searchInput').value =
    '';

  filteredData = [...rawData];

  currentPage = 1;

  renderDashboard();
}

[
  'filterBatch',
  'filterProgram',
  'filterStatus'
].forEach(id => {

  document
    .getElementById(id)
    .addEventListener('change', applyFilters);
});

document
  .getElementById('searchInput')
  .addEventListener('input', applyFilters);

// ────────────────────────────────────────────────────────────
// ANALYTICS
// ────────────────────────────────────────────────────────────
function computeStats(data) {

  const total = data.length;

  const reRegDone =
    data.filter(isReRegDone).length;

  const reRegPend =
    total - reRegDone;

  const reRegPct =
    pct(reRegDone, total);

  const iaSubmitted =
    data.filter(isIASubmitted).length;

  const iaPct =
    pct(iaSubmitted, total);

  // Batch-wise
  const batchMap = {};

  data.forEach(row => {

    const batch =
      getVal(row, 'batch') || 'Unknown';

    const program =
      getVal(row, 'program') || 'Unknown';

    if (!batchMap[batch]) {

      batchMap[batch] = {
        batch,
        program: new Set(),
        total: 0,
        done: 0,
        ia: 0
      };
    }

    batchMap[batch].total++;

    batchMap[batch].program.add(program);

    if (isReRegDone(row))
      batchMap[batch].done++;

    if (isIASubmitted(row))
      batchMap[batch].ia++;
  });

  const batchStats = Object
    .values(batchMap)
    .map(b => ({

      batch: b.batch,

      program:
        [...b.program].join(', '),

      total: b.total,

      done: b.done,

      pending:
        b.total - b.done,

      reRegPct:
        pct(b.done, b.total),

      ia: b.ia,

      iaPct:
        pct(b.ia, b.total)

    }));

  return {
    total,
    reRegDone,
    reRegPend,
    reRegPct,
    iaSubmitted,
    iaPct,
    batchStats
  };
}

// ────────────────────────────────────────────────────────────
// RENDER DASHBOARD
// ────────────────────────────────────────────────────────────
function renderDashboard() {

  const stats =
    computeStats(filteredData);

  renderKPIs(stats);

  renderBatchTable(stats.batchStats);

  renderCharts(stats);

  renderDetailTable();
}

// ────────────────────────────────────────────────────────────
// KPI
// ────────────────────────────────────────────────────────────
function renderKPIs(stats) {

  document.getElementById('totalStudents')
    .textContent = stats.total;

  document.getElementById('reRegDone')
    .textContent = stats.reRegDone;

  document.getElementById('reRegPending')
    .textContent = stats.reRegPend;

  document.getElementById('reRegPct')
    .textContent = stats.reRegPct + '%';

  document.getElementById('iaPct')
    .textContent = stats.iaPct + '%';
}

// ────────────────────────────────────────────────────────────
// BATCH TABLE
// ────────────────────────────────────────────────────────────
function renderBatchTable(batchStats) {

  const tbody =
    document.getElementById('batchSummaryBody');

  if (!batchStats.length) {

    tbody.innerHTML =
      '<tr><td colspan="8">No Data</td></tr>';

    return;
  }

  tbody.innerHTML = batchStats.map(b => `

    <tr>
      <td>${b.batch}</td>
      <td>${b.program}</td>
      <td>${b.total}</td>
      <td>${b.done}</td>
      <td>${b.pending}</td>
      <td>${b.reRegPct}%</td>
      <td>${b.ia}</td>
      <td>${b.iaPct}%</td>
    </tr>

  `).join('');
}

// ────────────────────────────────────────────────────────────
// CHARTS
// ────────────────────────────────────────────────────────────
function destroyCharts() {

  Object.values(charts).forEach(c => {

    if (c) c.destroy();
  });

  charts = {};
}

function renderCharts(stats) {

  destroyCharts();

  const labels =
    stats.batchStats.map(b => b.batch);

  // BAR CHART
  const ctx1 =
    document
      .getElementById('reRegBarChart')
      .getContext('2d');

  charts.bar = new Chart(ctx1, {

    type: 'bar',

    data: {

      labels,

      datasets: [{

        label: 'Re-Reg %',

        data:
          stats.batchStats
            .map(b => b.reRegPct),

        backgroundColor: '#4f46e5'

      }]
    }
  });

  // PIE CHART
  const ctx2 =
    document
      .getElementById('reRegPieChart')
      .getContext('2d');

  charts.pie = new Chart(ctx2, {

    type: 'doughnut',

    data: {

      labels: ['Done', 'Pending'],

      datasets: [{

        data: [
          stats.reRegDone,
          stats.reRegPend
        ],

        backgroundColor: [
          '#10b981',
          '#ef4444'
        ]
      }]
    }
  });
}

// ────────────────────────────────────────────────────────────
// DETAIL TABLE
// ────────────────────────────────────────────────────────────
function renderDetailTable() {

  const head =
    document.getElementById('detailTableHead');

  const body =
    document.getElementById('detailTableBody');

  const start =
    (currentPage - 1) * PAGE_SIZE;

  const end =
    start + PAGE_SIZE;

  const page =
    filteredData.slice(start, end);

  let showCols = allColumns;

  head.innerHTML =
    '<tr>' +
    showCols.map(c => `<th>${c}</th>`).join('') +
    '</tr>';

  body.innerHTML = page.map(row => {

    return '<tr>' +

      showCols.map(col => {

        let val = normalizeVal(row[col]);

        // Highlight LAST COLUMN status
        const isLast =
          col === allColumns[allColumns.length - 1];

        if (isLast) {

          const done =
            isReRegDone(row);

          return `
            <td>
              <span class="status-pill ${
                done
                  ? 'status-done'
                  : 'status-pending'
              }">
                ${val || 'Pending'}
              </span>
            </td>
          `;
        }

        return `<td>${val || '—'}</td>`;

      }).join('')

      + '</tr>';

  }).join('');
}

// ────────────────────────────────────────────────────────────
// EXPORT CSV
// ────────────────────────────────────────────────────────────
function exportFilteredData() {

  if (!filteredData.length) return;

  const ws =
    XLSX.utils.json_to_sheet(filteredData);

  const csv =
    XLSX.utils.sheet_to_csv(ws);

  const blob = new Blob(
    [csv],
    { type: 'text/csv' }
  );

  const url =
    URL.createObjectURL(blob);

  const a =
    document.createElement('a');

  a.href = url;

  a.download =
    'rereg_filtered_data.csv';

  a.click();

  URL.revokeObjectURL(url);
}

// ────────────────────────────────────────────────────────────
// PAGINATION
// ────────────────────────────────────────────────────────────
function goPage(page) {

  currentPage = page;

  renderDetailTable();
}

// ────────────────────────────────────────────────────────────
// LOADING
// ────────────────────────────────────────────────────────────
function showLoading(show) {

  document
    .getElementById('loadingScreen')
    .classList.toggle('hidden', !show);
}
