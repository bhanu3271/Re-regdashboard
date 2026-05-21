/* ============================================================
   RE-REGISTRATION DASHBOARD — FINAL SCRIPT
   ============================================================ */

'use strict';

/* ============================================================
   STATE
============================================================ */

let rawData = [];
let filteredData = [];
let allColumns = [];
let charts = {};
let currentPage = 1;

const PAGE_SIZE = 25;

/* ============================================================
   FILE HANDLING
============================================================ */

document
  .getElementById('fileInput')
  .addEventListener('change', handleFile);

const uploadCard =
  document.querySelector('.upload-card');

uploadCard.addEventListener('dragover', e => {
  e.preventDefault();
  uploadCard.classList.add('drag-over');
});

uploadCard.addEventListener('dragleave', () => {
  uploadCard.classList.remove('drag-over');
});

uploadCard.addEventListener('drop', e => {

  e.preventDefault();

  uploadCard.classList.remove('drag-over');

  const file = e.dataTransfer.files[0];

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

      const data =
        new Uint8Array(e.target.result);

      const workbook = XLSX.read(data, {
        type: 'array'
      });

      parseWorkbook(workbook, file.name);

    } catch (err) {

      showLoading(false);

      alert(
        'Error reading file: ' + err.message
      );
    }
  };

  reader.readAsArrayBuffer(file);
}

/* ============================================================
   PARSE WORKBOOK
============================================================ */

function parseWorkbook(workbook, fileName) {

  rawData = [];

  let bestSheet = null;
  let maxRows = 0;

  workbook.SheetNames.forEach(name => {

    const sheet = workbook.Sheets[name];

    const json =
      XLSX.utils.sheet_to_json(sheet, {
        defval: ''
      });

    if (json.length > maxRows) {

      maxRows = json.length;
      bestSheet = sheet;
    }
  });

  if (!bestSheet) {

    alert('No data found');

    showLoading(false);

    return;
  }

  rawData =
    XLSX.utils.sheet_to_json(bestSheet, {
      defval: ''
    });

  allColumns = Object.keys(rawData[0]);

  filteredData = [...rawData];

  populateFilters();

  renderDashboard();

  // last updated
  const lu =
    document.getElementById('lastUpdated');

  lu.classList.remove('hidden');

  lu.querySelector('span').textContent =
    `Loaded: ${fileName} · ${rawData.length} rows`;

  document
    .getElementById('uploadZone')
    .classList.add('hidden');

  document
    .getElementById('dashboard')
    .classList.remove('hidden');

  showLoading(false);
}

/* ============================================================
   HELPERS
============================================================ */

function normalizeVal(val) {

  if (val === null || val === undefined)
    return '';

  return String(val).trim();
}

/* ============================================================
   IMPORTANT:
   Re-Reg Status from LAST COLUMN ONLY
============================================================ */

function isReRegDone(row) {

  const values = Object.values(row);

  if (!values.length) return false;

  const lastValue = String(
    values[values.length - 1] || ''
  )
    .trim()
    .replace(/\s+/g, ' ')
    .toLowerCase();

  return [
    'done',
    'completed',
    'complete',
    'yes',
    'success',
    'registered',
    'true',
    '1'
  ].includes(lastValue);
}

function pct(num, den) {

  if (!den) return 0;

  return Math.round((num / den) * 1000) / 10;
}

/* ============================================================
   FILTERS
============================================================ */

function populateFilters() {

  const batchSel =
    document.getElementById('filterBatch');

  const programSel =
    document.getElementById('filterProgram');

  const batches = [
    ...new Set(
      rawData.map(r => r.Batch).filter(Boolean)
    )
  ];

  const programs = [
    ...new Set(
      rawData.map(r => r.Program).filter(Boolean)
    )
  ];

  batches.sort().forEach(b => {

    const opt =
      document.createElement('option');

    opt.value = b;
    opt.textContent = b;

    batchSel.appendChild(opt);
  });

  programs.sort().forEach(p => {

    const opt =
      document.createElement('option');

    opt.value = p;
    opt.textContent = p;

    programSel.appendChild(opt);
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
      row.Batch !== batch
    ) return false;

    if (
      program !== 'ALL' &&
      row.Program !== program
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

      const rowText =
        Object.values(row)
          .join(' ')
          .toLowerCase();

      if (!rowText.includes(search))
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

/* ============================================================
   STATS
============================================================ */

function computeStats(data) {

  const total = data.length;

  const done =
    data.filter(isReRegDone).length;

  const pending =
    total - done;

  const percent =
    pct(done, total);

  return {
    total,
    done,
    pending,
    percent
  };
}

/* ============================================================
   RENDER DASHBOARD
============================================================ */

function renderDashboard() {

  const stats =
    computeStats(filteredData);

  renderKPIs(stats);

  renderCharts(stats);

  renderDetailTable();
}

/* ============================================================
   KPI CARDS
============================================================ */

function renderKPIs(stats) {

  const grid =
    document.getElementById('kpiGrid');

  grid.innerHTML = `

    <div class="kpi-card">
      <div class="kpi-label">Total Students</div>
      <div class="kpi-value">${stats.total}</div>
    </div>

    <div class="kpi-card">
      <div class="kpi-label">Re-Reg Done</div>
      <div class="kpi-value">${stats.done}</div>
    </div>

    <div class="kpi-card">
      <div class="kpi-label">Pending</div>
      <div class="kpi-value">${stats.pending}</div>
    </div>

    <div class="kpi-card">
      <div class="kpi-label">Completion %</div>
      <div class="kpi-value">${stats.percent}%</div>
    </div>
  `;
}

/* ============================================================
   CHARTS
============================================================ */

function destroyCharts() {

  Object.values(charts).forEach(c => {

    if (c) c.destroy();
  });

  charts = {};
}

function renderCharts(stats) {

  destroyCharts();

  // Pie Chart
  const pieCtx =
    document
      .getElementById('reRegPieChart')
      .getContext('2d');

  charts.pie = new Chart(pieCtx, {

    type: 'doughnut',

    data: {

      labels: ['Done', 'Pending'],

      datasets: [{
        data: [
          stats.done,
          stats.pending
        ],

        backgroundColor: [
          '#10b981',
          '#ef4444'
        ]
      }]
    }
  });

  // Bar Chart
  const barCtx =
    document
      .getElementById('reRegBarChart')
      .getContext('2d');

  charts.bar = new Chart(barCtx, {

    type: 'bar',

    data: {

      labels: ['Completion'],

      datasets: [{

        label: 'Re-Reg %',

        data: [stats.percent],

        backgroundColor: '#4f46e5'
      }]
    },

    options: {
      scales: {
        y: {
          beginAtZero: true,
          max: 100
        }
      }
    }
  });
}

/* ============================================================
   DETAIL TABLE
============================================================ */

function renderDetailTable() {

  const head =
    document.getElementById('detailTableHead');

  const body =
    document.getElementById('detailTableBody');

  const meta =
    document.getElementById('tableMeta');

  const start =
    (currentPage - 1) * PAGE_SIZE;

  const end =
    start + PAGE_SIZE;

  const page =
    filteredData.slice(start, end);

  meta.textContent =
    `Showing ${start + 1} - ${Math.min(end, filteredData.length)}
     of ${filteredData.length}`;

  head.innerHTML =
    '<tr>' +
    allColumns.map(c =>
      `<th>${c}</th>`
    ).join('') +
    '</tr>';

  body.innerHTML = page.map(row => {

    return '<tr>' +

      allColumns.map((col, idx) => {

        const val =
          normalizeVal(row[col]);

        const isLast =
          idx === allColumns.length - 1;

        if (isLast) {

          return `
            <td>
              <span class="${
                isReRegDone(row)
                  ? 'status-done'
                  : 'status-pending'
              }">
                ${val}
              </span>
            </td>
          `;
        }

        return `<td>${val}</td>`;

      }).join('')

      + '</tr>';

  }).join('');
}

/* ============================================================
   EXPORT CSV
============================================================ */

function exportFilteredData() {

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
    'filtered_data.csv';

  a.click();

  URL.revokeObjectURL(url);
}

/* ============================================================
   LOADING
============================================================ */

function showLoading(show) {

  document
    .getElementById('loadingScreen')
    .classList.toggle('hidden', !show);
}
