/* ============================================================
   RE-REGISTRATION DASHBOARD — COMPLETE UPDATED SCRIPT
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

function findColumn(possibleNames) {

  return allColumns.find(col => {

    const c = col.toLowerCase();

    return possibleNames.some(name =>
      c.includes(name.toLowerCase())
    );
  });
}

function getColumnValue(row, possibleNames) {

  const col = findColumn(possibleNames);

  return normalizeVal(row[col]);
}

/* ============================================================
   RE-REG STATUS
============================================================ */

function isReRegDone(row) {

  const values = Object.values(row);

  if (!values.length) return false;

  const lastValue = String(
    values[values.length - 1] || ''
  )
    .trim()
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

/* ============================================================
   IA STATUS
============================================================ */

function getIAStatus(row) {

  const iaValue =
    getColumnValue(row, [
      'sem 1 ia',
      'ia status',
      'ia',
      'internal assessment'
    ]).toLowerCase();

  if (
    iaValue.includes('partial')
  ) {
    return 'PARTIAL';
  }

  if (
    iaValue.includes('submit') ||
    iaValue.includes('done') ||
    iaValue.includes('completed')
  ) {
    return 'SUBMITTED';
  }

  return 'NIL';
}

/* ============================================================
   PERCENT
============================================================ */

function pct(num, den) {

  if (!den) return 0;

  return Math.round((num / den) * 1000) / 10;
}

/* ============================================================
   FILTERS
============================================================ */

function populateFilters() {

  populateSingleFilter(
    'filterBatch',
    ['batch']
  );

  populateSingleFilter(
    'filterProgram',
    ['program']
  );

  populateSingleFilter(
    'filterSource',
    ['source']
  );

  populateSingleFilter(
    'filterSalesType',
    ['sales type', 'sales_type']
  );

  populateSingleFilter(
    'filterUGC',
    ['ugc']
  );
}

function populateSingleFilter(id, possibleNames) {

  const select =
    document.getElementById(id);

  if (!select) return;

  const values = [
    ...new Set(
      rawData
        .map(r =>
          getColumnValue(r, possibleNames)
        )
        .filter(Boolean)
    )
  ];

  values.sort();

  values.forEach(v => {

    const opt =
      document.createElement('option');

    opt.value = v;
    opt.textContent = v;

    select.appendChild(opt);
  });
}

function applyFilters() {

  const batch =
    document.getElementById('filterBatch').value;

  const program =
    document.getElementById('filterProgram').value;

  const source =
    document.getElementById('filterSource').value;

  const salesType =
    document.getElementById('filterSalesType').value;

  const ugc =
    document.getElementById('filterUGC').value;

  const status =
    document.getElementById('filterStatus').value;

  const search =
    document.getElementById('searchInput')
      .value
      .toLowerCase();

  filteredData = rawData.filter(row => {

    if (
      batch !== 'ALL' &&
      getColumnValue(row, ['batch']) !== batch
    ) return false;

    if (
      program !== 'ALL' &&
      getColumnValue(row, ['program']) !== program
    ) return false;

    if (
      source !== 'ALL' &&
      getColumnValue(row, ['source']) !== source
    ) return false;

    if (
      salesType !== 'ALL' &&
      getColumnValue(row, ['sales type']) !== salesType
    ) return false;

    if (
      ugc !== 'ALL' &&
      getColumnValue(row, ['ugc']) !== ugc
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

  [
    'filterBatch',
    'filterProgram',
    'filterSource',
    'filterSalesType',
    'filterUGC',
    'filterStatus'
  ].forEach(id => {

    const el =
      document.getElementById(id);

    if (el) el.value = 'ALL';
  });

  document.getElementById(
    'searchInput'
  ).value = '';

  filteredData = [...rawData];

  renderDashboard();
}

[
  'filterBatch',
  'filterProgram',
  'filterSource',
  'filterSalesType',
  'filterUGC',
  'filterStatus'
].forEach(id => {

  const el =
    document.getElementById(id);

  if (el) {
    el.addEventListener(
      'change',
      applyFilters
    );
  }
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
   BATCH STATS
============================================================ */

function getBatchStats(data) {

  const map = {};

  data.forEach(row => {

    const batch =
      getColumnValue(row, ['batch']) ||
      'Unknown';

    if (!map[batch]) {

      map[batch] = {
        batch,
        total: 0,
        done: 0,
        iaSubmitted: 0,
        iaPartial: 0,
        iaNil: 0
      };
    }

    map[batch].total++;

    if (isReRegDone(row)) {
      map[batch].done++;
    }

    const ia =
      getIAStatus(row);

    if (ia === 'SUBMITTED')
      map[batch].iaSubmitted++;

    else if (ia === 'PARTIAL')
      map[batch].iaPartial++;

    else
      map[batch].iaNil++;
  });

  return Object.values(map);
}

/* ============================================================
   PROGRAM STATS
============================================================ */

function getProgramStats(data) {

  const map = {};

  data.forEach(row => {

    const program =
      getColumnValue(row, ['program']) ||
      'Unknown';

    if (!map[program]) {

      map[program] = {
        program,
        total: 0,
        done: 0,
        iaSubmitted: 0,
        iaPartial: 0,
        iaNil: 0
      };
    }

    map[program].total++;

    if (isReRegDone(row)) {
      map[program].done++;
    }

    const ia =
      getIAStatus(row);

    if (ia === 'SUBMITTED')
      map[program].iaSubmitted++;

    else if (ia === 'PARTIAL')
      map[program].iaPartial++;

    else
      map[program].iaNil++;
  });

  return Object.values(map);
}

/* ============================================================
   RENDER DASHBOARD
============================================================ */

function renderDashboard() {

  const stats =
    computeStats(filteredData);

  renderKPIs(stats);

  renderBatchTable();

  renderProgramTable();

  renderCharts();

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
   BATCH TABLE
============================================================ */

function renderBatchTable() {

  const tbody =
    document.getElementById(
      'batchSummaryBody'
    );

  const stats =
    getBatchStats(filteredData);

  tbody.innerHTML = stats.map(s => {

    const pending =
      s.total - s.done;

    return `
      <tr>

        <td>${s.batch}</td>

        <td>${s.total}</td>

        <td>${s.done}</td>

        <td>${pending}</td>

        <td>${pct(s.done, s.total)}%</td>

        <td>
          ${s.iaSubmitted}
          (${pct(s.iaSubmitted, s.total)}%)
        </td>

        <td>
          ${s.iaPartial}
          (${pct(s.iaPartial, s.total)}%)
        </td>

        <td>
          ${s.iaNil}
          (${pct(s.iaNil, s.total)}%)
        </td>

      </tr>
    `;
  }).join('');
}

/* ============================================================
   PROGRAM TABLE
============================================================ */

function renderProgramTable() {

  const grid =
    document.getElementById(
      'programGrid'
    );

  const stats =
    getProgramStats(filteredData);

  grid.innerHTML = `

    <div class="table-responsive">

      <table class="data-table">

        <thead>

          <tr>

            <th>Program</th>
            <th>Total Students</th>
            <th>Re-Reg Done</th>
            <th>Pending</th>
            <th>Re-Reg %</th>
            <th>IA Submitted</th>
            <th>IA Partial</th>
            <th>IA Nil</th>

          </tr>

        </thead>

        <tbody>

          ${stats.map(s => {

            const pending =
              s.total - s.done;

            return `

              <tr>

                <td>${s.program}</td>

                <td>${s.total}</td>

                <td>${s.done}</td>

                <td>${pending}</td>

                <td>
                  ${pct(s.done, s.total)}%
                </td>

                <td>
                  ${s.iaSubmitted}
                  (${pct(s.iaSubmitted, s.total)}%)
                </td>

                <td>
                  ${s.iaPartial}
                  (${pct(s.iaPartial, s.total)}%)
                </td>

                <td>
                  ${s.iaNil}
                  (${pct(s.iaNil, s.total)}%)
                </td>

              </tr>

            `;
          }).join('')}

        </tbody>

      </table>

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

function renderCharts() {

  destroyCharts();

  const batchStats =
    getBatchStats(filteredData);

  const labels =
    batchStats.map(b => b.batch);

  /* PIE */

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

          batchStats.reduce(
            (a,b)=>a+b.done,0
          ),

          batchStats.reduce(
            (a,b)=>a+(b.total-b.done),0
          )

        ],

        backgroundColor: [
          '#10b981',
          '#ef4444'
        ]
      }]
    }
  });

  /* RE-REG BAR */

  const barCtx =
    document
      .getElementById('reRegBarChart')
      .getContext('2d');

  charts.bar = new Chart(barCtx, {

    type: 'bar',

    data: {

      labels,

      datasets: [{

        label: 'Re-Reg %',

        data: batchStats.map(
          b => pct(b.done, b.total)
        ),

        backgroundColor: '#4f46e5'
      }]
    },

    options: {
      responsive: true,
      scales: {
        y: {
          beginAtZero: true,
          max: 100
        }
      }
    }
  });

  /* IA BAR */

  const iaCtx =
    document
      .getElementById('iaBarChart')
      .getContext('2d');

  charts.ia = new Chart(iaCtx, {

    type: 'bar',

    data: {

      labels,

      datasets: [

        {
          label: 'IA Submitted %',

          data: batchStats.map(
            b => pct(
              b.iaSubmitted,
              b.total
            )
          ),

          backgroundColor: '#10b981'
        },

        {
          label: 'IA Partial %',

          data: batchStats.map(
            b => pct(
              b.iaPartial,
              b.total
            )
          ),

          backgroundColor: '#f59e0b'
        },

        {
          label: 'IA Nil %',

          data: batchStats.map(
            b => pct(
              b.iaNil,
              b.total
            )
          ),

          backgroundColor: '#ef4444'
        }

      ]
    },

    options: {
      responsive: true,
      scales: {
        y: {
          beginAtZero: true,
          max: 100
        }
      }
    }
  });

  /* STACKED */

  const stackCtx =
    document
      .getElementById('stackedBarChart')
      .getContext('2d');

  charts.stack = new Chart(stackCtx, {

    type: 'bar',

    data: {

      labels,

      datasets: [

        {
          label: 'Done',

          data: batchStats.map(
            b => b.done
          ),

          backgroundColor: '#10b981'
        },

        {
          label: 'Pending',

          data: batchStats.map(
            b => b.total - b.done
          ),

          backgroundColor: '#ef4444'
        }

      ]
    },

    options: {

      responsive: true,

      scales: {

        x: {
          stacked: true
        },

        y: {
          stacked: true
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
    document.getElementById(
      'detailTableHead'
    );

  const body =
    document.getElementById(
      'detailTableBody'
    );

  const meta =
    document.getElementById(
      'tableMeta'
    );

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

      allColumns.map(col => {

        const val =
          normalizeVal(row[col]);

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
