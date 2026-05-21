/* ============================================================
   RE-REGISTRATION DASHBOARD — UPDATED FINAL SCRIPT
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

  reader.onload = function (e) {

    try {

      const data =
        new Uint8Array(e.target.result);

      const workbook = XLSX.read(data, {
        type: 'array'
      });

      parseWorkbook(workbook, file.name);

    } catch (err) {

      console.error(err);

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

  if (!rawData.length) {

    alert('Sheet is empty');

    showLoading(false);

    return;
  }

  allColumns = Object.keys(rawData[0]);

  filteredData = [...rawData];

  populateFilters();

  renderDashboard();

  /* LAST UPDATED */

  const lu =
    document.getElementById('lastUpdated');

  if (lu) {

    lu.classList.remove('hidden');

    const span = lu.querySelector('span');

    if (span) {

      span.textContent =
        `Loaded: ${fileName} · ${rawData.length} rows`;
    }
  }

  /* SHOW DASHBOARD */

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

function getColumnValue(row, keywords) {

  for (const col of allColumns) {

    const lower =
      col.toLowerCase();

    const matched =
      keywords.some(k =>
        lower.includes(k)
      );

    if (matched) {

      return normalizeVal(row[col]);
    }
  }

  return '';
}

/* ============================================================
   RE-REG STATUS
   LAST COLUMN ONLY
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

  return Math.round(
    (num / den) * 1000
  ) / 10;
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
    ['sales type', 'salestype']
  );

  populateSingleFilter(
    'filterUGC',
    ['ugc']
  );
}

function populateSingleFilter(id, keywords) {

  const select =
    document.getElementById(id);

  if (!select) return;

  const values = [
    ...new Set(
      rawData
        .map(r =>
          getColumnValue(r, keywords)
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
    document.getElementById('filterBatch')?.value || 'ALL';

  const program =
    document.getElementById('filterProgram')?.value || 'ALL';

  const source =
    document.getElementById('filterSource')?.value || 'ALL';

  const salesType =
    document.getElementById('filterSalesType')?.value || 'ALL';

  const ugc =
    document.getElementById('filterUGC')?.value || 'ALL';

  const status =
    document.getElementById('filterStatus')?.value || 'ALL';

  const search =
    document.getElementById('searchInput')
      ?.value
      .toLowerCase() || '';

  filteredData = rawData.filter(row => {

    const batchVal =
      getColumnValue(row, ['batch']);

    const programVal =
      getColumnValue(row, ['program']);

    const sourceVal =
      getColumnValue(row, ['source']);

    const salesVal =
      getColumnValue(row, ['sales type', 'salestype']);

    const ugcVal =
      getColumnValue(row, ['ugc']);

    if (
      batch !== 'ALL' &&
      batchVal !== batch
    ) return false;

    if (
      program !== 'ALL' &&
      programVal !== program
    ) return false;

    if (
      source !== 'ALL' &&
      sourceVal !== source
    ) return false;

    if (
      salesType !== 'ALL' &&
      salesVal !== salesType
    ) return false;

    if (
      ugc !== 'ALL' &&
      ugcVal !== ugc
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

      const text =
        Object.values(row)
          .join(' ')
          .toLowerCase();

      if (!text.includes(search))
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

  const search =
    document.getElementById('searchInput');

  if (search) search.value = '';

  filteredData = [...rawData];

  renderDashboard();
}

/* FILTER EVENTS */

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
  ?.addEventListener(
    'input',
    applyFilters
  );

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
   DASHBOARD RENDER
============================================================ */

function renderDashboard() {

  const stats =
    computeStats(filteredData);

  renderKPIs(stats);

  renderCharts(stats);

  renderBatchTable();

  renderDetailTable();
}

/* ============================================================
   KPI CARDS
============================================================ */

function renderKPIs(stats) {

  const grid =
    document.getElementById('kpiGrid');

  if (!grid) return;

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

  if (!tbody) return;

  const batchMap = {};

  filteredData.forEach(row => {

    const batch =
      getColumnValue(row, ['batch']) || 'Unknown';

    const program =
      getColumnValue(row, ['program']) || 'Unknown';

    if (!batchMap[batch]) {

      batchMap[batch] = {
        batch,
        program,
        total: 0,
        done: 0
      };
    }

    batchMap[batch].total++;

    if (isReRegDone(row)) {

      batchMap[batch].done++;
    }
  });

  const rows =
    Object.values(batchMap);

  tbody.innerHTML = rows.map(r => {

    const pending =
      r.total - r.done;

    const percent =
      pct(r.done, r.total);

    return `
      <tr>
        <td>${r.batch}</td>
        <td>${r.program}</td>
        <td>${r.total}</td>
        <td>${r.done}</td>
        <td>${pending}</td>
        <td>${percent}%</td>
        <td>${r.done}</td>
        <td>${percent}%</td>
        <td>
          ${percent >= 50 ? '✅ Good' : '⚠️ Pending'}
        </td>
      </tr>
    `;
  }).join('');
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

  /* PIE */

  const pieCanvas =
    document.getElementById(
      'reRegPieChart'
    );

  if (pieCanvas) {

    charts.pie = new Chart(
      pieCanvas,
      {
        type: 'doughnut',

        data: {

          labels: [
            'Done',
            'Pending'
          ],

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
      }
    );
  }

  /* BAR */

  const barCanvas =
    document.getElementById(
      'reRegBarChart'
    );

  if (barCanvas) {

    charts.bar = new Chart(
      barCanvas,
      {
        type: 'bar',

        data: {

          labels: ['Completion'],

          datasets: [{

            label: 'Re-Reg %',

            data: [stats.percent],

            backgroundColor:
              '#4f46e5'
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
      }
    );
  }

  /* IA CHART */

  const iaCanvas =
    document.getElementById(
      'iaBarChart'
    );

  if (iaCanvas) {

    charts.ia = new Chart(
      iaCanvas,
      {
        type: 'bar',

        data: {

          labels: ['IA'],

          datasets: [{
            label: 'IA %',
            data: [stats.percent],
            backgroundColor:
              '#8b5cf6'
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
      }
    );
  }

  /* STACKED */

  const stackCanvas =
    document.getElementById(
      'stackedBarChart'
    );

  if (stackCanvas) {

    charts.stack = new Chart(
      stackCanvas,
      {
        type: 'bar',

        data: {

          labels: ['Students'],

          datasets: [

            {
              label: 'Done',
              data: [stats.done],
              backgroundColor:
                '#10b981'
            },

            {
              label: 'Pending',
              data: [stats.pending],
              backgroundColor:
                '#ef4444'
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
      }
    );
  }
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

  if (!head || !body) return;

  const start =
    (currentPage - 1) * PAGE_SIZE;

  const end =
    start + PAGE_SIZE;

  const page =
    filteredData.slice(start, end);

  if (meta) {

    meta.textContent =
      `Showing ${start + 1} - ${Math.min(end, filteredData.length)}
       of ${filteredData.length}`;
  }

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
    XLSX.utils.json_to_sheet(
      filteredData
    );

  const csv =
    XLSX.utils.sheet_to_csv(ws);

  const blob = new Blob(
    [csv],
    {
      type: 'text/csv'
    }
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
    ?.classList.toggle(
      'hidden',
      !show
    );
}
