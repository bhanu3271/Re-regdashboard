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

const uploadCard = document.querySelector('.upload-card');

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

  if (lu) {

    lu.classList.remove('hidden');

    const span = lu.querySelector('span');

    if (span) {

      span.textContent =
        `Loaded: ${fileName} · ${rawData.length} rows`;
    }
  }

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

function pct(num, den) {

  if (!den) return 0;

  return Math.round((num / den) * 1000) / 10;
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

  const iaColumn = allColumns.find(col =>
    col.toLowerCase().includes('sem 1 ia')
  );

  if (!iaColumn)
    return 'Nil Submission';

  const value =
    normalizeVal(row[iaColumn]).toLowerCase();

  if (
    value.includes('partial')
  ) {
    return 'Partially Submitted';
  }

  if (
    value.includes('submitted')
  ) {
    return 'Submitted';
  }

  return 'Nil Submission';
}

/* ============================================================
   FILTERS
============================================================ */

function populateFilters() {

  const batchSel =
    document.getElementById('filterBatch');

  const programSel =
    document.getElementById('filterProgram');

  batchSel.innerHTML =
    '<option value="ALL">All Batches</option>';

  programSel.innerHTML =
    '<option value="ALL">All Programs</option>';

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

  const batchMap = {};
  const programMap = {};

  data.forEach(row => {

    const batch =
      normalizeVal(row.Batch) || 'Unknown';

    const program =
      normalizeVal(row.Program) || 'Unknown';

    if (!batchMap[batch]) {

      batchMap[batch] = {
        total: 0,
        done: 0,
        iaSubmitted: 0,
        iaPartial: 0,
        iaNil: 0
      };
    }

    if (!programMap[program]) {

      programMap[program] = {
        total: 0,
        done: 0,
        iaSubmitted: 0,
        iaPartial: 0,
        iaNil: 0
      };
    }

    batchMap[batch].total++;
    programMap[program].total++;

    if (isReRegDone(row)) {

      batchMap[batch].done++;
      programMap[program].done++;
    }

    const iaStatus =
      getIAStatus(row);

    if (iaStatus === 'Submitted') {

      batchMap[batch].iaSubmitted++;
      programMap[program].iaSubmitted++;
    }
    else if (
      iaStatus === 'Partially Submitted'
    ) {

      batchMap[batch].iaPartial++;
      programMap[program].iaPartial++;
    }
    else {

      batchMap[batch].iaNil++;
      programMap[program].iaNil++;
    }
  });

  return {
    total,
    done,
    pending,
    percent,
    batchMap,
    programMap
  };
}

/* ============================================================
   RENDER DASHBOARD
============================================================ */

function renderDashboard() {

  const stats =
    computeStats(filteredData);

  renderKPIs(stats);

  renderBatchTable(stats.batchMap);

  renderProgramTable(stats.programMap);

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
   BATCH TABLE
============================================================ */

function renderBatchTable(batchMap) {

  const tbody =
    document.getElementById(
      'batchSummaryBody'
    );

  if (!tbody) return;

  let html = '';

  Object.entries(batchMap).forEach(
    ([batch, data]) => {

      const pending =
        data.total - data.done;

      const reRegPct =
        pct(data.done, data.total);

      const iaPct =
        pct(data.iaSubmitted, data.total);

      html += `
        <tr>
          <td>${batch}</td>
          <td>${data.total}</td>
          <td>${data.done}</td>
          <td>${pending}</td>
          <td>${reRegPct}%</td>
          <td>${data.iaSubmitted}</td>
          <td>${iaPct}%</td>
          <td>${data.iaPartial}</td>
          <td>${data.iaNil}</td>
        </tr>
      `;
    }
  );

  tbody.innerHTML = html;
}

/* ============================================================
   PROGRAM TABLE
============================================================ */

function renderProgramTable(programMap) {

  const tbody =
    document.getElementById(
      'programSummaryBody'
    );

  if (!tbody) return;

  let html = '';

  Object.entries(programMap).forEach(
    ([program, data]) => {

      const pending =
        data.total - data.done;

      const reRegPct =
        pct(data.done, data.total);

      const iaPct =
        pct(data.iaSubmitted, data.total);

      html += `
        <tr>
          <td>${program}</td>
          <td>${data.total}</td>
          <td>${data.done}</td>
          <td>${pending}</td>
          <td>${reRegPct}%</td>
          <td>${data.iaSubmitted}</td>
          <td>${iaPct}%</td>
          <td>${data.iaPartial}</td>
          <td>${data.iaNil}</td>
        </tr>
      `;
    }
  );

  tbody.innerHTML = html;
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

  const batches =
    Object.keys(stats.batchMap);

  const doneData =
    batches.map(
      b => stats.batchMap[b].done
    );

  const pendingData =
    batches.map(
      b =>
        stats.batchMap[b].total -
        stats.batchMap[b].done
    );

  // PIE CHART

  const pieCanvas =
    document.getElementById(
      'reRegPieChart'
    );

  if (pieCanvas) {

    charts.pie = new Chart(
      pieCanvas.getContext('2d'),
      {
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
      }
    );
  }

  // BAR CHART

  const barCanvas =
    document.getElementById(
      'reRegBarChart'
    );

  if (barCanvas) {

    charts.bar = new Chart(
      barCanvas.getContext('2d'),
      {
        type: 'bar',

        data: {
          labels: batches,

          datasets: [{
            label: 'Re-Reg Done',

            data: doneData,

            backgroundColor:
              '#4f46e5'
          }]
        }
      }
    );
  }

  // STACKED CHART

  const stackedCanvas =
    document.getElementById(
      'stackedBarChart'
    );

  if (stackedCanvas) {

    charts.stacked = new Chart(
      stackedCanvas.getContext('2d'),
      {
        type: 'bar',

        data: {
          labels: batches,

          datasets: [
            {
              label: 'Done',

              data: doneData,

              backgroundColor:
                '#10b981'
            },
            {
              label: 'Pending',

              data: pendingData,

              backgroundColor:
                '#ef4444'
            }
          ]
        },

        options: {
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

  // IA CHART

  const iaCanvas =
    document.getElementById(
      'iaBarChart'
    );

  if (iaCanvas) {

    charts.ia = new Chart(
      iaCanvas.getContext('2d'),
      {
        type: 'bar',

        data: {
          labels: batches,

          datasets: [
            {
              label: 'Submitted',

              data: batches.map(
                b =>
                  stats.batchMap[b]
                    .iaSubmitted
              ),

              backgroundColor:
                '#10b981'
            },
            {
              label:
                'Partially Submitted',

              data: batches.map(
                b =>
                  stats.batchMap[b]
                    .iaPartial
              ),

              backgroundColor:
                '#f59e0b'
            },
            {
              label: 'Nil Submission',

              data: batches.map(
                b =>
                  stats.batchMap[b]
                    .iaNil
              ),

              backgroundColor:
                '#ef4444'
            }
          ]
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
    '<th>IA Status</th></tr>';

  body.innerHTML = page.map(row => {

    return '<tr>' +

      allColumns.map(col => {

        const val =
          normalizeVal(row[col]);

        return `<td>${val}</td>`;

      }).join('')

      +

      `<td>${getIAStatus(row)}</td>`

      +

      '</tr>';

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
    .classList.toggle(
      'hidden',
      !show
    );
}
