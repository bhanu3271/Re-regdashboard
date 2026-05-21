'use strict';

/* ============================================================
   GLOBAL STATE
============================================================ */

let rawData = [];
let filteredData = [];
let allColumns = [];
let charts = {};
let currentPage = 1;

const PAGE_SIZE = 25;

/* ============================================================
   SAFE ELEMENT
============================================================ */

function el(id) {
  return document.getElementById(id);
}

/* ============================================================
   INIT
============================================================ */

document.addEventListener('DOMContentLoaded', () => {

  const fileInput = el('fileInput');

  if (fileInput) {
    fileInput.addEventListener('change', handleFile);
  }

  setupDragDrop();

  [
    'filterBatch',
    'filterProgram',
    'filterSource',
    'filterSalesType',
    'filterUGC',
    'filterIAStatus',
    'filterStatus'
  ].forEach(id => {

    const element = el(id);

    if (element) {
      element.addEventListener('change', applyFilters);
    }
  });

  const searchInput = el('searchInput');

  if (searchInput) {
    searchInput.addEventListener('input', applyFilters);
  }

});

/* ============================================================
   DRAG DROP
============================================================ */

function setupDragDrop() {

  const uploadCard =
    document.querySelector('.upload-card');

  if (!uploadCard) return;

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

    if (file) {
      processFile(file);
    }
  });
}

/* ============================================================
   FILE HANDLING
============================================================ */

function handleFile(e) {

  const file = e.target.files[0];

  if (file) {
    processFile(file);
  }
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

  allColumns =
    Object.keys(rawData[0] || {});

  filteredData = [...rawData];

  populateFilters();

  renderDashboard();

  const lu = el('lastUpdated');

  if (lu) {

    lu.classList.remove('hidden');

    const span = lu.querySelector('span');

    if (span) {

      span.textContent =
        `Loaded: ${fileName} · ${rawData.length} rows`;
    }
  }

  const uploadZone = el('uploadZone');
  const dashboard = el('dashboard');

  if (uploadZone) {
    uploadZone.classList.add('hidden');
  }

  if (dashboard) {
    dashboard.classList.remove('hidden');
  }

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

  return (
    Math.round((num / den) * 1000) / 10
  );
}

/* ============================================================
   GET VALUE FROM COLUMN
============================================================ */

function getValue(row, possibleNames) {

  for (const key of Object.keys(row)) {

    const lowerKey =
      key.toLowerCase().trim();

    for (const name of possibleNames) {

      if (
        lowerKey.includes(
          name.toLowerCase()
        )
      ) {

        return normalizeVal(row[key]);
      }
    }
  }

  return '';
}

/* ============================================================
   RE-REG STATUS
============================================================ */

function isReRegDone(row) {

  const values = Object.values(row);

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

  const iaValue = String(
    getValue(row, [
      'sem 1 ia',
      'ia status',
      'ia'
    ])
  ).toLowerCase();

  if (
    iaValue.includes('partial')
  ) {
    return 'PARTIAL';
  }

  if (
    iaValue.includes('nil') ||
    iaValue.includes('not submitted')
  ) {
    return 'NIL';
  }

  return 'SUBMITTED';
}

/* ============================================================
   FILTERS
============================================================ */

function populateFilters() {

  populateSelect(
    'filterBatch',
    getUniqueValues([
      'batch'
    ])
  );

  populateSelect(
    'filterProgram',
    getUniqueValues([
      'program'
    ])
  );

  populateSelect(
    'filterSource',
    getUniqueValues([
      'source'
    ])
  );

  populateSelect(
    'filterSalesType',
    getUniqueValues([
      'sales type'
    ])
  );

  /* UGC FIX */

  populateSelect(
    'filterUGC',
    getUniqueValues([
      'ugc included',
      'ugc',
      'is ugc'
    ])
  );
}

function getUniqueValues(names) {

  const values = [];

  rawData.forEach(row => {

    const value =
      normalizeVal(
        getValue(row, names)
      );

    if (value) {
      values.push(value);
    }
  });

  return [...new Set(values)].sort();
}

function populateSelect(id, values) {

  const select = el(id);

  if (!select) return;

  select.innerHTML =
    `<option value="ALL">All</option>`;

  values.forEach(value => {

    const option =
      document.createElement('option');

    option.value = value;
    option.textContent = value;

    select.appendChild(option);
  });
}

/* ============================================================
   APPLY FILTERS
============================================================ */

function applyFilters() {

  filteredData = rawData.filter(row => {

    const batch =
      el('filterBatch')?.value || 'ALL';

    const program =
      el('filterProgram')?.value || 'ALL';

    const source =
      el('filterSource')?.value || 'ALL';

    const sales =
      el('filterSalesType')?.value || 'ALL';

    const ugc =
      el('filterUGC')?.value || 'ALL';

    const iaStatus =
      el('filterIAStatus')?.value || 'ALL';

    const status =
      el('filterStatus')?.value || 'ALL';

    const search =
      (
        el('searchInput')?.value || ''
      ).toLowerCase();

    if (
      batch !== 'ALL' &&
      getValue(row, ['batch']) !== batch
    ) {
      return false;
    }

    if (
      program !== 'ALL' &&
      getValue(row, ['program']) !== program
    ) {
      return false;
    }

    if (
      source !== 'ALL' &&
      getValue(row, ['source']) !== source
    ) {
      return false;
    }

    if (
      sales !== 'ALL' &&
      getValue(row, ['sales type']) !== sales
    ) {
      return false;
    }

    /* UGC FIX */

    const ugcValue =
      getValue(row, [
        'ugc included',
        'ugc',
        'is ugc'
      ]);

    if (
      ugc !== 'ALL' &&
      ugcValue !== ugc
    ) {
      return false;
    }

    if (
      iaStatus !== 'ALL' &&
      getIAStatus(row) !== iaStatus
    ) {
      return false;
    }

    if (
      status === 'DONE' &&
      !isReRegDone(row)
    ) {
      return false;
    }

    if (
      status === 'PENDING' &&
      isReRegDone(row)
    ) {
      return false;
    }

    if (search) {

      const rowText =
        Object.values(row)
          .join(' ')
          .toLowerCase();

      if (
        !rowText.includes(search)
      ) {
        return false;
      }
    }

    return true;
  });

  currentPage = 1;

  renderDashboard();
}

/* ============================================================
   RESET FILTERS
============================================================ */

function resetFilters() {

  [
    'filterBatch',
    'filterProgram',
    'filterSource',
    'filterSalesType',
    'filterUGC',
    'filterIAStatus',
    'filterStatus'
  ].forEach(id => {

    const element = el(id);

    if (element) {
      element.value = 'ALL';
    }
  });

  if (el('searchInput')) {
    el('searchInput').value = '';
  }

  filteredData = [...rawData];

  renderDashboard();
}

/* ============================================================
   DASHBOARD RENDER
============================================================ */

function renderDashboard() {

  renderKPIs();

  renderBatchSummary();

  renderProgramSummary();

  renderCharts();

  renderDetailTable();
}

/* ============================================================
   KPI
============================================================ */

function renderKPIs() {

  const total =
    filteredData.length;

  const done =
    filteredData.filter(
      isReRegDone
    ).length;

  const pending =
    total - done;

  const grid = el('kpiGrid');

  if (!grid) return;

  grid.innerHTML = `

    <div class="kpi-card">
      <div class="kpi-label">
        Total Students
      </div>
      <div class="kpi-value">
        ${total}
      </div>
    </div>

    <div class="kpi-card">
      <div class="kpi-label">
        Re-Reg Done
      </div>
      <div class="kpi-value">
        ${done}
      </div>
    </div>

    <div class="kpi-card">
      <div class="kpi-label">
        Pending
      </div>
      <div class="kpi-value">
        ${pending}
      </div>
    </div>

    <div class="kpi-card">
      <div class="kpi-label">
        Completion %
      </div>
      <div class="kpi-value">
        ${pct(done, total)}%
      </div>
    </div>
  `;
}

/* ============================================================
   BATCH SUMMARY
============================================================ */

function renderBatchSummary() {

  const body =
    el('batchSummaryBody');

  if (!body) return;

  const groups = {};

  filteredData.forEach(row => {

    const batch =
      getValue(row, ['batch']) ||
      'Unknown';

    if (!groups[batch]) {
      groups[batch] = [];
    }

    groups[batch].push(row);
  });

  body.innerHTML = '';

  Object.keys(groups).forEach(batch => {

    const rows = groups[batch];

    renderSummaryRow(
      body,
      batch,
      rows
    );
  });
}

/* ============================================================
   PROGRAM SUMMARY
============================================================ */

function renderProgramSummary() {

  const body =
    el('programSummaryBody');

  if (!body) return;

  const groups = {};

  filteredData.forEach(row => {

    const program =
      getValue(row, ['program']) ||
      'Unknown';

    if (!groups[program]) {
      groups[program] = [];
    }

    groups[program].push(row);
  });

  body.innerHTML = '';

  Object.keys(groups).forEach(program => {

    const rows = groups[program];

    renderSummaryRow(
      body,
      program,
      rows
    );
  });
}

/* ============================================================
   SUMMARY ROW
============================================================ */

function renderSummaryRow(
  body,
  label,
  rows
) {

  const total = rows.length;

  const done =
    rows.filter(
      isReRegDone
    ).length;

  const pending =
    total - done;

  const submitted =
    rows.filter(r =>
      getIAStatus(r) === 'SUBMITTED'
    ).length;

  const partial =
    rows.filter(r =>
      getIAStatus(r) === 'PARTIAL'
    ).length;

  const nil =
    rows.filter(r =>
      getIAStatus(r) === 'NIL'
    ).length;

  const tr =
    document.createElement('tr');

  tr.innerHTML = `

    <td>${label}</td>

    <td>${total}</td>

    <td>${done}</td>

    <td>${pending}</td>

    <td>${pct(done, total)}%</td>

    <td>
      ${submitted}
      (${pct(submitted, total)}%)
    </td>

    <td>
      ${partial}
      (${pct(partial, total)}%)
    </td>

    <td>
      ${nil}
      (${pct(nil, total)}%)
    </td>

    <td>
      ${pct(submitted, total)}%
    </td>

    <td>
      ${
        pct(done, total) >= 75
          ? 'Good'
          : 'Needs Attention'
      }
    </td>
  `;

  body.appendChild(tr);
}

/* ============================================================
   CHARTS
============================================================ */

function destroyCharts() {

  Object.values(charts).forEach(chart => {

    if (chart) {
      chart.destroy();
    }
  });

  charts = {};
}

function renderCharts() {

  destroyCharts();

  const groups = {};

  filteredData.forEach(row => {

    const batch =
      getValue(row, ['batch']) ||
      'Unknown';

    if (!groups[batch]) {
      groups[batch] = [];
    }

    groups[batch].push(row);
  });

  const labels =
    Object.keys(groups);

  const reRegData = [];
  const iaData = [];
  const doneData = [];
  const pendingData = [];

  labels.forEach(batch => {

    const rows = groups[batch];

    const total =
      rows.length;

    const done =
      rows.filter(
        isReRegDone
      ).length;

    const submitted =
      rows.filter(r =>
        getIAStatus(r) === 'SUBMITTED'
      ).length;

    reRegData.push(
      pct(done, total)
    );

    iaData.push(
      pct(submitted, total)
    );

    doneData.push(done);

    pendingData.push(
      total - done
    );
  });

  /* RE-REG BAR */

  const reRegCanvas =
    el('reRegBarChart');

  if (reRegCanvas) {

    charts.bar = new Chart(
      reRegCanvas,
      {
        type: 'bar',
        data: {
          labels,
          datasets: [{
            label: 'Re-Reg %',
            data: reRegData,
            backgroundColor:
              '#4f46e5'
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
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

  /* IA BAR */

  const iaCanvas =
    el('iaBarChart');

  if (iaCanvas) {

    charts.ia = new Chart(
      iaCanvas,
      {
        type: 'bar',
        data: {
          labels,
          datasets: [{
            label: 'IA %',
            data: iaData,
            backgroundColor:
              '#10b981'
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
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

  /* PIE */

  const pieCanvas =
    el('reRegPieChart');

  if (pieCanvas) {

    const done =
      filteredData.filter(
        isReRegDone
      ).length;

    const pending =
      filteredData.length - done;

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
              done,
              pending
            ],
            backgroundColor: [
              '#10b981',
              '#ef4444'
            ]
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false
        }
      }
    );
  }

  /* STACKED */

  const stackedCanvas =
    el('stackedBarChart');

  if (stackedCanvas) {

    charts.stacked =
      new Chart(
        stackedCanvas,
        {
          type: 'bar',
          data: {
            labels,
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
            responsive: true,
            maintainAspectRatio: false,
            scales: {
              x: {
                stacked: true
              },
              y: {
                stacked: true,
                beginAtZero: true
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
    el('detailTableHead');

  const body =
    el('detailTableBody');

  const meta =
    el('tableMeta');

  if (!head || !body) return;

  head.innerHTML =
    '<tr>' +
    allColumns.map(col =>
      `<th>${col}</th>`
    ).join('') +
    '</tr>';

  const start =
    (currentPage - 1) *
    PAGE_SIZE;

  const end =
    start + PAGE_SIZE;

  const pageData =
    filteredData.slice(start, end);

  body.innerHTML =
    pageData.map(row => {

      return `
        <tr>
          ${allColumns.map(col => `
            <td>
              ${normalizeVal(row[col])}
            </td>
          `).join('')}
        </tr>
      `;
    }).join('');

  if (meta) {

    meta.textContent =
      `Showing ${
        start + 1
      } - ${
        Math.min(
          end,
          filteredData.length
        )
      } of ${
        filteredData.length
      }`;
  }
}

/* ============================================================
   EXPORT
============================================================ */

function exportFilteredData() {

  const ws =
    XLSX.utils.json_to_sheet(
      filteredData
    );

  const wb =
    XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    wb,
    ws,
    'Filtered Data'
  );

  XLSX.writeFile(
    wb,
    'Filtered_Data.xlsx'
  );
}

/* ============================================================
   LOADING
============================================================ */

function showLoading(show) {

  const loading =
    el('loadingScreen');

  if (!loading) return;

  loading.classList.toggle(
    'hidden',
    !show
  );
}
