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
   COLUMN DETECTION
============================================================ */

function findColumn(possibleNames) {

  return allColumns.find(col => {

    const lower = col.toLowerCase();

    return possibleNames.some(name =>
      lower.includes(name)
    );
  });
}

function getColumnValue(row, names) {

  const col = findColumn(names);

  if (!col) return '';

  return normalizeVal(row[col]);
}

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

      alert(err.message);
    }
  };

  reader.readAsArrayBuffer(file);
}

/* ============================================================
   PARSE WORKBOOK
============================================================ */

function parseWorkbook(workbook, fileName) {

  let bestSheet = null;
  let maxRows = 0;

  workbook.SheetNames.forEach(name => {

    const sheet =
      workbook.Sheets[name];

    const json =
      XLSX.utils.sheet_to_json(sheet, {
        defval: ''
      });

    if (json.length > maxRows) {

      maxRows = json.length;
      bestSheet = sheet;
    }
  });

  rawData =
    XLSX.utils.sheet_to_json(bestSheet, {
      defval: ''
    });

  allColumns =
    Object.keys(rawData[0]);

  filteredData = [...rawData];

  populateFilters();

  renderDashboard();

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

  if (
    val === null ||
    val === undefined
  ) return '';

  return String(val).trim();
}

function pct(num, den) {

  if (!den) return 0;

  return Math.round(
    (num / den) * 1000
  ) / 10;
}

/* ============================================================
   RE-REG STATUS
============================================================ */

function isReRegDone(row) {

  const values =
    Object.values(row);

  const last =
    String(values[values.length - 1] || '')
      .trim()
      .toLowerCase();

  return [
    'done',
    'completed',
    'complete',
    'yes',
    'success',
    'registered',
    '1',
    'true'
  ].includes(last);
}

/* ============================================================
   IA STATUS
============================================================ */

function getIAStatus(row) {

  const val = getColumnValue(
    row,
    ['sem 1 ia', 'ia']
  ).toLowerCase();

  if (val.includes('partial'))
    return 'Partial';

  if (val.includes('submit'))
    return 'Submitted';

  return 'Nil';
}

/* ============================================================
   FILTERS
============================================================ */

function populateFilters() {

  populateFilter(
    'filterBatch',
    ['batch']
  );

  populateFilter(
    'filterProgram',
    ['program']
  );

  populateFilter(
    'filterSource',
    ['source']
  );

  populateFilter(
    'filterSalesType',
    ['sales']
  );

  populateFilter(
    'filterUGC',
    ['ugc']
  );
}

function populateFilter(id, names) {

  const select =
    document.getElementById(id);

  if (!select) return;

  select.innerHTML =
    '<option value="ALL">All</option>';

  const values = [
    ...new Set(
      rawData
        .map(r =>
          getColumnValue(r, names)
        )
        .filter(Boolean)
    )
  ];

  values.sort().forEach(v => {

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

  const sales =
    document.getElementById('filterSalesType')?.value || 'ALL';

  const ugc =
    document.getElementById('filterUGC')?.value || 'ALL';

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
      sales !== 'ALL' &&
      getColumnValue(row, ['sales']) !== sales
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

      const txt =
        Object.values(row)
          .join(' ')
          .toLowerCase();

      if (!txt.includes(search))
        return false;
    }

    return true;
  });

  renderDashboard();
}

/* ============================================================
   EVENTS
============================================================ */

[
  'filterBatch',
  'filterProgram',
  'filterStatus',
  'filterSource',
  'filterSalesType',
  'filterUGC'
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

function buildSummaryMap(type) {

  const map = {};

  filteredData.forEach(row => {

    const key =
      getColumnValue(
        row,
        [type]
      ) || 'Unknown';

    if (!map[key]) {

      map[key] = {

        total: 0,
        done: 0,
        iaSubmitted: 0,
        iaPartial: 0,
        iaNil: 0
      };
    }

    map[key].total++;

    if (isReRegDone(row))
      map[key].done++;

    const ia =
      getIAStatus(row);

    if (ia === 'Submitted')
      map[key].iaSubmitted++;

    else if (ia === 'Partial')
      map[key].iaPartial++;

    else
      map[key].iaNil++;
  });

  return map;
}

/* ============================================================
   RENDER
============================================================ */

function renderDashboard() {

  const batchMap =
    buildSummaryMap('batch');

  const programMap =
    buildSummaryMap('program');

  renderBatchTable(batchMap);

  renderProgramTable(programMap);

  renderCharts(batchMap);

  renderDetailTable();
}

/* ============================================================
   BATCH TABLE
============================================================ */

function renderBatchTable(map) {

  const body =
    document.getElementById(
      'batchSummaryBody'
    );

  let html = '';

  Object.entries(map).forEach(
    ([batch, d]) => {

      const pending =
        d.total - d.done;

      const rrPct =
        pct(d.done, d.total);

      const iaPct =
        pct(d.iaSubmitted, d.total);

      html += `
        <tr>
          <td>${batch}</td>
          <td>${d.total}</td>
          <td>${d.done}</td>
          <td>${pending}</td>
          <td>${rrPct}%</td>
          <td>${d.iaSubmitted}</td>
          <td>${iaPct}%</td>
          <td>${d.iaPartial}</td>
          <td>${d.iaNil}</td>
        </tr>
      `;
    }
  );

  body.innerHTML = html;
}

/* ============================================================
   PROGRAM TABLE
============================================================ */

function renderProgramTable(map) {

  const body =
    document.getElementById(
      'programSummaryBody'
    );

  if (!body) return;

  let html = '';

  Object.entries(map).forEach(
    ([program, d]) => {

      const pending =
        d.total - d.done;

      const rrPct =
        pct(d.done, d.total);

      const iaPct =
        pct(d.iaSubmitted, d.total);

      html += `
        <tr>
          <td>${program}</td>
          <td>${d.total}</td>
          <td>${d.done}</td>
          <td>${pending}</td>
          <td>${rrPct}%</td>
          <td>${d.iaSubmitted}</td>
          <td>${iaPct}%</td>
          <td>${d.iaPartial}</td>
          <td>${d.iaNil}</td>
        </tr>
      `;
    }
  );

  body.innerHTML = html;
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

function renderCharts(batchMap) {

  destroyCharts();

  const labels =
    Object.keys(batchMap);

  const done =
    labels.map(
      l => batchMap[l].done
    );

  const pending =
    labels.map(
      l =>
        batchMap[l].total -
        batchMap[l].done
    );

  charts.bar = new Chart(
    document
      .getElementById('reRegBarChart')
      .getContext('2d'),
    {
      type: 'bar',

      data: {
        labels,

        datasets: [{
          label: 'Done',
          data: done,
          backgroundColor: '#4f46e5'
        }]
      }
    }
  );

  charts.pie = new Chart(
    document
      .getElementById('reRegPieChart')
      .getContext('2d'),
    {
      type: 'doughnut',

      data: {
        labels: ['Done', 'Pending'],

        datasets: [{
          data: [
            done.reduce((a,b)=>a+b,0),
            pending.reduce((a,b)=>a+b,0)
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

  head.innerHTML =
    '<tr>' +
    allColumns.map(c =>
      `<th>${c}</th>`
    ).join('') +
    '</tr>';

  body.innerHTML =
    filteredData
      .slice(0, PAGE_SIZE)
      .map(row => {

        return '<tr>' +

          allColumns.map(col => {

            return `
              <td>
                ${normalizeVal(row[col])}
              </td>
            `;

          }).join('')

          +

          '</tr>';
      })
      .join('');
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
