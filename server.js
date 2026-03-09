const express = require('express');
const ExcelJS = require('exceljs');

const app = express();
app.use(express.json());

const PORT = process.env.PORT || 3000;

// ── API Credentials (env vars with dev defaults) ──────────────────────────────

const PROEST_BASE = process.env.PROEST_BASE_URL || 'https://cloud.proest.com/external_api/v1';
const PROEST_PARTNER_KEY = process.env.PROEST_PARTNER_KEY || 'tRUCaYu1HpRURb1geiM_';
const PROEST_COMPANY_KEY = process.env.PROEST_COMPANY_KEY || '_JVadJkZ-Wzh9so_Zdy2';

const BUILDR_TOKEN_URL = process.env.BUILDR_TOKEN_URL || 'https://buildr.app/oauth/token';
const BUILDR_BASE = process.env.BUILDR_BASE_URL || 'https://api.buildr.com/api/2023-01';
const BUILDR_CLIENT_ID = process.env.BUILDR_CLIENT_ID || 'hBCahZo4zFvE2qD58vnvRFOtjG0W00Ia3UaZCCerNJU';
const BUILDR_CLIENT_SECRET = process.env.BUILDR_CLIENT_SECRET || 'SwB-6dLuZXPf9j-PNAIg4QQRt_b1Cw-oY2SPOLYeI1w';

// ── Token caches ──────────────────────────────────────────────────────────────

let proestToken = null;
let buildrToken = null;
let buildrTokenExpiry = 0;

// ── ProEst Auth ───────────────────────────────────────────────────────────────

async function getProestToken(forceRefresh = false) {
  if (proestToken && !forceRefresh) return proestToken;
  const res = await fetch(`${PROEST_BASE}/login`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
    body: JSON.stringify({ partner_key: PROEST_PARTNER_KEY, company_key: PROEST_COMPANY_KEY }),
  });
  if (!res.ok) throw new Error(`ProEst login failed: ${res.status}`);
  const data = await res.json();
  proestToken = data.token;
  return proestToken;
}

async function proestFetch(path, retried = false) {
  const token = await getProestToken();
  const res = await fetch(`${PROEST_BASE}${path}`, {
    headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' },
  });
  if (res.status === 401 && !retried) {
    proestToken = null;
    return proestFetch(path, true);
  }
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    throw new Error(`ProEst API error ${res.status}: ${text}`);
  }
  return res.json();
}

// ── Buildr Auth ───────────────────────────────────────────────────────────────

async function getBuildrToken() {
  if (buildrToken && Date.now() < buildrTokenExpiry) return buildrToken;
  const res = await fetch(BUILDR_TOKEN_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      grant_type: 'client_credentials',
      client_id: BUILDR_CLIENT_ID,
      client_secret: BUILDR_CLIENT_SECRET,
      scope: 'read write',
    }),
  });
  if (!res.ok) throw new Error(`Buildr auth failed: ${res.status}`);
  const data = await res.json();
  buildrToken = data.access_token;
  buildrTokenExpiry = Date.now() + (data.expires_in || 3600) * 1000 - 60000;
  return buildrToken;
}

async function buildrFetch(path) {
  const token = await getBuildrToken();
  const res = await fetch(`${BUILDR_BASE}${path}`, {
    headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' },
  });
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    throw new Error(`Buildr API error ${res.status}: ${text}`);
  }
  return res.json();
}

// ── API Routes ────────────────────────────────────────────────────────────────

// Search ProEst estimates
app.get('/api/proest/estimates', async (req, res) => {
  try {
    const query = req.query.query;
    if (!query) return res.status(400).json({ error: 'query parameter required' });
    const data = await proestFetch(`/estimates?query=${encodeURIComponent(query)}`);
    res.json(data);
  } catch (err) {
    console.error('ProEst search error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// Get ProEst estimate detail
app.get('/api/proest/estimates/:id', async (req, res) => {
  try {
    const data = await proestFetch(`/estimates/${req.params.id}`);
    res.json(data);
  } catch (err) {
    console.error('ProEst detail error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// Get all Buildr projects (paginated server-side)
app.get('/api/buildr/projects', async (req, res) => {
  try {
    let allProjects = [];
    let page = 1;
    while (true) {
      const data = await buildrFetch(`/projects?per_page=100&page=${page}`);
      const projects = data.projects || data.data || data;
      if (!Array.isArray(projects) || projects.length === 0) break;
      allProjects = allProjects.concat(projects);
      if (projects.length < 100) break;
      page++;
    }
    // Filter out closed_cancelled
    // Only show active, pursuit, upcoming, and complete projects
    const showStatuses = new Set(['active', 'pursuit', 'upcoming', 'complete']);
    allProjects = allProjects.filter(p => showStatuses.has(p.project_status));
    // Sort alphabetically
    allProjects.sort((a, b) => (a.name || '').localeCompare(b.name || ''));
    res.json(allProjects);
  } catch (err) {
    console.error('Buildr projects error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── Transfer / Excel Generation ───────────────────────────────────────────────

function convertCode(code) {
  // DD.SSSS.IIII → DD-SS-SS-IIII
  if (!code) return '';
  const parts = code.split('.');
  if (parts.length !== 3) return code;
  const [div, sub, item] = parts;
  const s1 = sub.substring(0, 2);
  const s2 = sub.substring(2, 4);
  return `${div}-${s1}-${s2}-${item}`;
}

function transformItems(items) {
  const rows = [];
  for (const item of items) {
    // Sum cost categories
    const material = item.material?.total || 0;
    const labor = item.labor?.total || 0;
    const subcontractor = item.subcontractor?.total || 0;
    const equipment = item.equipment?.total || 0;
    const other = item.other?.total || 0;
    const totalCost = material + labor + subcontractor + equipment + other;

    // Skip $0 items
    if (totalCost === 0) continue;

    let itemCode = convertCode(item.code);

    // Division 70 → 50
    if (itemCode.startsWith('70')) {
      itemCode = '50' + itemCode.substring(2);
    }

    // Build notes from divisions array
    const notes = Array.isArray(item.divisions)
      ? item.divisions.map(d => d.description || d.name || '').filter(Boolean).join(' > ')
      : '';

    rows.push({
      itemCode,
      description: item.description || '',
      quantity: 1,
      unit: 'LS',
      totalDirectCost: totalCost,
      notes,
    });
  }
  return rows;
}

async function buildExcel(rows, estimateCode, estimateName) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Buildr Import');

  // Column widths
  sheet.columns = [
    { key: 'itemCode', width: 18 },
    { key: 'description', width: 40 },
    { key: 'quantity', width: 10 },
    { key: 'unit', width: 8 },
    { key: 'totalDirectCost', width: 18 },
    { key: 'notes', width: 30 },
  ];

  // Header row
  const headers = ['Item Code', 'Description', 'Quantity', 'Unit', 'Total Direct Cost', 'Notes'];
  const headerRow = sheet.addRow(headers);
  headerRow.eachCell(cell => {
    cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
  });

  // Data rows
  for (const row of rows) {
    const dataRow = sheet.addRow([
      row.itemCode,
      row.description,
      row.quantity,
      row.unit,
      row.totalDirectCost,
      row.notes,
    ]);
    dataRow.eachCell(cell => {
      cell.font = { name: 'Arial', size: 10 };
      cell.border = { bottom: { style: 'thin', color: { argb: 'FFB4C6E7' } } };
    });
    // Currency format for cost column (column 5)
    dataRow.getCell(5).numFmt = '$#,##0.00';
  }

  return workbook.xlsx.writeBuffer();
}

app.post('/api/transfer', async (req, res) => {
  try {
    const { estimate_id, project_name } = req.body;
    if (!estimate_id) return res.status(400).json({ error: 'estimate_id required' });

    // Fetch estimate detail
    const estimate = await proestFetch(`/estimates/${estimate_id}`);
    const items = estimate.items || [];
    const estimateName = estimate.description || estimate.name || 'Estimate';
    const estimateCode = estimate.code || String(estimate_id);

    if (items.length === 0) {
      return res.status(400).json({ error: 'No items found in estimate' });
    }

    const rows = transformItems(items);
    if (rows.length === 0) {
      return res.status(400).json({ error: 'All items have $0 total — nothing to transfer' });
    }

    const buffer = await buildExcel(rows, estimateCode, estimateName);
    const filename = `${estimateCode} - ${estimateName}_BUILDR.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(Buffer.from(buffer));
  } catch (err) {
    console.error('Transfer error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── HTML UI ───────────────────────────────────────────────────────────────────

const HTML = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ProEst to Buildr Transfer | Y&amp;C</title>
<style>
  :root {
    --navy: #1F4E79;
    --navy-light: #2A6BA3;
    --navy-dark: #163A5C;
    --accent: #B4C6E7;
    --bg: #F4F6F9;
    --card: #FFFFFF;
    --text: #2C3E50;
    --text-light: #6B7C93;
    --success: #27AE60;
    --error: #E74C3C;
    --border: #DCE3EB;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    background: var(--bg);
    color: var(--text);
    min-height: 100vh;
  }
  header {
    background: var(--navy);
    color: white;
    padding: 16px 24px;
    display: flex;
    align-items: center;
    gap: 16px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.15);
  }
  header h1 {
    font-size: 20px;
    font-weight: 600;
    letter-spacing: -0.3px;
  }
  header .subtitle {
    font-size: 13px;
    opacity: 0.7;
    font-weight: 400;
  }
  .container {
    max-width: 700px;
    margin: 32px auto;
    padding: 0 20px;
  }
  .card {
    background: var(--card);
    border-radius: 10px;
    padding: 28px;
    margin-bottom: 20px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
    border: 1px solid var(--border);
  }
  .card h2 {
    font-size: 15px;
    font-weight: 600;
    color: var(--navy);
    margin-bottom: 16px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
  }
  .input-row {
    display: flex;
    gap: 10px;
  }
  input[type="text"], .search-input {
    flex: 1;
    padding: 10px 14px;
    border: 1.5px solid var(--border);
    border-radius: 6px;
    font-size: 14px;
    outline: none;
    transition: border-color 0.2s;
  }
  input[type="text"]:focus, .search-input:focus {
    border-color: var(--navy-light);
  }
  button {
    padding: 10px 20px;
    border: none;
    border-radius: 6px;
    font-size: 14px;
    font-weight: 600;
    cursor: pointer;
    transition: background 0.2s, opacity 0.2s;
  }
  button:disabled {
    opacity: 0.5;
    cursor: not-allowed;
  }
  .btn-primary {
    background: var(--navy);
    color: white;
  }
  .btn-primary:hover:not(:disabled) {
    background: var(--navy-light);
  }
  .btn-success {
    background: var(--success);
    color: white;
    width: 100%;
    padding: 14px;
    font-size: 16px;
  }
  .btn-success:hover:not(:disabled) {
    background: #219A52;
  }
  .estimate-info {
    display: none;
    margin-top: 16px;
    padding: 16px;
    background: #F0F5FA;
    border-radius: 8px;
    border-left: 4px solid var(--navy);
  }
  .estimate-info .name {
    font-weight: 600;
    font-size: 16px;
    margin-bottom: 8px;
  }
  .estimate-info .meta {
    display: flex;
    gap: 24px;
    font-size: 13px;
    color: var(--text-light);
  }
  .estimate-info .meta span {
    font-weight: 600;
    color: var(--text);
  }
  .dropdown-wrapper {
    position: relative;
  }
  .dropdown-wrapper input {
    width: 100%;
  }
  .dropdown-list {
    display: none;
    position: absolute;
    top: 100%;
    left: 0;
    right: 0;
    max-height: 240px;
    overflow-y: auto;
    background: white;
    border: 1.5px solid var(--border);
    border-top: none;
    border-radius: 0 0 6px 6px;
    z-index: 10;
    box-shadow: 0 4px 12px rgba(0,0,0,0.1);
  }
  .dropdown-list.open { display: block; }
  .dropdown-item {
    padding: 10px 14px;
    font-size: 14px;
    cursor: pointer;
    border-bottom: 1px solid #F0F0F0;
  }
  .dropdown-item:hover, .dropdown-item.active {
    background: #F0F5FA;
  }
  .dropdown-item .project-status {
    font-size: 11px;
    color: var(--text-light);
    margin-left: 8px;
  }
  .status-bar {
    display: none;
    margin-top: 16px;
    padding: 12px 16px;
    border-radius: 6px;
    font-size: 14px;
    font-weight: 500;
  }
  .status-bar.info {
    display: block;
    background: #EBF5FB;
    color: var(--navy);
    border: 1px solid var(--accent);
  }
  .status-bar.success {
    display: block;
    background: #EAFAF1;
    color: var(--success);
    border: 1px solid #A9DFBF;
  }
  .status-bar.error {
    display: block;
    background: #FDEDEC;
    color: var(--error);
    border: 1px solid #F5B7B1;
  }
  .spinner {
    display: inline-block;
    width: 14px;
    height: 14px;
    border: 2px solid rgba(255,255,255,0.3);
    border-top-color: white;
    border-radius: 50%;
    animation: spin 0.6s linear infinite;
    vertical-align: middle;
    margin-right: 6px;
  }
  .spinner.dark {
    border-color: rgba(31,78,121,0.2);
    border-top-color: var(--navy);
  }
  @keyframes spin { to { transform: rotate(360deg); } }
  .hidden { display: none !important; }
  .step-num {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 24px;
    height: 24px;
    border-radius: 50%;
    background: var(--navy);
    color: white;
    font-size: 12px;
    font-weight: 700;
    margin-right: 8px;
    flex-shrink: 0;
  }
  .card h2 { display: flex; align-items: center; }
</style>
</head>
<body>

<header>
  <div>
    <h1>ProEst to Buildr Transfer</h1>
    <div class="subtitle">Yorke &amp; Curtis, Inc.</div>
  </div>
</header>

<div class="container">

  <!-- Step 1: ProEst Lookup -->
  <div class="card">
    <h2><span class="step-num">1</span> ProEst Estimate</h2>
    <div class="input-row">
      <input type="text" id="estimateCode" placeholder="Enter estimate code (e.g. 25001)" />
      <button class="btn-primary" id="lookupBtn" onclick="lookupEstimate()">Look Up</button>
    </div>
    <div class="estimate-info" id="estimateInfo">
      <div class="name" id="estimateName"></div>
      <div class="meta">
        <div>Items: <span id="estimateItems">0</span></div>
        <div>Total Cost: <span id="estimateCost">$0.00</span></div>
        <div>Code: <span id="estimateCodeDisplay"></span></div>
      </div>
    </div>
    <div class="status-bar" id="lookupStatus"></div>
  </div>

  <!-- Step 2: Buildr Project -->
  <div class="card">
    <h2><span class="step-num">2</span> Buildr Project</h2>
    <div class="dropdown-wrapper" id="dropdownWrapper">
      <input type="text" class="search-input" id="projectSearch"
             placeholder="Search projects..."
             oninput="filterProjects()"
             onfocus="openDropdown()"
             autocomplete="off" />
      <div class="dropdown-list" id="dropdownList"></div>
    </div>
    <div class="status-bar" id="projectStatus"></div>
  </div>

  <!-- Step 3: Transfer -->
  <div class="card">
    <h2><span class="step-num">3</span> Transfer</h2>
    <button class="btn-success" id="transferBtn" onclick="doTransfer()" disabled>
      Generate &amp; Download Excel
    </button>
    <div class="status-bar" id="transferStatus"></div>
  </div>

</div>

<script>
let estimateId = null;
let estimateData = null;
let selectedProject = null;
let allProjects = [];

// ── ProEst Lookup ─────────────────────────────────────────────────────────

async function lookupEstimate() {
  const code = document.getElementById('estimateCode').value.trim();
  if (!code) return;

  const btn = document.getElementById('lookupBtn');
  const status = document.getElementById('lookupStatus');
  const info = document.getElementById('estimateInfo');

  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span>Looking up...';
  info.style.display = 'none';
  setStatus(status, 'info', 'Searching ProEst...');

  try {
    const res = await fetch('/api/proest/estimates?query=' + encodeURIComponent(code));
    if (!res.ok) throw new Error((await res.json()).error || 'Search failed');
    const estimates = await res.json();

    // Find matching estimate
    const list = estimates.estimates || (Array.isArray(estimates) ? estimates : (estimates.data || []));
    if (list.length === 0) {
      setStatus(status, 'error', 'No estimates found for "' + code + '"');
      btn.disabled = false;
      btn.textContent = 'Look Up';
      return;
    }

    // Take first match
    const est = list[0];
    estimateId = est.id;

    setStatus(status, 'info', 'Loading estimate details...');

    // Get full detail
    const detailRes = await fetch('/api/proest/estimates/' + est.id);
    if (!detailRes.ok) throw new Error((await detailRes.json()).error || 'Detail fetch failed');
    estimateData = await detailRes.json();

    const items = estimateData.items || [];
    const name = estimateData.description || estimateData.name || 'Unknown';
    const estCode = estimateData.code || code;

    // Calculate total cost
    let totalCost = 0;
    for (const item of items) {
      totalCost += (item.material?.total || 0)
        + (item.labor?.total || 0)
        + (item.subcontractor?.total || 0)
        + (item.equipment?.total || 0)
        + (item.other?.total || 0);
    }

    document.getElementById('estimateName').textContent = name;
    document.getElementById('estimateItems').textContent = items.length;
    document.getElementById('estimateCost').textContent = '$' + totalCost.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2});
    document.getElementById('estimateCodeDisplay').textContent = estCode;
    info.style.display = 'block';
    setStatus(status, 'success', 'Estimate loaded successfully');
    updateTransferBtn();
  } catch (err) {
    setStatus(status, 'error', err.message);
  }

  btn.disabled = false;
  btn.textContent = 'Look Up';
}

// ── Buildr Projects ───────────────────────────────────────────────────────

async function loadProjects() {
  const status = document.getElementById('projectStatus');
  setStatus(status, 'info', '<span class="spinner dark"></span> Loading Buildr projects...');

  try {
    const res = await fetch('/api/buildr/projects');
    if (!res.ok) throw new Error((await res.json()).error || 'Failed to load projects');
    allProjects = await res.json();
    setStatus(status, 'success', allProjects.length + ' projects loaded');
    setTimeout(() => { status.className = 'status-bar'; }, 2000);
  } catch (err) {
    setStatus(status, 'error', 'Failed to load projects: ' + err.message);
  }
}

function filterProjects() {
  const query = document.getElementById('projectSearch').value.toLowerCase();
  const list = document.getElementById('dropdownList');
  const filtered = allProjects.filter(p =>
    (p.name || '').toLowerCase().includes(query)
  ).slice(0, 50);

  list.innerHTML = '';
  for (const p of filtered) {
    const div = document.createElement('div');
    div.className = 'dropdown-item';
    div.innerHTML = p.name + (p.status ? '<span class="project-status">' + p.status + '</span>' : '');
    div.onclick = () => selectProject(p);
    list.appendChild(div);
  }
  list.classList.add('open');
}

function openDropdown() {
  filterProjects();
}

function selectProject(p) {
  selectedProject = p;
  document.getElementById('projectSearch').value = p.name;
  document.getElementById('dropdownList').classList.remove('open');
  updateTransferBtn();
}

// Close dropdown on outside click
document.addEventListener('click', (e) => {
  const wrapper = document.getElementById('dropdownWrapper');
  if (!wrapper.contains(e.target)) {
    document.getElementById('dropdownList').classList.remove('open');
  }
});

// Allow Enter key on estimate code
document.getElementById('estimateCode').addEventListener('keydown', (e) => {
  if (e.key === 'Enter') lookupEstimate();
});

// ── Transfer ──────────────────────────────────────────────────────────────

function updateTransferBtn() {
  document.getElementById('transferBtn').disabled = !(estimateId && selectedProject);
}

async function doTransfer() {
  const btn = document.getElementById('transferBtn');
  const status = document.getElementById('transferStatus');

  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Generating Excel...';
  setStatus(status, 'info', 'Building spreadsheet...');

  try {
    const res = await fetch('/api/transfer', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        estimate_id: estimateId,
        project_name: selectedProject.name,
      }),
    });

    if (!res.ok) {
      const err = await res.json();
      throw new Error(err.error || 'Transfer failed');
    }

    // Extract filename from Content-Disposition header
    const disposition = res.headers.get('Content-Disposition') || '';
    const filenameMatch = disposition.match(/filename="(.+?)"/);
    const filename = filenameMatch ? filenameMatch[1] : 'buildr_import.xlsx';

    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);

    setStatus(status, 'success', 'Download started: ' + filename);
  } catch (err) {
    setStatus(status, 'error', err.message);
  }

  btn.disabled = false;
  btn.innerHTML = 'Generate &amp; Download Excel';
  updateTransferBtn();
}

// ── Helpers ───────────────────────────────────────────────────────────────

function setStatus(el, type, msg) {
  el.className = 'status-bar ' + type;
  el.innerHTML = msg;
}

// Load projects on page load
loadProjects();
</script>

</body>
</html>`;

// ── Serve HTML ────────────────────────────────────────────────────────────────

app.get('/', (req, res) => {
  res.setHeader('Content-Type', 'text/html');
  res.send(HTML);
});

// ── Start ─────────────────────────────────────────────────────────────────────

app.listen(PORT, () => {
  console.log(`ProEst to Buildr Transfer running on http://localhost:${PORT}`);
});
