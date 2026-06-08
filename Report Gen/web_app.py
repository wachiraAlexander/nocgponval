"""
================================================================================
  REPORT GENERATOR  -  WEB APP  (Flask)
================================================================================

Browser-based version of Report Gen. Reuses ALL report-generation
logic from app.py (pipelines, styling, adaptive packing, assignee cleaning).
Only the UI layer is different.

USAGE
    pip install flask pandas openpyxl
    python web_app.py

    Then on the same machine:        http://localhost:5000
    From other machines on subnet:   http://<server-ip>:5000

DEPLOYMENT
    The server binds to 0.0.0.0 so it accepts connections from any machine
    on your local network. Find your IP with:
        Windows:  ipconfig
        macOS:    ipconfig getifaddr en0
        Linux:    hostname -I

    Open the firewall on port 5000 if needed.

DESIGN
    Stateless — uploaded files live only for the duration of one request.
    Output files are streamed back as a download immediately and not kept
    on disk. No login. Each user works on their own file independently.

Authors: Emmanuel Mutua & Alex Wachira
"""

from __future__ import annotations

import io
import logging
import socket
import sys
import tempfile
import traceback
from datetime import datetime
from pathlib import Path
from typing import Dict

# Reuse the entire report-gen layer from the desktop app.
# This relies on app.py being in the same directory.
sys.path.insert(0, str(Path(__file__).resolve().parent))
import app as report_lib  # noqa: E402

from flask import (  # noqa: E402
    Flask,
    Response,
    flash,
    jsonify,
    redirect,
    render_template_string,
    request,
    send_file,
    url_for,
)
from werkzeug.utils import secure_filename  # noqa: E402

# ----------------------------------------------------------------------------
# Configuration
# ----------------------------------------------------------------------------
HOST = "0.0.0.0"          # listen on all interfaces (so subnet can reach us)
PORT = 5000
MAX_UPLOAD_MB = 50         # generous limit for big Issue.xlsx files
ALLOWED_EXTENSIONS = {".xlsx", ".xlsm"}

app_flask = Flask(__name__)
app_flask.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024
app_flask.secret_key = "report-gen-local-only"  # only used for flash messages

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s  %(levelname)-7s  %(message)s")
log = logging.getLogger("report-web")


# ----------------------------------------------------------------------------
# HTML template (single file, inline; styled to match the desktop app's vibe)
# ----------------------------------------------------------------------------
PAGE_HTML = r"""
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Report Gen — JTL NOC Monitoring</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
  :root {
    --navy: #1F3864;
    --navy-2: #2E5597;
    --navy-3: #14274A;
    --bg: #F4F6FB;
    --card: #FFFFFF;
    --border: #DCE3EE;
    --border-strong: #B6C2D5;
    --muted: #6A7282;
    --label: #595F6D;
    --green: #2E7D32;
    --red: #C62828;
    --orange: #ED7D31;
    --gpon-blue: #4472C4;
    --ent-purple: #7030A0;
    --ring-green: #70AD47;
    --gold: #B8860B;
    --shade-row: #F8FAFE;
    --bb-color: #4472C4;
    --ran-color: #ED7D31;
    --ring-color: #70AD47;
    --power-color: #C00000;
    --equip-color: #7030A0;
    --cesr-color: #808080;
    --shadow-sm: 0 1px 2px rgba(20, 30, 60, 0.06);
    --shadow-md: 0 2px 8px rgba(20, 30, 60, 0.08);
  }
  * { box-sizing: border-box; }
  html, body { margin: 0; padding: 0; }
  body {
    font-family: -apple-system, "Segoe UI", system-ui, "Helvetica Neue", sans-serif;
    background: var(--bg);
    color: #1A2233;
    font-size: 14px;
    line-height: 1.4;
  }

  /* ---- Header ---- */
  header {
    background: var(--navy);
    color: white;
    padding: 14px 0;
    box-shadow: var(--shadow-md);
  }
  .header-inner {
    max-width: 1180px;
    margin: 0 auto;
    padding: 0 20px;
    display: flex;
    align-items: center;
    justify-content: space-between;
  }
  header h1 {
    margin: 0;
    font-size: 18px;
    font-weight: 700;
    letter-spacing: .2px;
  }
  header h1 .accent { opacity: .65; font-weight: 400; margin-left: 6px; }
  header .meta { font-size: 12px; opacity: .85; }

  /* ---- Layout ---- */
  main {
    max-width: 1180px;
    margin: 0 auto;
    padding: 22px 20px 60px;
  }
  .tabs {
    display: flex;
    gap: 4px;
    margin-bottom: 0;
  }
  .tab {
    background: #E5EAF3;
    border: 1px solid var(--border);
    border-bottom: none;
    border-radius: 8px 8px 0 0;
    padding: 11px 22px;
    font-weight: 600;
    color: #4B5567;
    cursor: pointer;
    user-select: none;
    font-size: 14px;
    transition: background .12s;
  }
  .tab:hover:not(.active) { background: #DCE3EE; }
  .tab.active {
    background: var(--card);
    color: var(--navy);
    border-bottom: 1px solid var(--card);
    margin-bottom: -1px;
    position: relative;
    z-index: 1;
  }
  .tab.gpon.active { box-shadow: inset 0 4px 0 var(--gpon-blue); }
  .tab.ent.active  { box-shadow: inset 0 4px 0 var(--ent-purple); }

  .panel {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 0 12px 12px 12px;
    padding: 20px;
    box-shadow: var(--shadow-sm);
  }

  /* ---- Section: drop zone ---- */
  .section {
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 14px 16px;
    margin-bottom: 14px;
    background: #FBFCFE;
  }
  .section-title {
    font-size: 11px;
    font-weight: 700;
    color: var(--label);
    text-transform: uppercase;
    letter-spacing: .6px;
    margin-bottom: 10px;
  }
  .drop {
    border: 2px dashed var(--border-strong);
    border-radius: 8px;
    padding: 22px 20px;
    text-align: center;
    background: #F2F6FC;
    color: var(--navy);
    cursor: pointer;
    transition: all .15s ease;
  }
  .drop:hover, .drop.hover {
    border-color: var(--navy-2);
    background: #EAF1FB;
  }
  .drop strong { display: block; font-size: 14px; margin-bottom: 4px; }
  .drop small { color: var(--muted); font-size: 12px; }
  .drop input[type=file] { display: none; }

  .file-row {
    display: flex;
    align-items: center;
    gap: 12px;
    margin-top: 10px;
    padding: 10px 12px;
    background: #EAF2EA;
    border-left: 4px solid var(--green);
    border-radius: 4px;
    font-size: 13px;
    display: none;
  }
  .file-row.show { display: flex; }
  .file-row .name { font-weight: 600; color: #1A2233; }
  .file-row .size { color: var(--muted); font-size: 12px; }
  .file-row .clear {
    margin-left: auto;
    background: none;
    border: none;
    color: var(--muted);
    cursor: pointer;
    font-size: 13px;
    padding: 4px 8px;
    border-radius: 4px;
  }
  .file-row .clear:hover { background: #DCE3EE; color: #1A2233; }

  /* ---- Two-column layout for preview ---- */
  .preview-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 14px;
    margin-top: 0;
    display: none;
  }
  .preview-grid.show { display: grid; }

  /* ---- KPI cards ---- */
  .kpis {
    display: grid;
    grid-template-columns: repeat(5, 1fr);
    gap: 6px;
    margin-bottom: 12px;
  }
  .kpi {
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 10px 6px;
    text-align: center;
    background: #FFFFFF;
  }
  .kpi .label {
    font-size: 9px;
    font-weight: 700;
    color: var(--label);
    text-transform: uppercase;
    letter-spacing: .5px;
    line-height: 1.2;
    height: 22px;
    display: flex;
    align-items: center;
    justify-content: center;
  }
  .kpi .value {
    font-size: 22px;
    font-weight: 700;
    color: var(--navy);
    line-height: 1.1;
    margin-top: 4px;
  }
  .kpi.empty .value { color: var(--border-strong); }

  /* ---- Source sheet info panel ---- */
  .info-panel {
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 12px 14px;
    background: #FBFCFE;
    flex: 1;
  }
  .info-panel h4 {
    margin: 0 0 8px;
    font-size: 12px;
    font-weight: 700;
    color: var(--label);
    text-transform: uppercase;
    letter-spacing: .5px;
  }
  .info-panel pre {
    margin: 0;
    font-family: "Consolas", "Monaco", "Courier New", monospace;
    font-size: 12px;
    color: #2A3344;
    white-space: pre-wrap;
    word-wrap: break-word;
    max-height: 300px;
    overflow-y: auto;
  }
  .info-panel pre .key { color: var(--label); }
  .info-panel pre .num { color: var(--navy); font-weight: 600; }

  /* ---- Distribution preview chart ---- */
  .chart {
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 12px 14px;
    background: #FFFFFF;
  }
  .chart h4 {
    margin: 0 0 10px;
    font-size: 12px;
    font-weight: 700;
    color: var(--label);
    text-transform: uppercase;
    letter-spacing: .5px;
  }
  .bars {
    display: grid;
    grid-template-columns: 165px 1fr 50px;
    gap: 4px 8px;
    align-items: center;
  }
  .bars .name {
    font-size: 12px;
    text-align: right;
    color: #2A3344;
    line-height: 1.6;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }
  .bars .bar-track {
    background: #EEF1F7;
    height: 18px;
    border-radius: 3px;
    overflow: hidden;
    position: relative;
  }
  .bars .bar-fill {
    height: 100%;
    border-radius: 3px;
    transition: width .25s ease;
  }
  .bars .count {
    font-size: 13px;
    font-weight: 700;
    color: var(--navy);
    text-align: left;
    padding-left: 2px;
  }
  .bar-bb     { background: var(--bb-color); }
  .bar-ran    { background: var(--ran-color); }
  .bar-ring   { background: var(--ring-color); }
  .bar-power  { background: var(--power-color); }
  .bar-equip  { background: var(--equip-color); }
  .bar-cesr   { background: var(--cesr-color); }
  .bar-status { background: var(--gold); }
  .bar-other  { background: #1F3864; }
  .empty-state {
    grid-column: 1/-1;
    text-align: center;
    color: var(--muted);
    font-style: italic;
    padding: 30px 10px;
    font-size: 13px;
  }

  /* ---- Action bar ---- */
  .actions {
    margin-top: 12px;
    display: flex;
    gap: 10px;
    align-items: center;
    padding-top: 12px;
    border-top: 1px solid var(--border);
  }
  button.primary {
    background: var(--navy);
    color: white;
    border: none;
    padding: 11px 22px;
    font-size: 14px;
    font-weight: 600;
    border-radius: 6px;
    cursor: pointer;
    transition: background .12s;
  }
  button.primary:hover:not(:disabled) { background: var(--navy-2); }
  button.primary:active:not(:disabled) { background: var(--navy-3); }
  button.primary:disabled {
    background: #B6C2D5;
    cursor: not-allowed;
  }
  .progress-wrap {
    flex: 1;
    height: 8px;
    background: #EEF1F7;
    border-radius: 4px;
    overflow: hidden;
    display: none;
  }
  .progress-wrap.show { display: block; }
  .progress-bar {
    height: 100%;
    background: var(--navy);
    width: 0;
    transition: width .35s ease;
    border-radius: 4px;
  }
  .progress-bar.indeterminate {
    width: 30%;
    animation: indet 1.4s ease-in-out infinite;
  }
  @keyframes indet {
    0%   { margin-left: -30%; }
    100% { margin-left: 100%; }
  }
  .status {
    font-size: 13px;
    color: var(--muted);
    text-align: right;
    min-width: 200px;
  }
  .status.success { color: var(--green); font-weight: 600; }
  .status.error   { color: var(--red);   font-weight: 600; }
  .spinner {
    display: inline-block;
    width: 14px;
    height: 14px;
    border: 2px solid #B6C2D5;
    border-top-color: var(--navy);
    border-radius: 50%;
    animation: spin 1s linear infinite;
    vertical-align: -2px;
    margin-right: 6px;
  }
  @keyframes spin { to { transform: rotate(360deg); } }

  /* ---- Error box ---- */
  .err {
    background: #FDECEA;
    border-left: 4px solid var(--red);
    padding: 10px 14px;
    border-radius: 4px;
    color: #6A1B1B;
    margin-top: 12px;
    font-size: 13px;
    white-space: pre-wrap;
    font-family: "Consolas", "Monaco", monospace;
  }

  /* ---- Footer ---- */
  footer {
    text-align: center;
    color: var(--muted);
    font-size: 12px;
    padding: 22px 20px;
  }
  footer .pill {
    display: inline-block;
    padding: 3px 10px;
    background: #E5EAF3;
    border-radius: 12px;
    font-family: "Consolas", "Monaco", monospace;
    color: var(--navy);
    font-weight: 600;
    margin: 0 4px;
  }

  /* ---- Responsive ---- */
  @media (max-width: 900px) {
    .preview-grid { grid-template-columns: 1fr; }
    .kpis { grid-template-columns: repeat(5, 1fr); gap: 4px; }
    .kpi .value { font-size: 18px; }
    .bars { grid-template-columns: 130px 1fr 40px; }
  }
</style>
</head>
<body>
<header>
  <div class="header-inner">
    <h1>Report Gen <span class="accent">— JTL NOC Monitoring</span></h1>
    <div class="meta">v{{ version }} &nbsp;•&nbsp; © Emmanuel Mutua &amp; Alex Wachira</div>
  </div>
</header>

<main>
  <div class="tabs">
    <div class="tab gpon active" data-kind="gpon">GPON</div>
    <div class="tab ent" data-kind="enterprise">Enterprise</div>
  </div>

  <div class="panel">
    <!-- Input section -->
    <div class="section">
      <div class="section-title" id="input-title">GPON input file</div>
      <div class="drop" id="drop">
        <strong>Drop the Issue.xlsx file here, or click to browse</strong>
        <small>Excel files only (.xlsx) &nbsp;•&nbsp; up to {{ max_mb }} MB</small>
        <input type="file" id="file" accept=".xlsx,.xlsm">
      </div>
      <div class="file-row" id="fileinfo">
        <div>
          <div class="name" id="filename">—</div>
          <div class="size" id="filesize"></div>
        </div>
        <button class="clear" id="clear" title="Remove this file">✕ Clear</button>
      </div>
    </div>

    <!-- Preview grid: source-info panel + distribution chart -->
    <div class="kpis" id="kpis"></div>
    <div class="preview-grid" id="preview-grid">
      <div class="info-panel">
        <h4>Source sheet info</h4>
        <pre id="info-text"></pre>
      </div>
      <div class="chart">
        <h4>Distribution preview</h4>
        <div class="bars" id="bars"></div>
      </div>
    </div>

    <!-- Action bar -->
    <div class="actions">
      <button class="primary" id="generate" disabled>Generate <span id="kindlabel">GPON</span> report</button>
      <div class="progress-wrap" id="prog"><div class="progress-bar" id="progbar"></div></div>
      <div class="status" id="status">Choose a file to begin.</div>
    </div>
    <div id="errbox"></div>
  </div>
</main>

<footer>
  Running on <span class="pill">{{ server_addr }}</span>
  &nbsp;•&nbsp; Files are processed in memory and not stored on the server.
</footer>

<script>
  const drop      = document.getElementById('drop');
  const fileInput = document.getElementById('file');
  const fileInfo  = document.getElementById('fileinfo');
  const filename  = document.getElementById('filename');
  const filesize  = document.getElementById('filesize');
  const clearBtn  = document.getElementById('clear');
  const status    = document.getElementById('status');
  const generate  = document.getElementById('generate');
  const kindlabel = document.getElementById('kindlabel');
  const inputTitle = document.getElementById('input-title');
  const previewGrid = document.getElementById('preview-grid');
  const kpis      = document.getElementById('kpis');
  const infoText  = document.getElementById('info-text');
  const bars      = document.getElementById('bars');
  const errbox    = document.getElementById('errbox');
  const prog      = document.getElementById('prog');
  const progbar   = document.getElementById('progbar');

  let currentKind = 'gpon';
  let chosenFile  = null;

  // ---- KPI specs per kind ----
  const KPI_SPECS = {
    gpon: [
      { label: 'Total Rows',       key: 'total_rows' },
      { label: 'Dropped (Closed)', key: 'dropped_closed' },
      { label: 'GPON Rows',        key: 'gpon_rows' },
      { label: 'Doubles',          key: 'double_tickets' },
      { label: 'Splitters',        key: 'splitter_rows' },
    ],
    enterprise: [
      { label: 'Total Rows',       key: 'total_rows' },
      { label: 'Dropped (Closed)', key: 'dropped_closed' },
      { label: 'Dropped (GPON)',   key: 'dropped_gpon' },
      { label: 'Will Process',     key: 'kept_rows' },
      { label: 'CESR Matches',     key: 'cesr_count' },
    ],
  };

  // ---- Color rule for distribution bars ----
  function barClass(label) {
    const s = label.toLowerCase();
    if (s.startsWith('bb '))           return 'bar-bb';
    if (s.includes('6150') || s.includes('6120') || s.includes('1050')) return 'bar-ran';
    if (s.startsWith('access ring'))   return 'bar-ring';
    if (s.startsWith('power'))         return 'bar-power';
    if (s.startsWith('equipment'))     return 'bar-equip';
    if (s.startsWith('cesr'))          return 'bar-cesr';
    // GPON status sheets / specials
    if (['customer action','under monitoring','temporary restoration'].includes(s)) return 'bar-status';
    if (s.includes('double'))          return 'bar-power';
    if (s.includes('splitter'))        return 'bar-ring';
    return 'bar-other';
  }

  // ---- Initial render: empty KPI cards ----
  function renderEmptyKpis() {
    const specs = KPI_SPECS[currentKind];
    kpis.innerHTML = specs.map(s =>
      '<div class="kpi empty"><div class="label">' + s.label + '</div><div class="value">—</div></div>'
    ).join('');
  }
  renderEmptyKpis();

  // ---- Tab switching ----
  document.querySelectorAll('.tab').forEach(tab => {
    tab.addEventListener('click', () => {
      if (tab.classList.contains('active')) return;
      document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      currentKind = tab.dataset.kind;
      kindlabel.textContent = currentKind === 'gpon' ? 'GPON' : 'Enterprise';
      inputTitle.textContent = (currentKind === 'gpon' ? 'GPON' : 'Enterprise') + ' input file';
      renderEmptyKpis();
      previewGrid.classList.remove('show');
      if (chosenFile) startPreview();
    });
  });

  // ---- Drag & drop ----
  drop.addEventListener('click', () => fileInput.click());
  ['dragover','dragenter'].forEach(ev =>
    drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('hover'); })
  );
  ['dragleave','drop'].forEach(ev =>
    drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('hover'); })
  );
  drop.addEventListener('drop', e => {
    if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
  });
  fileInput.addEventListener('change', e => {
    if (e.target.files.length) handleFile(e.target.files[0]);
  });
  clearBtn.addEventListener('click', () => {
    chosenFile = null;
    fileInput.value = '';
    fileInfo.classList.remove('show');
    previewGrid.classList.remove('show');
    renderEmptyKpis();
    generate.disabled = true;
    status.textContent = 'Choose a file to begin.';
    status.className = 'status';
    clearError();
  });

  function handleFile(f) {
    const lower = f.name.toLowerCase();
    if (!(lower.endsWith('.xlsx') || lower.endsWith('.xlsm'))) {
      showError('Only .xlsx files are supported.');
      return;
    }
    chosenFile = f;
    filename.textContent = f.name;
    filesize.textContent = (f.size / 1024 / 1024).toFixed(2) + ' MB';
    fileInfo.classList.add('show');
    generate.disabled = false;
    clearError();
    startPreview();
  }

  function clearError() { errbox.innerHTML = ''; }
  function showError(msg) { errbox.innerHTML = '<div class="err">' + escapeHtml(msg) + '</div>'; }

  function setStatus(msg, kind) {
    status.textContent = msg;
    status.className = 'status' + (kind ? ' ' + kind : '');
  }
  function setStatusBusy(msg) {
    status.innerHTML = '<span class="spinner"></span>' + escapeHtml(msg);
    status.className = 'status';
  }

  // ---- Preview ----
  async function startPreview() {
    if (!chosenFile) return;
    setStatusBusy('Reading preview…');
    previewGrid.classList.remove('show');
    clearError();

    const fd = new FormData();
    fd.append('file', chosenFile);
    fd.append('kind', currentKind);
    try {
      const res = await fetch('/preview', { method: 'POST', body: fd });
      if (!res.ok) {
        const txt = await res.text();
        showError('Preview failed: ' + txt);
        setStatus('Preview failed', 'error');
        return;
      }
      const data = await res.json();
      renderPreview(data);
      setStatus('Preview ready — click Generate when you\'re set.');
      previewGrid.classList.add('show');
    } catch (e) {
      showError('Preview failed: ' + e.message);
      setStatus('Preview failed', 'error');
    }
  }

  function renderPreview(d) {
    const specs = KPI_SPECS[currentKind];
    kpis.innerHTML = specs.map(s => {
      const v = d[s.key];
      const display = (v === undefined || v === null) ? '—' : v;
      return '<div class="kpi"><div class="label">' + s.label + '</div><div class="value">' + display + '</div></div>';
    }).join('');

    // Distribution counts
    let counts;
    if (currentKind === 'gpon') {
      counts = Object.assign({}, d.status_counts || {});
      if (d.double_tickets) counts['Double Tickets'] = d.double_tickets;
      if (d.splitter_rows)  counts['Splitter rows']  = d.splitter_rows;
    } else {
      counts = Object.assign({}, d.bucket_counts || {});
      if (d.cesr_count) counts['CESR'] = d.cesr_count;
    }
    const items = Object.entries(counts).filter(([_, v]) => v > 0);

    if (!items.length) {
      bars.innerHTML = '<div class="empty-state">No buckets matched.</div>';
    } else {
      items.sort((a, b) => b[1] - a[1]);
      const max = Math.max(...items.map(kv => kv[1]));
      bars.innerHTML = items.map(([label, val]) => {
        const pct = max ? (val / max * 100).toFixed(1) : 0;
        const cls = barClass(label);
        return (
          '<div class="name" title="' + escapeHtml(label) + '">' + escapeHtml(label) + '</div>' +
          '<div class="bar-track"><div class="bar-fill ' + cls + '" style="width:' + pct + '%"></div></div>' +
          '<div class="count">' + val + '</div>'
        );
      }).join('');
    }

    // Source-info text panel
    const lines = [];
    lines.push('File:           ' + chosenFile.name);
    lines.push('Sheets:         ' + (d.sheet_names || []).join(', '));
    lines.push('Total rows:     ' + d.total_rows);
    lines.push('Cleaned:        ' + (d.cleaned_total || 0) + '  (dropped closed: ' + (d.dropped_closed || 0) + ')');

    if (currentKind === 'gpon') {
      lines.push('GPON rows:      ' + d.gpon_rows + '   (non-GPON: ' + d.non_gpon_rows + ')');
      lines.push('Doubles:        ' + d.double_tickets);
      lines.push('Splitters:      ' + d.splitter_rows);
      lines.push('');
      lines.push('Status breakdown (GPON only):');
      Object.entries(d.status_counts || {}).forEach(([k, v]) => {
        lines.push('  ' + pad(k, 24) + ' ' + v);
      });
      const cats = d.category_counts || {};
      const sortedCats = Object.entries(cats).sort((a, b) => b[1] - a[1]).slice(0, 8);
      if (sortedCats.length) {
        lines.push('');
        lines.push('Top Category 1 buckets:');
        sortedCats.forEach(([k, v]) => {
          const truncated = k.length > 55 ? k.substring(0, 55) : k;
          lines.push('  ' + pad(truncated, 56) + ' ' + v);
        });
      }
    } else {
      lines.push('After GPON exclusion: ' + d.kept_rows);
      lines.push('');
      lines.push('Bucket distribution:');
      const buckets = d.bucket_counts || {};
      const nonZero = Object.entries(buckets).filter(([_, v]) => v > 0)
                                              .sort((a, b) => b[1] - a[1]);
      if (!nonZero.length) {
        lines.push('  (no rows matched any bucket)');
      } else {
        nonZero.forEach(([k, v]) => {
          lines.push('  ' + pad(k, 32) + ' ' + v);
        });
      }
      if (d.cesr_count) {
        lines.push('  ' + pad('CESR (extra sheet)', 32) + ' ' + d.cesr_count);
      }
    }
    infoText.textContent = lines.join('\n');
  }

  function pad(s, n) {
    s = String(s);
    if (s.length >= n) return s;
    return s + ' '.repeat(n - s.length);
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  }

  // ---- Generate ----
  generate.addEventListener('click', async () => {
    if (!chosenFile) return;
    generate.disabled = true;
    prog.classList.add('show');
    progbar.classList.add('indeterminate');
    progbar.style.width = '';
    setStatusBusy('Generating report — this may take a moment…');
    clearError();

    const fd = new FormData();
    fd.append('file', chosenFile);
    fd.append('kind', currentKind);
    try {
      const res = await fetch('/generate', { method: 'POST', body: fd });
      if (!res.ok) {
        const txt = await res.text();
        showError('Generation failed: ' + txt);
        setStatus('Failed', 'error');
        prog.classList.remove('show');
        progbar.classList.remove('indeterminate');
        generate.disabled = false;
        return;
      }
      const blob = await res.blob();
      const dispo = res.headers.get('Content-Disposition') || '';
      let outName = currentKind === 'gpon' ? 'GPON Report.xlsx' : 'Enterprise Report.xlsx';
      const m = /filename="?([^"]+)"?/.exec(dispo);
      if (m) outName = m[1];

      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = outName;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(a.href);

      progbar.classList.remove('indeterminate');
      progbar.style.width = '100%';
      setStatus('✓ Done — your download has started.', 'success');
    } catch (e) {
      showError('Generation failed: ' + e.message);
      setStatus('Failed', 'error');
      progbar.classList.remove('indeterminate');
    } finally {
      generate.disabled = false;
      setTimeout(() => prog.classList.remove('show'), 800);
    }
  });
</script>
</body>
</html>
"""


# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------
def _is_allowed(filename: str) -> bool:
    lower = filename.lower()
    return any(lower.endswith(ext) for ext in ALLOWED_EXTENSIONS)


def _save_upload_to_temp(file_storage) -> Path:
    """Save the uploaded file to a unique temp path and return that path."""
    safe_name = secure_filename(file_storage.filename) or "upload.xlsx"
    tmp_dir = Path(tempfile.mkdtemp(prefix="reportgen_"))
    out_path = tmp_dir / safe_name
    file_storage.save(out_path)
    return out_path


def _output_filename(kind: str) -> str:
    """'<Date> GPON Report.xlsx' / '<Date> Enterprise Report.xlsx'."""
    label = "GPON" if kind == "gpon" else "Enterprise"
    return f"{report_lib._date_title()} {label} Report.xlsx"


def _server_ip() -> str:
    """Best-effort LAN IP discovery for the footer."""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


# ----------------------------------------------------------------------------
# Routes
# ----------------------------------------------------------------------------
@app_flask.route("/", methods=["GET"])
def index():
    server_addr = f"http://{_server_ip()}:{PORT}"
    return render_template_string(
        PAGE_HTML,
        version=report_lib.APP_VERSION,
        max_mb=MAX_UPLOAD_MB,
        server_addr=server_addr,
    )


@app_flask.route("/preview", methods=["POST"])
def preview():
    f = request.files.get("file")
    kind = request.form.get("kind", "gpon")
    if not f or not _is_allowed(f.filename):
        return ("Invalid file (must be .xlsx)", 400)

    in_path = None
    try:
        in_path = _save_upload_to_temp(f)
        if kind == "gpon":
            info = report_lib.gpon_preview(in_path)
        else:
            info = report_lib.enterprise_preview(in_path)
        return jsonify(info)
    except Exception as exc:
        log.error("Preview failed: %s", exc)
        log.error(traceback.format_exc())
        return (f"{exc}", 500)
    finally:
        # Clean up the temp file and parent dir
        try:
            if in_path and in_path.exists():
                in_path.unlink()
                in_path.parent.rmdir()
        except Exception:
            pass


@app_flask.route("/generate", methods=["POST"])
def generate():
    f = request.files.get("file")
    kind = request.form.get("kind", "gpon")
    if not f or not _is_allowed(f.filename):
        return ("Invalid file (must be .xlsx)", 400)

    in_path = None
    out_path = None
    try:
        in_path = _save_upload_to_temp(f)
        out_name = _output_filename(kind)
        out_path = in_path.parent / out_name

        if kind == "gpon":
            report_lib.gpon_process(in_path, out_path)
        else:
            report_lib.enterprise_process(in_path, out_path)

        # Stream the file back into memory so we can delete the temp dir
        buf = io.BytesIO(out_path.read_bytes())
        buf.seek(0)
        log.info("Served %s: %s (%d bytes)", kind, out_name, buf.getbuffer().nbytes)
        return send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=out_name,
        )
    except Exception as exc:
        log.error("Generate failed: %s", exc)
        log.error(traceback.format_exc())
        return (f"{exc}", 500)
    finally:
        # Always clean up
        try:
            if out_path and out_path.exists():
                out_path.unlink()
        except Exception:
            pass
        try:
            if in_path and in_path.exists():
                in_path.unlink()
                in_path.parent.rmdir()
        except Exception:
            pass


@app_flask.errorhandler(413)
def too_large(_e):
    return (f"File too large. Maximum size is {MAX_UPLOAD_MB} MB.", 413)


# ----------------------------------------------------------------------------
# Entry point
# ----------------------------------------------------------------------------
def main():
    ip = _server_ip()
    print()
    print("=" * 70)
    print(f"  Report Gen \u2014 JTL NOC Monitoring  (Web)  v{report_lib.APP_VERSION}")
    print(f"  Authors: Emmanuel Mutua & Alex Wachira")
    print("=" * 70)
    print()
    print(f"  Local URL:    http://localhost:{PORT}")
    print(f"  Network URL:  http://{ip}:{PORT}")
    print()
    print(f"  Share the Network URL with anyone on the same network.")
    print(f"  Press Ctrl+C to stop the server.")
    print()
    print("=" * 70)
    print()
    app_flask.run(host=HOST, port=PORT, debug=False, use_reloader=False)


if __name__ == "__main__":
    main()
