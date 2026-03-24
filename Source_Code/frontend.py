"""
Web-based Frontend for the Automation Prototype.
Launches a local Flask server and opens the UI in the system's default browser.
"""
import sys
import io
import logging
import threading
import webbrowser
from tkinter import filedialog, Tk
from flask import Flask, request, jsonify, render_template_string

# Suppress Werkzeug's development-server warning so it doesn't appear as a
# PowerShell NativeCommandError (the message is written to stderr at WARNING level).
_wz_log = logging.getLogger("werkzeug")
_wz_log.setLevel(logging.ERROR)
_wz_log.propagate = False

# ── Import backend functions ──────────────────────────────────────────────────
from AutomationPrototype import (
    automated_testing_activity,
    severity_counter,
    assessment_results,
    recommendations,
    appendix,
    finding_details,
    DEFAULT_RECOMMENDATIONS_PATH,
    DEFAULT_EXECUTIVE_REPORT_PATH,
    DEFAULT_ACTIVITY_REPORT_PATH,
    DEFAULT_FINDINGS_REPORT_PATH,
    DEFAULT_TECHNICAL_REPORT_PATH,
)

app = Flask(__name__)

# ── HTML page ─────────────────────────────────────────────────────────────────
HTML = r"""
<!DOCTYPE html>
<html lang="en" data-theme="dark">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>RAT — Report Automation Tool</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@300;400;500;600;700&display=swap" rel="stylesheet">
  <style>
    :root[data-theme="dark"] {
      --bg:           #0b0f16;
      --surface:      #111722;
      --surface2:     #161d2a;
      --border:       #1e2a3a;
      --border-med:   #283649;
      --text:         #dce6f5;
      --text-muted:   #7e94b4;
      --text-dim:     #38506a;
      --accent:       #2563eb;
      --accent-hover: #1d4ed8;
      --accent-dim:   rgba(37,99,235,0.08);
      --accent-ring:  rgba(37,99,235,0.22);
      --green:        #16a34a;
      --green-hover:  #15803d;
      --log-bg:       #060a10;
      --warn:         #f87171;
      --ok:           #4ade80;
      --info:         #60a5fa;
      --purple:       #c084fc;
    }
    :root[data-theme="light"] {
      --bg:           #eef2f9;
      --surface:      #ffffff;
      --surface2:     #e5eaf4;
      --border:       #c8d4e6;
      --border-med:   #b0bfd4;
      --text:         #0c1525;
      --text-muted:   #4e6080;
      --text-dim:     #9aadc4;
      --accent:       #1d4ed8;
      --accent-hover: #1e40af;
      --accent-dim:   rgba(29,78,216,0.07);
      --accent-ring:  rgba(29,78,216,0.18);
      --green:        #15803d;
      --green-hover:  #166534;
      --log-bg:       #0c1220;
      --warn:         #dc2626;
      --ok:           #16a34a;
      --info:         #1d4ed8;
      --purple:       #7c3aed;
    }

    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    html, body { height: 100%; }
    body {
      font-family: "Space Grotesk", -apple-system, sans-serif;
      background: var(--bg); color: var(--text);
      display: flex; flex-direction: column;
      font-size: 13px; line-height: 1.5;
      -webkit-font-smoothing: antialiased;
    }

    .header {
      border-bottom: 1px solid var(--border);
      padding: 0 32px; height: 58px;
      display: flex; align-items: center; gap: 20px;
      flex-shrink: 0;
    }
    .wordmark { display: flex; align-items: center; gap: 14px; flex: 1; }
    .wordmark-name { font-size: 15px; font-weight: 700; letter-spacing: -0.02em; }
    .wordmark-rule { width: 1px; height: 16px; background: var(--border-med); flex-shrink: 0; }
    .wordmark-desc { font-size: 11.5px; color: var(--text-muted); }
    .header-badge {
      font-size: 10px; font-weight: 700; letter-spacing: 0.06em;
      background: var(--accent-dim); color: var(--accent);
      border: 1px solid rgba(37,99,235,0.18);
      padding: 3px 8px; border-radius: 3px;
    }

    .btn {
      display: inline-flex; align-items: center; justify-content: center; gap: 7px;
      padding: 0 18px; height: 34px; border-radius: 4px; border: 1px solid transparent;
      font-family: "Space Grotesk", inherit; font-size: 13px; font-weight: 600;
      cursor: pointer; transition: background 0.1s, border-color 0.1s;
      white-space: nowrap; line-height: 1;
    }
    .btn:focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; }
    .btn:disabled { opacity: 0.38; cursor: not-allowed; pointer-events: none; }
    .btn-theme {
      background: transparent; color: var(--text-muted);
      border-color: var(--border-med); height: 30px; font-size: 12px; padding: 0 12px;
    }
    .btn-theme:hover { background: var(--surface2); color: var(--text); }
    .btn-primary { background: var(--accent); color: #fff; border-color: var(--accent); }
    .btn-primary:hover { background: var(--accent-hover); border-color: var(--accent-hover); }
    .btn-green { background: var(--green); color: #fff; border-color: var(--green); }
    .btn-green:hover { background: var(--green-hover); border-color: var(--green-hover); }
    .btn-ghost {
      background: transparent; color: var(--text-muted);
      border-color: var(--border); height: 30px; font-size: 11.5px; padding: 0 11px;
    }
    .btn-ghost:hover { background: var(--surface2); color: var(--text); }
    .btn-action { height: 46px; font-size: 14px; font-weight: 700; letter-spacing: -0.01em; }

    .spinner {
      display: none; width: 13px; height: 13px; flex-shrink: 0;
      border: 2px solid rgba(255,255,255,0.2); border-top-color: #fff;
      border-radius: 50%; animation: spin 0.5s linear infinite;
    }
    .btn.loading .spinner { display: block; }
    .btn.loading .btn-label { opacity: 0.6; }
    @keyframes spin { to { transform: rotate(360deg); } }

    .main {
      flex: 1; min-height: 0;
      max-width: 1120px; width: 100%; margin: 0 auto;
      padding: 28px 32px 24px;
      display: flex; flex-direction: column;
    }
    .section-label {
      font-size: 9px; font-weight: 700;
      text-transform: uppercase; letter-spacing: 0.12em;
      color: var(--text-dim); margin-bottom: 8px;
    }

    .file-table { width: 100%; border-collapse: collapse; }
    .file-table tr { border-top: 1px solid var(--border); transition: background 0.1s; }
    .file-table tr:last-child { border-bottom: 1px solid var(--border); }
    .file-table tr:hover { background: var(--surface2); }
    .file-table tr:has(input:focus) { background: var(--surface2); box-shadow: inset 2px 0 0 var(--accent); }
    .file-table td { padding: 9px 0; vertical-align: middle; }
    .td-num {
      width: 44px; padding-right: 14px;
      font-size: 11px; font-weight: 700; color: var(--text-dim);
      font-variant-numeric: tabular-nums; user-select: none;
    }
    .td-label { width: 210px; padding-right: 18px; }
    .td-label label {
      font-size: 12.5px; font-weight: 500; color: var(--text-muted);
      cursor: pointer; display: flex; flex-direction: column; gap: 1px;
    }
    .label-type { font-size: 9.5px; font-weight: 700; letter-spacing: 0.06em; color: var(--text-dim); }
    .td-input { padding-right: 12px; }
    .td-input input {
      width: 100%; padding: 7px 11px; border-radius: 4px;
      border: 1px solid var(--border); background: var(--surface);
      color: var(--text); font-size: 11.5px;
      font-family: "Cascadia Code", "Consolas", monospace;
      outline: none; transition: border-color 0.12s, box-shadow 0.12s;
    }
    .td-input input:focus { border-color: var(--accent); box-shadow: 0 0 0 3px var(--accent-ring); }
    .td-input input::placeholder { color: var(--text-dim); }
    .td-btn { width: 80px; }

    .sp-20 { height: 20px; flex-shrink: 0; }
    .sp-28 { height: 28px; flex-shrink: 0; }
    .sp-32 { height: 32px; flex-shrink: 0; }

    .actions {
      display: grid;
      grid-template-columns: 1fr 1fr 1fr auto;
      gap: 10px; align-items: center;
    }
    .actions-row2 {
      display: grid;
      grid-template-columns: 1fr 1fr 1fr auto;
      gap: 10px; align-items: center;
      margin-top: 10px;
    }

    .log-wrap { flex: 1; min-height: 0; display: flex; flex-direction: column; }
    .log-titlebar {
      background: var(--surface); border: 1px solid var(--border);
      border-bottom: none; border-radius: 5px 5px 0 0;
      display: flex; align-items: center; gap: 10px;
      padding: 0 14px; height: 32px; flex-shrink: 0;
    }
    .log-title {
      font-size: 9px; font-weight: 700;
      text-transform: uppercase; letter-spacing: 0.12em;
      color: var(--text-muted); flex: 1;
    }
    .log-statustext {
      font-family: "Cascadia Code", "Consolas", monospace;
      font-size: 11px; color: var(--text-dim); transition: color 0.2s;
    }
    .log-statustext.busy { color: var(--info); }
    .status-dot {
      width: 6px; height: 6px; border-radius: 50%;
      background: var(--ok); flex-shrink: 0; transition: background 0.2s;
    }
    .status-dot.busy { background: var(--info); animation: pulse 1.1s ease-in-out infinite; }
    @keyframes pulse { 0%,100% { opacity:1; } 50% { opacity:0.28; } }

    #log {
      flex: 1; min-height: 220px;
      background: var(--log-bg); border: 1px solid var(--border);
      border-radius: 0 0 5px 5px; padding: 12px 16px;
      font-family: "Cascadia Code", "Consolas", monospace;
      font-size: 12px; line-height: 1.72;
      overflow-y: auto; white-space: pre-wrap; word-break: break-all;
    }
    #log::-webkit-scrollbar { width: 5px; }
    #log::-webkit-scrollbar-track { background: transparent; }
    #log::-webkit-scrollbar-thumb { background: var(--border-med); border-radius: 3px; }
    #log::-webkit-scrollbar-thumb:hover { background: var(--text-dim); }

    .log-TS     { color: #29394f; }
    .log-INFO   { color: var(--info); }
    .log-OK     { color: var(--ok); }
    .log-WARN   { color: var(--warn); }
    .log-PLAIN  { color: #8da4c4; }
    .log-HEADER { color: var(--purple); font-weight: 700; }
    .log-INIT   { color: #29394f; font-style: italic; }

    .fade-up {
      opacity: 0; transform: translateY(10px);
      animation: fadeUp 0.38s cubic-bezier(0.16,1,0.3,1) forwards;
    }
    @keyframes fadeUp { to { opacity:1; transform:translateY(0); } }
    .d1 { animation-delay: 0.04s; } .d2 { animation-delay: 0.10s; }
    .d3 { animation-delay: 0.17s; } .d4 { animation-delay: 0.24s; }
    .d5 { animation-delay: 0.30s; } .d6 { animation-delay: 0.36s; }
    .d7 { animation-delay: 0.42s; }
    @media (prefers-reduced-motion: reduce) {
      .fade-up { animation: none; opacity:1; transform:none; }
    }
  </style>
</head>
<body>
  <header class="header fade-up">
    <div class="wordmark">
      <span class="wordmark-name">Report Automation Tool</span>
      <span class="wordmark-rule"></span>
      <span class="wordmark-desc">pen test reporting pipeline</span>
    </div>
    <span class="header-badge">ICSI&thinsp;499</span>
    <button class="btn btn-theme" id="themeBtn" onclick="toggleTheme()">&#9728;&ensp;Light</button>
  </header>

  <main class="main">
    <div class="sp-28 fade-up d1"></div>

    <!-- File inputs -->
    <div class="section-label fade-up d1">Input Files</div>
    <table class="file-table fade-up d2" aria-label="Report file paths">
      <tbody>
        <tr>
          <td class="td-num">01</td>
          <td class="td-label">
            <label for="activityPath">Activity Report<span class="label-type">PDF</span></label>
          </td>
          <td class="td-input">
            <input id="activityPath" type="text" value="{{ activity_default }}" placeholder="./Reports/activity-report.pdf" />
          </td>
          <td class="td-btn"><button class="btn btn-ghost" onclick="browseFile('activityPath','pdf')">Browse&hellip;</button></td>
        </tr>
        <tr>
          <td class="td-num">02</td>
          <td class="td-label">
            <label for="findingsPath">Findings Report<span class="label-type">DOCX &mdash; output</span></label>
          </td>
          <td class="td-input">
            <input id="findingsPath" type="text" value="{{ findings_default }}" placeholder="./Reports/findings-report.docx" />
          </td>
          <td class="td-btn"><button class="btn btn-ghost" onclick="browseFile('findingsPath','docx')">Browse&hellip;</button></td>
        </tr>
        <tr>
          <td class="td-num">03</td>
          <td class="td-label">
            <label for="technicalPath">Technical Report<span class="label-type">PDF</span></label>
          </td>
          <td class="td-input">
            <input id="technicalPath" type="text" value="{{ technical_default }}" placeholder="./Reports/technical-report.pdf" />
          </td>
          <td class="td-btn"><button class="btn btn-ghost" onclick="browseFile('technicalPath','pdf')">Browse&hellip;</button></td>
        </tr>
        <tr>
          <td class="td-num">04</td>
          <td class="td-label">
            <label for="executivePath">Executive Report<span class="label-type">PDF</span></label>
          </td>
          <td class="td-input">
            <input id="executivePath" type="text" value="{{ executive_default }}" placeholder="./Reports/executive-report.pdf" />
          </td>
          <td class="td-btn"><button class="btn btn-ghost" onclick="browseFile('executivePath','pdf')">Browse&hellip;</button></td>
        </tr>
        <tr>
          <td class="td-num">05</td>
          <td class="td-label">
            <label for="recommendationPath">Findings &amp; Recommendations<span class="label-type">XLSX</span></label>
          </td>
          <td class="td-input">
            <input id="recommendationPath" type="text" value="{{ recommendation_default }}" placeholder="./Reports/FindingsDetailsAndRecommendations.xlsx" />
          </td>
          <td class="td-btn"><button class="btn btn-ghost" onclick="browseFile('recommendationPath','xlsx')">Browse&hellip;</button></td>
        </tr>
      </tbody>
    </table>

    <div class="sp-28 fade-up d3"></div>

    <!-- Parameters -->
    <div class="section-label fade-up d3">Parameters</div>
    <table class="file-table fade-up d4" aria-label="Run parameters">
      <tbody>
        <tr>
          <td class="td-num">P1</td>
          <td class="td-label">
            <label for="pageRange">Activity Log Pages<span class="label-type">e.g. 3-5 &mdash; Run Testing</span></label>
          </td>
          <td class="td-input"><input id="pageRange" type="text" placeholder="e.g. 3-5" /></td>
          <td class="td-btn"></td>
        </tr>
        <tr>
          <td class="td-num">P2</td>
          <td class="td-label">
            <label for="pageNumber">Assessment Results Page #<span class="label-type">e.g. 7 &mdash; Run Assessment</span></label>
          </td>
          <td class="td-input"><input id="pageNumber" type="text" placeholder="e.g. 7" /></td>
          <td class="td-btn"></td>
        </tr>
        <tr>
          <td class="td-num">P3</td>
          <td class="td-label">
            <label for="environment">Environment<span class="label-type">Internal / External &mdash; Recommendations</span></label>
          </td>
          <td class="td-input"><input id="environment" type="text" placeholder="Internal or External" /></td>
          <td class="td-btn"></td>
        </tr>
      </tbody>
    </table>

    <div class="sp-32 fade-up d5"></div>

    <!-- Actions row 1 -->
    <div class="actions fade-up d5">
      <button class="btn btn-primary btn-action" id="runTestingBtn" onclick="runTesting()">
        <span class="spinner"></span><span class="btn-label">&#9654;&ensp;Run Automated Testing</span>
      </button>
      <button class="btn btn-primary btn-action" id="runAssessmentsBtn" onclick="runAssessments()">
        <span class="spinner"></span><span class="btn-label">&#9654;&ensp;Run Assessment Results</span>
      </button>
      <button class="btn btn-primary btn-action" id="runRecommendationsBtn" onclick="runRecommendations()">
        <span class="spinner"></span><span class="btn-label">&#9654;&ensp;Run Recommendations</span>
      </button>
      <span></span>
    </div>

    <!-- Actions row 2 -->
    <div class="actions-row2 fade-up d6">
      <button class="btn btn-primary btn-action" id="runAppendixBtn" onclick="runAppendix()">
        <span class="spinner"></span><span class="btn-label">&#9654;&ensp;Run Appendix</span>
      </button>
      <button class="btn btn-primary btn-action" id="runFindingsBtn" onclick="runFindings()">
        <span class="spinner"></span><span class="btn-label">&#9654;&ensp;Run Findings Details</span>
      </button>
      <button class="btn btn-green btn-action" id="runSeverityBtn" onclick="runSeverity()">
        <span class="spinner"></span><span class="btn-label">&#9650;&ensp;Count Vulnerabilities</span>
      </button>
      <button class="btn btn-ghost" onclick="clearLog()">&#10005;&ensp;Clear Log</button>
    </div>

    <div class="sp-28 fade-up d7"></div>

    <!-- Log -->
    <div class="log-wrap fade-up d7">
      <div class="log-titlebar">
        <span class="log-title">Output</span>
        <span class="log-statustext" id="logStatus">ready</span>
        <span class="status-dot" id="statusDot"></span>
      </div>
      <div id="log"></div>
    </div>
  </main>

  <script>
    function toggleTheme() {
      const html = document.documentElement;
      const isDark = html.getAttribute('data-theme') === 'dark';
      html.setAttribute('data-theme', isDark ? 'light' : 'dark');
      document.getElementById('themeBtn').innerHTML =
        isDark ? '\uD83C\uDF19&ensp;Dark' : '\u2600&ensp;Light';
    }

    function esc(s) {
      return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
    }
    function appendLog(text, type) {
      const log = document.getElementById('log');
      const ts = new Date().toLocaleTimeString('en-US', {hour12: false});
      log.innerHTML +=
        '<span class="log-TS">[' + ts + ']</span> ' +
        '<span class="log-' + type + '">' + esc(text) + '</span>\n';
      log.scrollTop = log.scrollHeight;
    }
    function clearLog() {
      document.getElementById('log').innerHTML = '';
      setStatus('ready');
    }
    function setStatus(msg, busy) {
      const el = document.getElementById('logStatus');
      const dot = document.getElementById('statusDot');
      el.textContent = msg;
      el.classList.toggle('busy', !!busy);
      dot.classList.toggle('busy', !!busy);
    }

    const ACTION_BTNS = [
      'runTestingBtn','runAssessmentsBtn','runRecommendationsBtn',
      'runAppendixBtn','runFindingsBtn','runSeverityBtn'
    ];
    function setBusy(busy) {
      ACTION_BTNS.forEach(id => {
        const b = document.getElementById(id);
        b.disabled = busy;
        b.classList.toggle('loading', busy);
      });
      setStatus(busy ? 'running\u2026' : 'done', busy);
    }

    function browseFile(inputId, ext) {
      fetch('/browse?ext=' + encodeURIComponent(ext))
        .then(r => r.json())
        .then(d => { if (d.path) document.getElementById(inputId).value = d.path; })
        .catch(() => {});
    }

    async function runAction(url, payload, header) {
      setBusy(true);
      appendLog('\u2500\u2500 ' + header + ' \u2500\u2500', 'HEADER');
      Object.entries(payload).forEach(([k, v]) => appendLog(k + '  \u2192 ' + v, 'INFO'));
      try {
        const resp = await fetch(url, {
          method: 'POST',
          headers: {'Content-Type': 'application/json'},
          body: JSON.stringify(payload)
        });
        const data = await resp.json();
        data.logs.forEach(({text, type}) => appendLog(text, type));
      } catch(e) {
        appendLog('\u274c  request failed: ' + e, 'WARN');
      } finally {
        setBusy(false);
      }
    }

    async function runTesting() {
      const activity  = document.getElementById('activityPath').value.trim();
      const findings  = document.getElementById('findingsPath').value.trim();
      const pageRange = document.getElementById('pageRange').value.trim();
      if (!activity || !findings) { appendLog('\u26a0  Activity Report and Findings Report paths are required.', 'WARN'); return; }
      if (!pageRange)              { appendLog('\u26a0  Activity Log page range (P1) is required.', 'WARN'); return; }
      await runAction('/run_testing', {activity, findings, pageRange}, 'automated testing activity');
    }

    async function runAssessments() {
      const executive  = document.getElementById('executivePath').value.trim();
      const findings   = document.getElementById('findingsPath').value.trim();
      const pageNumber = document.getElementById('pageNumber').value.trim();
      if (!executive || !findings) { appendLog('\u26a0  Executive Report and Findings Report paths are required.', 'WARN'); return; }
      if (!pageNumber)             { appendLog('\u26a0  Assessment Results page number (P2) is required.', 'WARN'); return; }
      await runAction('/run_assessments', {executive, findings, pageNumber}, 'assessment results');
    }

    async function runRecommendations() {
      const recommendation = document.getElementById('recommendationPath').value.trim();
      const technical      = document.getElementById('technicalPath').value.trim();
      const findings       = document.getElementById('findingsPath').value.trim();
      const environment    = document.getElementById('environment').value.trim();
      if (!recommendation || !technical || !findings) { appendLog('\u26a0  Recommendations XLSX, Technical Report, and Findings Report paths are required.', 'WARN'); return; }
      if (!environment) { appendLog('\u26a0  Environment (P3) is required (e.g. Internal or External).', 'WARN'); return; }
      await runAction('/run_recommendations', {recommendation, technical, findings, environment}, 'recommendations');
    }

    async function runAppendix() {
      const technical = document.getElementById('technicalPath').value.trim();
      const findings  = document.getElementById('findingsPath').value.trim();
      if (!technical || !findings) { appendLog('\u26a0  Technical Report and Findings Report paths are required.', 'WARN'); return; }
      await runAction('/run_appendix', {technical, findings}, 'appendix');
    }

    async function runFindings() {
      const technical      = document.getElementById('technicalPath').value.trim();
      const findings       = document.getElementById('findingsPath').value.trim();
      const recommendation = document.getElementById('recommendationPath').value.trim();
      if (!technical || !findings || !recommendation) { appendLog('\u26a0  Technical Report, Findings Report, and Recommendations XLSX paths are required.', 'WARN'); return; }
      await runAction('/run_findings', {technical, findings, recommendation}, 'findings details');
    }

    async function runSeverity() {
      const technical = document.getElementById('technicalPath').value.trim();
      if (!technical) { appendLog('\u26a0  Technical Report path is required.', 'WARN'); return; }
      await runAction('/run_severity', {technical}, 'vulnerability count');
    }

    appendLog('report automation tool ready.', 'INIT');
  </script>
</body>
</html>
"""


# ── Helper ────────────────────────────────────────────────────────────────────
def _line_type(line: str) -> str:
    if "\u2705" in line:
        return "OK"
    if "\u274c" in line or "error" in line.lower():
        return "WARN"
    return "PLAIN"


def _capture_and_run(fn, *args):
    buf = io.StringIO()
    old = sys.stdout
    sys.stdout = buf
    exc = None
    try:
        fn(*args)
    except Exception as e:
        exc = e
    finally:
        sys.stdout = old
    logs = [
        {"text": line, "type": _line_type(line)}
        for line in buf.getvalue().splitlines()
        if line.strip()
    ]
    return logs, exc


# ── Routes ────────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template_string(
        HTML,
        activity_default=DEFAULT_ACTIVITY_REPORT_PATH,
        findings_default=DEFAULT_FINDINGS_REPORT_PATH,
        technical_default=DEFAULT_TECHNICAL_REPORT_PATH,
        executive_default=DEFAULT_EXECUTIVE_REPORT_PATH,
        recommendation_default=DEFAULT_RECOMMENDATIONS_PATH,
    )


@app.route("/browse")
def browse():
    ext = request.args.get("ext", "")
    if ext == "pdf":
        filetypes = [("PDF files", "*.pdf")]
    elif ext == "docx":
        filetypes = [("Word files", "*.docx")]
    elif ext == "xlsx":
        filetypes = [("Excel files", "*.xlsx")]
    else:
        filetypes = [("All files", "*.*")]
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.askopenfilename(filetypes=filetypes, parent=root)
    root.destroy()
    return jsonify({"path": path or ""})


@app.route("/run_testing", methods=["POST"])
def run_testing():
    data       = request.get_json(force=True)
    activity   = (data.get("activity")  or "").strip()
    findings   = (data.get("findings")  or "").strip()
    page_range = (data.get("pageRange") or "").strip()
    logs, exc  = _capture_and_run(automated_testing_activity, activity, findings, page_range)
    if exc:
        logs.append({"text": f"\u274c Error: {exc}", "type": "WARN"})
    else:
        logs.append({"text": f"\u2705 Complete. Output saved to: {findings}", "type": "OK"})
    return jsonify({"logs": logs})


@app.route("/run_assessments", methods=["POST"])
def run_assessments():
    data        = request.get_json(force=True)
    executive   = (data.get("executive")  or "").strip()
    findings    = (data.get("findings")   or "").strip()
    page_number = (data.get("pageNumber") or "").strip()
    logs, exc   = _capture_and_run(assessment_results, executive, findings, page_number)
    if exc:
        logs.append({"text": f"\u274c Error: {exc}", "type": "WARN"})
    else:
        logs.append({"text": f"\u2705 Complete. Output saved to: {findings}", "type": "OK"})
    return jsonify({"logs": logs})


@app.route("/run_recommendations", methods=["POST"])
def run_recommendations():
    data           = request.get_json(force=True)
    recommendation = (data.get("recommendation") or "").strip()
    technical      = (data.get("technical")      or "").strip()
    findings       = (data.get("findings")       or "").strip()
    environment    = (data.get("environment")    or "").strip()
    logs, exc      = _capture_and_run(recommendations, recommendation, technical, findings, environment)
    if exc:
        logs.append({"text": f"\u274c Error: {exc}", "type": "WARN"})
    else:
        logs.append({"text": f"\u2705 Complete. Output saved to: {findings}", "type": "OK"})
    return jsonify({"logs": logs})


@app.route("/run_appendix", methods=["POST"])
def run_appendix():
    data      = request.get_json(force=True)
    technical = (data.get("technical") or "").strip()
    findings  = (data.get("findings")  or "").strip()
    logs, exc = _capture_and_run(appendix, technical, findings)
    if exc:
        logs.append({"text": f"\u274c Error: {exc}", "type": "WARN"})
    else:
        logs.append({"text": f"\u2705 Complete. Output saved to: {findings}", "type": "OK"})
    return jsonify({"logs": logs})


@app.route("/run_findings", methods=["POST"])
def run_findings():
    data           = request.get_json(force=True)
    technical      = (data.get("technical")      or "").strip()
    findings       = (data.get("findings")       or "").strip()
    recommendation = (data.get("recommendation") or "").strip()
    logs, exc      = _capture_and_run(finding_details, technical, findings, recommendation, "External")
    if exc:
        logs.append({"text": f"\u274c Error: {exc}", "type": "WARN"})
    else:
        logs.append({"text": f"\u2705 Complete. Output saved to: {findings}", "type": "OK"})
    return jsonify({"logs": logs})


@app.route("/run_severity", methods=["POST"])
def run_severity():
    data      = request.get_json(force=True)
    technical = (data.get("technical") or "").strip()
    logs      = []
    buf       = io.StringIO()
    old       = sys.stdout
    sys.stdout = buf
    try:
        frames = severity_counter(technical)
        sys.stdout = old
        for line in buf.getvalue().splitlines():
            if line.strip():
                logs.append({"text": line, "type": _line_type(line)})
        counts: dict = {}
        for f in frames:
            counts[f.rating] = counts.get(f.rating, 0) + 1
        logs.append({"text": f"Found {len(frames)} total vulnerabilities:", "type": "OK"})
        for sev in ["Critical", "High", "Medium", "Low", "Informational"]:
            if sev in counts:
                logs.append({"text": f"  {sev:<16} {counts[sev]}", "type": "PLAIN"})
        for frame in frames:
            logs.append({"text": f"  [{frame.rating}]  {frame.discoveredName}", "type": "PLAIN"})
    except Exception as exc:
        sys.stdout = old
        for line in buf.getvalue().splitlines():
            if line.strip():
                logs.append({"text": line, "type": _line_type(line)})
        logs.append({"text": f"\u274c Error: {exc}", "type": "WARN"})
    return jsonify({"logs": logs})


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    threading.Timer(1.0, lambda: webbrowser.open("http://localhost:5000")).start()
    app.run(host="127.0.0.1", port=5000, debug=False, use_reloader=False)
