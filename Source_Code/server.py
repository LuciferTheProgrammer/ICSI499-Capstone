"""
Web server for the Report Automation Tool.
Run with: python server.py
Then open http://localhost:5000 in your browser.
"""
import os
import sys
import io
import json
import queue
import threading
from datetime import datetime

# Ensure AutomationPrototype is importable from the same directory
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from flask import Flask, Response, request, jsonify, render_template

from AutomationPrototype import (
    automated_testing_activity,
    severity_counter,
    assessment_results,
    recommendations,
    appendix,
    finding_details,
    Logistics,
    IPAddress,
    Host_Discovery,
    Narrative_Exploitation,
    Informational,
    Findings_Summary,
    DEFAULT_RECOMMENDATIONS_PATH,
    DEFAULT_EXECUTIVE_REPORT_PATH,
    DEFAULT_ACTIVITY_REPORT_PATH,
    DEFAULT_FINDINGS_REPORT_PATH,
    DEFAULT_TECHNICAL_REPORT_PATH,
)

app = Flask(__name__)

# ── Global state ──────────────────────────────────────────────────────────────
_output_queue: queue.Queue = queue.Queue()
_is_running = False
_running_lock = threading.Lock()


class _QueueWriter(io.TextIOBase):
    """Captures stdout and forwards lines to the SSE queue."""
    _SUPPRESS = "Consider using the pymupdf_layout package"

    def __init__(self, q: queue.Queue):
        self._q = q

    def write(self, text: str) -> int:
        if text and text.strip() and self._SUPPRESS not in text:
            self._q.put(text.rstrip("\n"))
        return len(text)

    def flush(self):
        pass


def _run_task(func, *args):
    global _is_running
    old_stdout = sys.stdout
    sys.stdout = _QueueWriter(_output_queue)
    try:
        func(*args)
        _output_queue.put("__OK__")
    except Exception as exc:
        _output_queue.put(f"❌ Error: {exc}")
        _output_queue.put("__DONE__")
    finally:
        sys.stdout = old_stdout
        with _running_lock:
            _is_running = False
        _output_queue.put("__DONE__")


def _start(func, *args):
    """Start func in a background thread if not already running."""
    global _is_running
    with _running_lock:
        if _is_running:
            return False
        _is_running = True
    threading.Thread(target=_run_task, args=(func, *args), daemon=True).start()
    return True


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/config")
def config():
    return jsonify({
        "activity":       DEFAULT_ACTIVITY_REPORT_PATH,
        "findings":       DEFAULT_FINDINGS_REPORT_PATH,
        "technical":      DEFAULT_TECHNICAL_REPORT_PATH,
        "executive":      DEFAULT_EXECUTIVE_REPORT_PATH,
        "recommendation": DEFAULT_RECOMMENDATIONS_PATH,
    })


@app.route("/api/status")
def status():
    return jsonify({"running": _is_running})


@app.route("/stream")
def stream():
    """Server-Sent Events endpoint – streams log output to the browser."""
    def generate():
        while True:
            try:
                msg = _output_queue.get(timeout=25)
                yield f"data: {json.dumps(msg)}\n\n"
            except queue.Empty:
                yield "data: \"__PING__\"\n\n"

    return Response(
        generate(),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.route("/api/run", methods=["POST"])
def run():
    data = request.get_json(force=True) or {}
    action = data.get("action", "")
    paths  = data.get("paths", {})
    extra  = data.get("extra", {})

    activity       = paths.get("activity", "").strip()
    findings       = paths.get("findings", "").strip()
    technical      = paths.get("technical", "").strip()
    executive      = paths.get("executive", "").strip()
    recommendation = paths.get("recommendation", "").strip()

    def err(msg):
        return jsonify({"error": msg}), 400

    def busy():
        return jsonify({"error": "A task is already running. Please wait."}), 409

    # ── Automated Testing ─────────────────────────────────────────────────────
    if action == "automated_testing":
        page_range = extra.get("page_range", "").strip()
        if not all([activity, findings, technical, page_range]):
            return err("activity, findings, technical paths and page_range are required")
        def task():
            automated_testing_activity(activity, findings, page_range)
            Informational(technical, findings)
        if not _start(task):
            return busy()

    # ── Assessment Results ────────────────────────────────────────────────────
    elif action == "assessment_results":
        page_number = extra.get("page_number", "").strip()
        if not all([executive, findings, page_number]):
            return err("executive, findings paths and page_number are required")
        if not _start(assessment_results, executive, findings, page_number):
            return busy()

    # ── Recommendations ───────────────────────────────────────────────────────
    elif action == "recommendations":
        if not all([recommendation, technical, findings]):
            return err("recommendation, technical and findings paths are required")
        if not _start(recommendations, recommendation, technical, findings):
            return busy()

    # ── Appendix ──────────────────────────────────────────────────────────────
    elif action == "appendix":
        if not all([technical, findings]):
            return err("technical and findings paths are required")
        if not _start(appendix, technical, findings):
            return busy()

    # ── Findings Details ──────────────────────────────────────────────────────
    elif action == "findings_details":
        if not all([recommendation, technical, findings]):
            return err("recommendation (xlsx), technical and findings paths are required")
        if not _start(finding_details, technical, findings, recommendation):
            return busy()

    # ── Logistics ─────────────────────────────────────────────────────────────
    elif action == "logistics":
        if not all([technical, findings]):
            return err("technical and findings paths are required")
        if not _start(Logistics, technical, findings):
            return busy()

    # ── IP Address ────────────────────────────────────────────────────────────
    elif action == "ip_address":
        if not all([technical, findings]):
            return err("technical and findings paths are required")
        if not _start(IPAddress, technical, findings):
            return busy()

    # ── Host Discovery ────────────────────────────────────────────────────────
    elif action == "host_discovery":
        if not all([technical, findings]):
            return err("technical and findings paths are required")
        if not _start(Host_Discovery, technical, findings):
            return busy()

    # ── Exploitation ──────────────────────────────────────────────────────────
    elif action == "exploitation":
        customer = extra.get("customer", "").strip()
        if not findings or not customer:
            return err("findings path and customer name are required")
        if not _start(Narrative_Exploitation, findings, customer):
            return busy()

    # ── Informational ─────────────────────────────────────────────────────────
    elif action == "informational":
        if not all([technical, findings]):
            return err("technical and findings paths are required")
        if not _start(Informational, technical, findings):
            return busy()

    # ── Findings Summary ──────────────────────────────────────────────────────
    elif action == "findings_summary":
        if not all([technical, findings]):
            return err("technical and findings paths are required")
        if not _start(Findings_Summary, technical, findings):
            return busy()

    # ── Count Vulnerabilities ─────────────────────────────────────────────────
    elif action == "count_vulnerabilities":
        if not technical:
            return err("technical path is required")
        def vuln_task():
            frames = severity_counter(technical)
            counts: dict = {}
            for f in frames:
                counts[f.rating] = counts.get(f.rating, 0) + 1
            print(f"Found {len(frames)} total vulnerabilities:")
            for sev in ["Critical", "High", "Medium", "Low", "Informational"]:
                if sev in counts:
                    print(f"  {sev:<16} {counts[sev]}")
            for frame in frames:
                print(f"  [{frame.rating}]  {frame.discoveredName}")
        if not _start(vuln_task):
            return busy()

    else:
        return err(f"Unknown action: {action}")

    return jsonify({"status": "started"})


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("Starting Report Automation Tool web server...")
    print("Open http://localhost:5000 in your browser")
    app.run(host="127.0.0.1", port=5000, debug=False, threaded=True)
