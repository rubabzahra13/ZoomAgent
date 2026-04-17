#!/usr/bin/env python3
"""
Zoom Downloader - Localhost web app to run the Zoom summaries downloader.
"""

import os
import subprocess
import sys
from pathlib import Path

from flask import Flask, Response, render_template_string, request

app = Flask(__name__)
SCRIPT_DIR = Path(__file__).resolve().parent
DOWNLOADER_SCRIPT = SCRIPT_DIR / "zoom_summaries_auto_name.py"


def run_downloader_stream(mode: str = "all", date_str: str = ""):
    """Run the downloader script and yield stdout line by line."""
    if not DOWNLOADER_SCRIPT.exists():
        yield f"data: Error: Script not found: {DOWNLOADER_SCRIPT}\n\n"
        return
    if mode not in ("all", "daily"):
        mode = "all"
    cmd = [sys.executable, "-u", str(DOWNLOADER_SCRIPT), "--mode", mode]
    if mode == "daily" and date_str.strip():
        cmd.extend(["--date", date_str.strip()])
    env = os.environ.copy()
    env["PYTHONUNBUFFERED"] = "1"
    try:
        proc = subprocess.Popen(
            cmd,
            cwd=str(SCRIPT_DIR),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
            env=env,
        )
        for line in iter(proc.stdout.readline, ""):
            yield f"data: {line}\n\n"
        proc.wait()
    except Exception as e:
        yield f"data: Error: {e}\n\n"


@app.route("/")
def index():
    return render_template_string(INDEX_HTML)


@app.route("/api/run")
def api_run():
    """Stream downloader output as Server-Sent Events."""
    mode = request.args.get("mode", "all")
    date_str = request.args.get("date", "") or ""
    return Response(
        run_downloader_stream(mode=mode, date_str=date_str),
        mimetype="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
            "Connection": "keep-alive",
        },
    )


INDEX_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Zoom Summaries Downloader</title>
  <style>
    * { box-sizing: border-box; }
    body {
      font-family: system-ui, -apple-system, sans-serif;
      max-width: 720px;
      margin: 0 auto;
      padding: 2rem;
      background: #0f0f12;
      color: #e4e4e7;
      min-height: 100vh;
    }
    h1 { font-size: 1.5rem; margin-bottom: 0.5rem; }
    .sub { color: #71717a; font-size: 0.9rem; margin-bottom: 1.5rem; }
    .card {
      background: #18181b;
      border: 1px solid #27272a;
      border-radius: 12px;
      padding: 1.5rem;
      margin-bottom: 1rem;
    }
    button {
      background: #3b82f6;
      color: white;
      border: none;
      padding: 0.75rem 1.5rem;
      border-radius: 8px;
      font-size: 1rem;
      cursor: pointer;
      font-weight: 500;
    }
    button:hover { background: #2563eb; }
    button:disabled { background: #52525b; cursor: not-allowed; }
    label.mode {
      display: flex;
      align-items: flex-start;
      gap: 0.5rem;
      margin: 0.5rem 0;
      cursor: pointer;
      font-size: 0.95rem;
    }
    label.mode input { margin-top: 0.2rem; }
    .date-row {
      margin-top: 0.75rem;
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      gap: 0.5rem;
    }
    input[type="date"] {
      background: #09090b;
      border: 1px solid #27272a;
      color: #e4e4e7;
      border-radius: 8px;
      padding: 0.5rem 0.75rem;
      font-family: inherit;
    }
    #log {
      background: #09090b;
      border: 1px solid #27272a;
      border-radius: 8px;
      padding: 1rem;
      font-family: ui-monospace, monospace;
      font-size: 0.85rem;
      line-height: 1.5;
      height: 360px;
      overflow-y: auto;
      white-space: pre-wrap;
      word-break: break-word;
    }
    #log:empty::before { content: "Log output will appear here…"; color: #52525b; }
    .success { color: #22c55e; }
    .error { color: #ef4444; }
  </style>
</head>
<body>
  <h1>Zoom Summaries Downloader</h1>
  <p class="sub">Run the downloader from your browser. A Chromium window will open for login and downloads.</p>

  <div class="card">
    <strong>Automation</strong>
    <label class="mode">
      <input type="radio" name="mode" value="all" checked>
      <span><strong>All summaries</strong> — one run downloads every summary on every list page.</span>
    </label>
    <label class="mode">
      <input type="radio" name="mode" value="daily">
      <span><strong>One calendar day</strong> — only rows whose Date Created is that local day (midnight–end of day).</span>
    </label>
    <div class="date-row" id="dateRow" style="display: none;">
      <label for="dayPicker">Day (optional):</label>
      <input type="date" id="dayPicker" title="Leave empty to use today on this computer">
      <span style="color: #71717a; font-size: 0.85rem;">Empty = today</span>
    </div>
    <p style="margin: 1rem 0 0;">
      <button id="runBtn">Run</button>
    </p>
    <p style="margin: 1rem 0 0; color: #71717a; font-size: 0.9rem;">
      Saves to <strong>Downloads</strong> as <code>Topic_DateCreated.docx</code>.
    </p>
  </div>

  <div class="card">
    <strong>Log</strong>
    <div id="log"></div>
  </div>

  <script>
    const runBtn = document.getElementById('runBtn');
    const logEl = document.getElementById('log');
    const dateRow = document.getElementById('dateRow');
    const dayPicker = document.getElementById('dayPicker');

    document.querySelectorAll('input[name="mode"]').forEach((el) => {
      el.addEventListener('change', () => {
        dateRow.style.display = el.value === 'daily' ? 'flex' : 'none';
      });
    });

    function appendLog(text, className = '') {
      const line = document.createElement('div');
      if (className) line.className = className;
      line.textContent = text;
      logEl.appendChild(line);
      logEl.scrollTop = logEl.scrollHeight;
    }

    function runUrl() {
      const mode = document.querySelector('input[name="mode"]:checked').value;
      const params = new URLSearchParams({ mode });
      if (mode === 'daily' && dayPicker.value) params.set('date', dayPicker.value);
      return '/api/run?' + params.toString();
    }

    runBtn.addEventListener('click', async () => {
      runBtn.disabled = true;
      logEl.innerHTML = '';
      appendLog('Starting downloader…');

      try {
        const res = await fetch(runUrl());
        const reader = res.body.getReader();
        const decoder = new TextDecoder();
        let buffer = '';

        while (true) {
          const { done, value } = await reader.read();
          if (done) break;
          buffer += decoder.decode(value, { stream: true });
          const lines = buffer.split('\\n');
          buffer = lines.pop() || '';
          for (const line of lines) {
            const m = line.match(/^data: (.*)$/);
            if (m) {
              const msg = m[1];
              const isErr = /error|✗|failed/i.test(msg);
              appendLog(msg, isErr ? 'error' : (msg.includes('✓') ? 'success' : ''));
            }
          }
        }
        if (buffer) {
          const m = buffer.match(/^data: (.*)$/);
          if (m) appendLog(m[1]);
        }
        appendLog('\\nDone.');
      } catch (e) {
        appendLog('Error: ' + e.message, 'error');
      }
      runBtn.disabled = false;
    });
  </script>
</body>
</html>
"""


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False, threaded=True)
