"""
Foreman 2-Week Lookahead Generator — Web App
=============================================
Flask app. Single file. Deploy anywhere Python runs.

Run locally:
    pip install flask openpyxl
    python app.py
    → open http://localhost:5000

Deploy to Railway / Render:
    See README at bottom of this file.
"""

import os
import sys
import json
import tempfile
import zipfile
from datetime import date, datetime
from pathlib import Path
from flask import Flask, render_template_string, request, send_file, jsonify, session
import openpyxl

# ── Ensure engine and parser are on path ─────────────────────────────────────
HERE = Path(__file__).parent
ENGINE_DIR = HERE
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(ENGINE_DIR))

from lookahead_engine import build as engine_build, _load_calendar  # type: ignore
from xml_parser import parse_msp_xml, summarise                     # type: ignore

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "lookahead-dev-key-change-in-prod")

# ── Default colours (user can override in UI) ─────────────────────────────────
DEFAULT_SECTION_COLOURS = {
    "EARTHWORKS":     "6E2F0B",
    "SOG":            "0B4D6E",
    "SUSPENDED SLAB": "0B6E3A",
    "STRUCTURE":      "4A235A",
    "FAÇADE":         "0E6655",
    "FITOUT":         "1A5276",
    "GENERAL":        "333333",
}

DEFAULT_ZONE_COLOURS = {
    "Zone 1":  "2C3E7A", "Zone 2": "1A5276",
    "Zone 3A": "117A65", "Zone 3B": "1E8449",
    "Zone 4":  "784212", "Zone 5": "6C3483",
}

# ── HTML template ─────────────────────────────────────────────────────────────
HTML = """
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Foreman Lookahead Generator</title>
<style>
  :root {
    --navy: #1A1A2E; --accent: #0F3460; --gold: #C9A84C;
    --white: #fff; --light: #F5F7FA; --border: #DDE3EC;
    --green: #006100; --amber: #9C5700; --red: #9C0006;
    --radius: 8px; --shadow: 0 2px 12px rgba(0,0,0,.10);
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font: 14px/1.6 'Segoe UI', Arial, sans-serif; background: var(--light); color: #222; }

  header {
    background: var(--navy); color: var(--white); padding: 18px 32px;
    display: flex; align-items: center; gap: 16px;
    box-shadow: 0 2px 8px rgba(0,0,0,.3);
  }
  header h1 { font-size: 18px; font-weight: 700; letter-spacing: .5px; }
  header span { color: var(--gold); font-size: 12px; margin-left: auto; }

  main { max-width: 820px; margin: 32px auto; padding: 0 16px 64px; }

  .card {
    background: var(--white); border-radius: var(--radius);
    box-shadow: var(--shadow); border: 1px solid var(--border);
    margin-bottom: 24px; overflow: hidden;
  }
  .card-head {
    background: var(--accent); color: var(--white);
    padding: 12px 20px; font-weight: 700; font-size: 13px;
    display: flex; align-items: center; gap: 8px;
  }
  .card-head .step {
    background: var(--gold); color: var(--navy);
    border-radius: 50%; width: 22px; height: 22px;
    display: flex; align-items: center; justify-content: center;
    font-size: 11px; font-weight: 900; flex-shrink: 0;
  }
  .card-body { padding: 20px; }

  label { display: block; font-weight: 600; font-size: 12px;
          color: #555; margin-bottom: 4px; text-transform: uppercase; letter-spacing: .4px; }
  input[type=text], input[type=file], select {
    width: 100%; padding: 9px 12px; border: 1px solid var(--border);
    border-radius: 6px; font-size: 13px; color: #222;
    background: #FAFBFC; transition: border .15s;
  }
  input[type=text]:focus, select:focus { outline: none; border-color: var(--accent); }
  .field { margin-bottom: 16px; }
  .hint { font-size: 11px; color: #888; margin-top: 3px; }
  .row2 { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }

  .upload-zone {
    border: 2px dashed var(--border); border-radius: var(--radius);
    padding: 24px; text-align: center; cursor: pointer; transition: all .15s;
    background: #FAFBFC;
  }
  .upload-zone:hover, .upload-zone.over { border-color: var(--accent); background: #EEF2FA; }
  .upload-zone input { display: none; }
  .upload-zone .icon { font-size: 28px; margin-bottom: 8px; }
  .upload-zone .label { font-size: 13px; color: #555; }
  .upload-zone .sub { font-size: 11px; color: #999; margin-top: 4px; }
  .file-chosen { font-size: 12px; color: var(--green); margin-top: 8px; font-weight: 600; }

  .btn {
    display: inline-flex; align-items: center; gap: 8px;
    padding: 11px 24px; border-radius: 6px; border: none;
    font-size: 13px; font-weight: 700; cursor: pointer; transition: all .15s;
    text-decoration: none;
  }
  .btn-primary { background: var(--accent); color: var(--white); }
  .btn-primary:hover { background: #0a2748; }
  .btn-success { background: #006100; color: var(--white); }
  .btn-success:hover { background: #004800; }
  .btn-lg { padding: 14px 32px; font-size: 14px; width: 100%; justify-content: center; }

  .progress { display: none; margin-top: 16px; }
  .progress-bar {
    height: 6px; background: var(--border); border-radius: 3px; overflow: hidden;
  }
  .progress-bar-inner {
    height: 100%; background: var(--accent);
    width: 0%; transition: width .3s; border-radius: 3px;
  }
  .progress-text { font-size: 12px; color: #555; margin-top: 6px; text-align: center; }

  .results { display: none; }
  .stat-row {
    display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 20px;
  }
  .stat {
    flex: 1; min-width: 110px; background: var(--light);
    border: 1px solid var(--border); border-radius: 6px;
    padding: 12px 16px; text-align: center;
  }
  .stat .num { font-size: 24px; font-weight: 900; color: var(--accent); }
  .stat .lbl { font-size: 10px; color: #888; text-transform: uppercase; letter-spacing: .5px; }
  .stat.green .num { color: var(--green); }
  .stat.amber .num { color: var(--amber); }
  .stat.red .num   { color: var(--red); }

  .download-area { text-align: center; padding: 12px 0; }
  .download-note { font-size: 11px; color: #888; margin-top: 10px; line-height: 1.5; }

  .section-list { font-size: 12px; color: #555; background: #F0F4FB;
    border-radius: 6px; padding: 10px 14px; margin-top: 8px; }
  .section-list span { display: inline-block; margin: 2px 4px;
    background: #DDE8F8; border-radius: 4px; padding: 2px 8px; font-size: 11px; }

  .error-box {
    background: #FFF0F0; border: 1px solid #F5C6C6; border-radius: 6px;
    padding: 12px 16px; color: #9C0006; font-size: 13px; margin-top: 12px; display: none;
  }

  .divider { border: none; border-top: 1px solid var(--border); margin: 20px 0; }

  .steps-mini { display: flex; gap: 8px; margin-bottom: 20px; flex-wrap: wrap; }
  .steps-mini .s {
    flex: 1; min-width: 160px; background: #F0F4FB; border-radius: 6px;
    padding: 10px 12px; font-size: 11px; color: #444; border-left: 3px solid var(--accent);
  }
  .steps-mini .s strong { display: block; font-size: 12px; color: var(--navy); margin-bottom: 2px; }
</style>
</head>
<body>

<header>
  <h1>📋 Foreman Lookahead Generator</h1>
  <span>2-Week Rolling Lookahead · Excel + Macros</span>
</header>

<main>

  <div class="steps-mini">
    <div class="s"><strong>1. Upload files</strong>MS Project XML + WD Calendar</div>
    <div class="s"><strong>2. Name the job</strong>Project name, subtitle, output filename</div>
    <div class="s"><strong>3. Generate</strong>Excel ready in seconds</div>
    <div class="s"><strong>4. Add macros</strong>Paste .bas into Excel VBA editor</div>
  </div>

  <form id="mainForm" enctype="multipart/form-data">

    <!-- ── STEP 1: Files ─────────────────────────────────────────────── -->
    <div class="card">
      <div class="card-head"><span class="step">1</span> Upload Programme &amp; Calendar</div>
      <div class="card-body">
        <div class="row2">
          <div class="field">
            <label>MS Project XML Export</label>
            <div class="upload-zone" id="zone-xml" onclick="document.getElementById('f-xml').click()">
              <div class="icon">📁</div>
              <div class="label">Click to select file</div>
              <div class="sub">.xml exported from MS Project</div>
              <input type="file" id="f-xml" name="xml_file" accept=".xml">
            </div>
            <div class="file-chosen" id="xml-name"></div>
          </div>
          <div class="field">
            <label>Working Day Calendar</label>
            <div class="upload-zone" id="zone-cal" onclick="document.getElementById('f-cal').click()">
              <div class="icon">📅</div>
              <div class="label">Click to select file</div>
              <div class="sub">.xlsx — Table sheet, col D=date, col F=type</div>
              <input type="file" id="f-cal" name="cal_file" accept=".xlsx,.xls">
            </div>
            <div class="file-chosen" id="cal-name"></div>
          </div>
        </div>
      </div>
    </div>

    <!-- ── STEP 2: Project details ───────────────────────────────────── -->
    <div class="card">
      <div class="card-head"><span class="step">2</span> Project Details</div>
      <div class="card-body">
        <div class="row2">
          <div class="field">
            <label>Project Name</label>
            <input type="text" name="project_name" placeholder="e.g. BOILER HOUSE" value="{{ default_name }}">
          </div>
          <div class="field">
            <label>Project Subtitle</label>
            <input type="text" name="project_subtitle" placeholder="e.g. 60 Chandler Hwy, Alphington" value="{{ default_subtitle }}">
          </div>
        </div>
        <div class="row2">
          <div class="field">
            <label>Output Filename (Excel)</label>
            <input type="text" name="output_filename" placeholder="e.g. BH_Foreman_Lookahead.xlsx" value="{{ default_xlsx }}">
            <div class="hint">Will be saved as .xlsx — remind foreman to save as .xlsm to use macros</div>
          </div>
          <div class="field">
            <label>MS Project Outline Level for Activities</label>
            <select name="outline_level">
              <option value="4">Level 4</option>
              <option value="5" selected>Level 5 (default)</option>
              <option value="6">Level 6</option>
              <option value="3">Level 3</option>
            </select>
            <div class="hint">Outline level where activity rows live in your programme</div>
          </div>
        </div>
      </div>
    </div>

    <!-- ── Generate button ───────────────────────────────────────────── -->
    <div class="card">
      <div class="card-body">
        <button type="button" class="btn btn-primary btn-lg" onclick="generate()">
          ⚡ Generate Lookahead
        </button>
        <div class="progress" id="progress">
          <div class="progress-bar"><div class="progress-bar-inner" id="pbar"></div></div>
          <div class="progress-text" id="ptext">Parsing programme…</div>
        </div>
        <div class="error-box" id="errbox"></div>
      </div>
    </div>

  </form>

  <!-- ── Results ───────────────────────────────────────────────────── -->
  <div class="results card" id="results">
    <div class="card-head" style="background:#006100">✅ Lookahead Generated!</div>
    <div class="card-body">
      <div class="stat-row" id="stat-row"></div>
      <div class="section-list" id="sections-found"></div>
      <hr class="divider">
      <div class="download-area">
        <a class="btn btn-success btn-lg" id="dl-link" href="#" download>
          ⬇ Download Lookahead Package (.zip)
        </a>
        <div class="download-note">
          Package includes: <strong>Excel workbook</strong> + <strong>VBA macros (.bas)</strong><br>
          Open the .xlsx → Save As .xlsm → Alt+F11 → Insert Module → paste .bas → assign buttons
        </div>
      </div>
    </div>
  </div>

</main>

<script>
// File picker labels
["xml","cal"].forEach(id => {
  const inp = document.getElementById("f-" + id);
  inp.addEventListener("change", () => {
    const nm = document.getElementById(id + "-name");
    nm.textContent = inp.files[0] ? "✓ " + inp.files[0].name : "";
  });
});

// Drag-over highlight
document.querySelectorAll(".upload-zone").forEach(z => {
  z.addEventListener("dragover", e => { e.preventDefault(); z.classList.add("over"); });
  z.addEventListener("dragleave", () => z.classList.remove("over"));
  z.addEventListener("drop", e => {
    e.preventDefault(); z.classList.remove("over");
    const inp = z.querySelector("input[type=file]");
    if (e.dataTransfer.files.length) {
      inp.files = e.dataTransfer.files;
      inp.dispatchEvent(new Event("change"));
    }
  });
});

function setProgress(pct, text) {
  document.getElementById("pbar").style.width = pct + "%";
  document.getElementById("ptext").textContent = text;
}

function showError(msg) {
  const b = document.getElementById("errbox");
  b.textContent = "⚠ " + msg;
  b.style.display = "block";
  document.getElementById("progress").style.display = "none";
}

async function generate() {
  const xmlFile = document.getElementById("f-xml").files[0];
  const calFile = document.getElementById("f-cal").files[0];
  if (!xmlFile) return showError("Please select an MS Project XML file.");
  if (!calFile) return showError("Please select a Working Day Calendar (.xlsx) file.");

  document.getElementById("errbox").style.display = "none";
  document.getElementById("results").style.display = "none";
  document.getElementById("progress").style.display = "block";
  setProgress(10, "Uploading files…");

  const fd = new FormData(document.getElementById("mainForm"));
  setProgress(30, "Parsing MS Project XML…");

  try {
    const resp = await fetch("/generate", { method: "POST", body: fd });
    setProgress(80, "Building Excel workbook…");

    if (!resp.ok) {
      const err = await resp.json().catch(() => ({ error: "Server error " + resp.status }));
      return showError(err.error || "Unknown error");
    }

    const data = await resp.json();
    setProgress(100, "Done!");

    // Stats
    const s = data.summary;
    document.getElementById("stat-row").innerHTML = `
      <div class="stat"><div class="num">${s.total}</div><div class="lbl">Activities</div></div>
      <div class="stat green"><div class="num">${s.in_progress}</div><div class="lbl">In Progress</div></div>
      <div class="stat amber"><div class="num">${s.starting_soon}</div><div class="lbl">Starting Soon</div></div>
      <div class="stat"><div class="num">${s.future}</div><div class="lbl">Future</div></div>
      <div class="stat"><div class="num">${s.past}</div><div class="lbl">Past/Complete</div></div>
    `;

    document.getElementById("sections-found").innerHTML =
      "<strong>Sections found:</strong> " +
      s.sections.map(sc => `<span>${sc}</span>`).join("") +
      " | <strong>Zones:</strong> " +
      s.zones.map(z => `<span>${z}</span>`).join("");

    document.getElementById("dl-link").href = "/download/" + data.token;
    document.getElementById("dl-link").download = data.filename;
    document.getElementById("progress").style.display = "none";
    document.getElementById("results").style.display = "block";

  } catch (e) {
    showError("Network error: " + e.message);
  }
}
</script>
</body>
</html>
"""


# ── In-memory job store (keyed by token) ──────────────────────────────────────
# For production use Redis or a temp-file approach. This is fine for single-server.
_jobs = {}


# ── Routes ────────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template_string(HTML,
        default_name="",
        default_subtitle="",
        default_xlsx="Foreman_Lookahead.xlsx",
    )


@app.route("/generate", methods=["POST"])
def generate():
    import traceback, secrets

    xml_file = request.files.get("xml_file")
    cal_file = request.files.get("cal_file")
    if not xml_file or not cal_file:
        return jsonify({"error": "Both files are required."}), 400

    project_name     = request.form.get("project_name", "PROJECT").strip().upper() or "PROJECT"
    project_subtitle = request.form.get("project_subtitle", "").strip()
    output_filename  = request.form.get("output_filename", "Foreman_Lookahead.xlsx").strip()
    vba_filename     = output_filename.replace(".xlsx", "_macros.bas")
    outline_level    = int(request.form.get("outline_level", "5"))

    if not output_filename.endswith(".xlsx"):
        output_filename += ".xlsx"

    tmpdir = tempfile.mkdtemp(prefix="lookahead_")

    try:
        # Save uploads
        xml_path = os.path.join(tmpdir, "programme.xml")
        cal_path = os.path.join(tmpdir, "calendar.xlsx")
        xml_file.save(xml_path)
        cal_file.save(cal_path)

        # Load calendar
        working, non_working = _load_calendar(cal_path, "Table")

        # Parse MS Project XML
        programme = parse_msp_xml(xml_path, working, outline_level=outline_level)
        if not programme:
            return jsonify({"error": (
                f"No activities found at OutlineLevel {outline_level} in the XML. "
                "Try a different outline level (the dropdown in Project Details)."
            )}), 400

        # Gather sections/zones for colour mapping
        sections = list(dict.fromkeys(a["section"] for a in programme))
        zones    = list(dict.fromkeys(a["zone"]    for a in programme))

        section_colours = {
            s: DEFAULT_SECTION_COLOURS.get(s, DEFAULT_SECTION_COLOURS["GENERAL"])
            for s in sections
        }
        zone_colours = {
            z: DEFAULT_ZONE_COLOURS.get(z, "555555")
            for z in zones
        }

        # Build a config-like object
        class Cfg:
            PROJECT_NAME     = project_name
            PROJECT_SUBTITLE = project_subtitle
            OUTPUT_FILENAME  = output_filename
            VBA_FILENAME     = vba_filename
            CALENDAR_PATH    = cal_path
            CALENDAR_SHEET   = "Table"
            SECTION_COLOURS  = section_colours
            ZONE_COLOURS     = zone_colours
            PROGRAMME        = programme

        xlsx_path, bas_path = engine_build(Cfg, output_dir=tmpdir, quiet=True)

        # Zip both files
        zip_name = output_filename.replace(".xlsx", "_package.zip")
        zip_path = os.path.join(tmpdir, zip_name)
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.write(xlsx_path, os.path.basename(xlsx_path))
            zf.write(bas_path,  os.path.basename(bas_path))

        # Store for download
        token = secrets.token_urlsafe(16)
        _jobs[token] = {"zip_path": zip_path, "filename": zip_name, "tmpdir": tmpdir}

        return jsonify({
            "token":    token,
            "filename": zip_name,
            "summary":  summarise(programme),
        })

    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/download/<token>")
def download(token):
    job = _jobs.get(token)
    if not job:
        return "Download expired or not found", 404
    return send_file(job["zip_path"], as_attachment=True,
                     download_name=job["filename"],
                     mimetype="application/zip")


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("DEBUG", "1") == "1"
    print(f"\n{'='*50}")
    print(f"  Foreman Lookahead Generator")
    print(f"  Running at: http://localhost:{port}")
    print(f"{'='*50}\n")
    app.run(host="0.0.0.0", port=port, debug=debug)


"""
══════════════════════════════════════════════════════════
DEPLOYMENT GUIDE
══════════════════════════════════════════════════════════

── LOCAL ──
pip install flask openpyxl
python app.py
→ http://localhost:5000

── RAILWAY (free tier) ──
1. Create account at railway.app
2. New Project → Deploy from GitHub repo
3. Add this file + lookahead_engine.py + xml_parser.py to a repo
4. Set start command: gunicorn app:app
5. Add requirements.txt:
      flask
      openpyxl
      gunicorn

── RENDER (free tier) ──
Same as Railway — connect GitHub repo, Render auto-detects Flask.

── SECRET_KEY ──
Set SECRET_KEY env var to a random string in production.

── NOTES ──
• _jobs dict is in-memory — restarting the server clears pending downloads.
  For production, use Redis or write zips to /tmp and clean up on a schedule.
• No auth — add Flask-Login or HTTP Basic Auth if you want to restrict access.
══════════════════════════════════════════════════════════
"""
