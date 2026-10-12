# Foreman Lookahead Generator — Web App

Upload an MS Project XML + Working Day Calendar → download a ready-made 2-week foreman lookahead Excel in seconds.

## Files in this package

| File | Purpose |
|---|---|
| `app.py` | Flask web app — run this |
| `xml_parser.py` | Parses MS Project .xml exports into activity lists |
| `lookahead_engine.py` | Builds the Excel workbook (generic engine) |
| `requirements.txt` | Python dependencies |
| `configs/boiler_house.py` | Example manual config (for reference / override) |
| `build.py` | CLI build script (bypasses web UI) |

## Run locally

```bash
pip install flask openpyxl
python app.py
# Open http://localhost:5000
```

## Deploy to Railway (free, 5 minutes)

1. Create account at https://railway.app
2. New Project → Deploy from GitHub Repo
3. Push all files to a GitHub repo
4. Railway auto-detects Flask — set start command: `gunicorn app:app`
5. Done — Railway gives you a public URL

## Deploy to Render (free)

Same as Railway — connect your GitHub repo, Render auto-detects Python/Flask.

## Using the app

1. Export your programme from MS Project: File → Save As → XML
2. Upload the XML + your Working Day Calendar (.xlsx)
3. Fill in project name, subtitle
4. Hit Generate → download the .zip
5. Open the Excel, Save As .xlsm, add macros (instructions in the .bas file)

## For a new job (CLI approach, no web UI)

```bash
cp configs/boiler_house.py configs/new_job.py
# Edit new_job.py — project name, calendar path, PROGRAMME list
python build.py configs/new_job.py
```

## Outline Level

The app defaults to OutlineLevel 5 for activity rows. If your XML uses a different 
level, change the dropdown in Project Details before generating.
