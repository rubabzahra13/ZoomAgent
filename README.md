# Zoom Meeting Summaries Downloader

Automatically downloads Zoom meeting summaries as Word documents (DOCX) and renames them based on **Topic** and **Date Created**.

## Installation

```bash
pip install -r requirements.txt
playwright install chromium
```

## Run on localhost

Start the web app and open it in your browser:

```bash
python3 app.py
```

Then go to **http://127.0.0.1:5000**, choose **All summaries** or **One calendar day** (optional date; empty = today), and click **Run**. A Chromium window will open for Zoom; log in if needed and the script will stream progress in the browser.

## Command-line usage

**Download every summary on every page (one pass):**

```bash
python3 zoom_summaries_auto_name.py
# or explicitly:
python3 zoom_summaries_auto_name.py --mode all
```

**Download only summaries created on one calendar day** (local date on this machine; from midnight through end of that day). Default day is **today**; optional `--date`:

```bash
python3 zoom_summaries_auto_name.py --mode daily
python3 zoom_summaries_auto_name.py --mode daily --date 2026-04-17
```

The script still walks every list page; rows whose Date Created is not on that day are skipped.

## How It Works

1. Opens a visible browser window and navigates to Zoom summaries page
2. Auto-logs in with saved credentials (or wait for manual login if 2FA required)
3. For each summary in the table:
   - Clicks the topic link to open detail page
   - Clicks "Open in Docs" to open in Zoom Docs
   - Clicks the 3-dot menu (...)
   - Clicks "Export" → "Word"
   - Saves the downloaded file with the format: `Topic_DateCreated.docx`
4. Returns to the list and processes the next summary

## Filename Format

Files are saved as: `Topic_DateCreated.docx`

**Example outputs:**
```
Brandscaling's_Zoom_Meeting_Jan_23,_2026_03-29_PM.docx
Brandscaling_-_Michelle_and_Richard_Jan_22,_2026_02-55_PM.docx
```

## Filename Cleaning

- Invalid characters (`/ \ : * ? " < > |`) are removed
- Spaces are replaced with underscores
- Colons in times (e.g., `02:32`) become dashes (`02-32`)
- Max filename length: 150 characters
- Duplicates get `_2`, `_3`, etc. appended

## Configuration

Set credentials in your environment (recommended) so they are not stored in the repo:

```bash
export ZOOM_EMAIL="your-email@example.com"
export ZOOM_PASSWORD="your-password"
```

If unset, the script still runs but **auto-login is disabled** — sign in manually in the Chromium window.

## Requirements

- Python 3.7+
- Playwright
- Chromium browser (installed via `playwright install chromium`)
