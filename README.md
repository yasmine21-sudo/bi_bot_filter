# PBIRS Daily Screenshot Automation

This project opens a PBIRS report, discovers every sheet dynamically, captures a screenshot for each sheet, and emails the images once per day.

For a step-by-step environment and server deployment guide, see [SETUP_AND_HOSTING_GUIDE.md](./SETUP_AND_HOSTING_GUIDE.md).

## What it does

- Uses Playwright with Chromium in headless mode
- Waits for the report to finish rendering before each screenshot
- Detects report sheets dynamically
- Saves one full-page PNG screenshot per sheet locally
- Emails the screenshots as attachments
- Logs every run to `logs/pbirs_capture.log`
- Includes Windows Task Scheduler scripts that run through `.venv`

## Authentication support

The report probe returned `401 Unauthorized`, so authentication is required for your environment.

- `AUTH_MODE=integrated`: Best fit for Windows-integrated PBIRS access. Run the scheduled task as the same domain or service account that is allowed to open the report.
- `AUTH_MODE=basic`: Uses `HTTP_USERNAME` and `HTTP_PASSWORD`.
- `AUTH_MODE=form`: Uses `PBIRS_USERNAME` and `PBIRS_PASSWORD` against a visible login form.

## Setup

1. Create the virtual environment:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

2. Install the Python package dependency:

```powershell
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

3. Install the Playwright browser:

```powershell
python -m playwright install chromium
```

4. Create your config file:

```powershell
Copy-Item .env.example .env
```

5. Update `.env` with:

- `REPORT_URL`
- Authentication values
- SMTP values
- `EMAIL_TO`
- Optional browser and timeout tuning

6. Run a manual test:

```powershell
.\.venv\Scripts\python.exe .\main.py
```

To validate SMTP by itself without screenshots:

```powershell
.\.venv\Scripts\python.exe .\send_test_email.py
```

## Daily scheduling on Windows

Run the task registration script after `.venv` and `.env` are ready:

```powershell
.\scripts\register-daily-task.ps1 -TaskName "PBIRS Daily Capture" -StartTime "08:00"
```

This creates a daily scheduled task that calls `scripts/run-daily-job.ps1`, which in turn runs `main.py` through `.venv\Scripts\python.exe`.

If your PBIRS server uses Windows integrated auth, make sure the scheduled task runs as a domain or service account that can already access the report in a browser.

## Optional cron example

If you later move this to Linux, the equivalent cron entry is:

```cron
0 8 * * * /path/to/project/.venv/bin/python /path/to/project/main.py
```

## How tab detection works

The script uses two strategies:

1. Power BI API first

- It looks for the embedded Power BI report object via `window.powerbi.get(...)`.
- If found, it calls `getPages()` to enumerate the report sheets.
- It switches sheets with `setPage(...)`.

2. DOM fallback

- If the API is not exposed, it scans the page for visible tab-like elements such as `[role="tab"]`, `aria-selected`, and common Power BI tab classes.
- It sorts visible tabs by screen position so the capture order matches what a user sees.
- It clicks each discovered tab and waits for the report to settle before taking a screenshot.

## How screenshot capture works

- The browser uses a large viewport and `DEVICE_SCALE_FACTOR=2` by default for sharper output.
- After each tab switch, the script waits until spinners are gone and the report surface remains stable across multiple polling cycles.
- It expands the report iframe to the rendered sheet size and captures a full-page screenshot so the entire sheet is included.
- Screenshots are written to `output/YYYYMMDD/`.
- File names include the tab order, tab name, and timestamp.

## Reliability notes

- Increase `REPORT_RENDER_TIMEOUT_MS` if the report is slow.
- Increase `POST_TAB_CLICK_WAIT_MS` if a sheet contains heavy visuals.
- Review `logs/pbirs_capture.log` after the first run to confirm which auth path and tab-detection path were used.
- If your login page uses custom selectors, set `LOGIN_USERNAME_SELECTOR`, `LOGIN_PASSWORD_SELECTOR`, and `LOGIN_SUBMIT_SELECTOR` in `.env`.
