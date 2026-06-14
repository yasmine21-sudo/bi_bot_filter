# Setup And Hosting Guide

This guide walks through:

- configuring the `.env` file
- running the app manually
- deploying it to a company server
- scheduling it to run every day

The recommended target is a Windows server or Windows VM inside your company network, because the PBIRS URL is internal and the project already includes Windows Task Scheduler scripts.

## 1. What you need

Before deployment, make sure the server has:

- Windows Server or Windows 10/11 on the internal network
- Python 3.13 or another supported Python 3.x version
- network access to `http://10.20.10.63`
- network access to your internal SMTP server
- a domain or service account that is allowed to open the PBIRS report

Recommended account setup:

- Create a dedicated service account such as `DOMAIN\svc_pbirs_capture`
- Grant that account access to the PBIRS report
- Grant that account permission to send through your SMTP relay if required

If PBIRS uses Windows Integrated Authentication, this account is especially important because the scheduled task must run under the same identity that can access the report.

## 2. Copy the project to your server

Place the project in a stable folder such as:

```text
C:\Apps\pbirs-daily-capture
```

The folder should contain:

- `main.py`
- `requirements.txt`
- `.env.example`
- `scripts\run-daily-job.ps1`
- `scripts\register-daily-task.ps1`

## 3. Create the virtual environment

Open PowerShell in the project folder and run:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m playwright install chromium
```

If your company browser policies work better with Edge, keep `BROWSER_CHANNEL=msedge` in `.env`. If not, you can leave `BROWSER_CHANNEL` empty and let Playwright use bundled Chromium.

## 4. Create and configure `.env`

Create the runtime config file:

```powershell
Copy-Item .env.example .env
```

Then update `.env`.

## 5. `.env` field guide

### Core report settings

- `REPORT_URL`
  Use the PBIRS embed URL.
- `OUTPUT_DIR`
  Folder where screenshots are saved.
- `LOG_DIR`
  Folder where logs are written.
- `HEADLESS`
  Keep `true` for scheduled runs.
- `TIMEZONE`
  Used in file names and email timestamps.

Example:

```env
REPORT_URL=http://10.20.10.63/reports/powerbi/scrape?&rs:Embed=true
OUTPUT_DIR=output
LOG_DIR=logs
HEADLESS=true
TIMEZONE=Africa/Lagos
```

### Browser and screenshot quality

- `BROWSER_CHANNEL`
  Use `msedge` if Microsoft Edge is installed on the server and works well with company auth policies.
- `VIEWPORT_WIDTH` and `VIEWPORT_HEIGHT`
  Control the browser size.
- `DEVICE_SCALE_FACTOR`
  `2` gives sharper screenshots.
- `SCREENSHOT_PREFIX`
  Used in naming output files.

Recommended values:

```env
BROWSER_CHANNEL=msedge
VIEWPORT_WIDTH=1920
VIEWPORT_HEIGHT=1080
DEVICE_SCALE_FACTOR=2
SCREENSHOT_PREFIX=pbirs
```

### Timeout and reliability settings

- `NAVIGATION_TIMEOUT_MS`
  Max time for page navigation.
- `REPORT_RENDER_TIMEOUT_MS`
  Max time to wait for slow report rendering.
- `REPORT_STABLE_INTERVAL_MS`
  Polling interval while checking if the report has settled.
- `REPORT_STABLE_POLLS`
  Number of stable checks required before capture.
- `POST_TAB_CLICK_WAIT_MS`
  Extra delay after switching tabs.

Good starting values:

```env
NAVIGATION_TIMEOUT_MS=120000
REPORT_RENDER_TIMEOUT_MS=180000
REPORT_STABLE_INTERVAL_MS=2000
REPORT_STABLE_POLLS=3
POST_TAB_CLICK_WAIT_MS=5000
```

If a report is especially slow, raise `REPORT_RENDER_TIMEOUT_MS` to `240000` or `300000`.

### Authentication settings

Choose one auth mode:

- `AUTH_MODE=integrated`
- `AUTH_MODE=basic`
- `AUTH_MODE=form`
- `AUTH_MODE=none`

#### Recommended for company server: Integrated auth

Use this when PBIRS is protected by Windows authentication.

```env
AUTH_MODE=integrated
AUTH_SERVER_WHITELIST=10.20.10.63
HTTP_USERNAME=
HTTP_PASSWORD=
PBIRS_USERNAME=
PBIRS_PASSWORD=
LOGIN_USERNAME_SELECTOR=
LOGIN_PASSWORD_SELECTOR=
LOGIN_SUBMIT_SELECTOR=
```

Important:

- Run the scheduled task as the domain or service account that has PBIRS access.
- Test the report manually from that same account before relying on the automated run.

#### Basic auth example

```env
AUTH_MODE=basic
AUTH_SERVER_WHITELIST=
HTTP_USERNAME=your_username
HTTP_PASSWORD=your_password
```

#### Form login example

```env
AUTH_MODE=form
PBIRS_USERNAME=your_username
PBIRS_PASSWORD=your_password
LOGIN_USERNAME_SELECTOR=
LOGIN_PASSWORD_SELECTOR=
LOGIN_SUBMIT_SELECTOR=
```

If the login page uses custom fields, fill in the three selector values.

### Email settings

- `SMTP_HOST`
  Your internal mail relay or SMTP server
- `SMTP_PORT`
  Usually `25`, `587`, or `465` depending on company setup
- `SMTP_USE_TLS`
  Usually `true` for port `587`
- `SMTP_USERNAME` and `SMTP_PASSWORD`
  Only if your SMTP server requires login
- `SMTP_ENVELOPE_FROM`
  Optional SMTP envelope sender. On Exchange-like servers, this should usually be the same mailbox the authenticated account is allowed to send as.
- `EMAIL_FROM`
  Sender address
- `EMAIL_REPLY_TO`
  Optional reply-to address if replies should go somewhere else
- `EMAIL_TO`
  Comma-separated list of recipients
- `EMAIL_SUBJECT_PREFIX`
  Prefix used before the timestamp

Example:

```env
SMTP_HOST=smtp.company.local
SMTP_PORT=587
SMTP_USE_TLS=true
SMTP_USERNAME=
SMTP_PASSWORD=
SMTP_ENVELOPE_FROM=pbirs-bot@company.com
EMAIL_FROM=pbirs-bot@company.com
EMAIL_REPLY_TO=
EMAIL_TO=bi-team@company.com,ops-team@company.com
EMAIL_SUBJECT_PREFIX=PBIRS Daily Capture
```

### Schedule setting

This value is mainly for documentation and operations reference:

```env
SCHEDULE_TIME=08:00
```

The actual schedule is created in Windows Task Scheduler.

## 6. Recommended `.env` for a company server

If your company uses Windows-integrated auth and internal SMTP, this is the most likely configuration:

```env
REPORT_URL=http://10.20.10.63/reports/powerbi/scrape?&rs:Embed=true
OUTPUT_DIR=output
LOG_DIR=logs
HEADLESS=true
BROWSER_CHANNEL=msedge
VIEWPORT_WIDTH=1920
VIEWPORT_HEIGHT=1080
DEVICE_SCALE_FACTOR=2
NAVIGATION_TIMEOUT_MS=120000
REPORT_RENDER_TIMEOUT_MS=180000
REPORT_STABLE_INTERVAL_MS=2000
REPORT_STABLE_POLLS=3
POST_TAB_CLICK_WAIT_MS=5000
SCREENSHOT_PREFIX=pbirs
TIMEZONE=Africa/Lagos
AUTH_MODE=integrated
AUTH_SERVER_WHITELIST=10.20.10.63
HTTP_USERNAME=
HTTP_PASSWORD=
PBIRS_USERNAME=
PBIRS_PASSWORD=
LOGIN_USERNAME_SELECTOR=
LOGIN_PASSWORD_SELECTOR=
LOGIN_SUBMIT_SELECTOR=
SMTP_HOST=smtp.company.local
SMTP_PORT=587
SMTP_USE_TLS=true
SMTP_USERNAME=
SMTP_PASSWORD=
EMAIL_FROM=pbirs-bot@company.com
EMAIL_TO=bi-team@company.com
EMAIL_SUBJECT_PREFIX=PBIRS Daily Capture
SCHEDULE_TIME=08:00
```

## 7. Run the app manually

Use this before scheduling:

```powershell
.\.venv\Scripts\python.exe .\main.py
```

What to check after the run:

- screenshots appear under `output\YYYYMMDD\`
- log file updates under `logs\pbirs_capture.log`
- email arrives with PNG attachments

If the run fails, open:

- `logs\pbirs_capture.log`

## 8. Host it on your company server

For this app, “hosting” means installing it on an internal always-on server and running it there every day.

Recommended deployment pattern:

1. Use an internal Windows server or VM that stays online.
2. Install the app into `C:\Apps\pbirs-daily-capture`.
3. Run it under a dedicated service account.
4. Store `.env` only on the server, not in source control.
5. Use Task Scheduler for the daily run.
6. Keep `logs` and `output` on a disk with enough space.

Recommended server checklist:

- the server can open `http://10.20.10.63` over the network
- the service account can access the PBIRS report
- the server can reach the SMTP host and port
- PowerShell execution policy allows the scheduled script to run
- Microsoft Edge is installed if `BROWSER_CHANNEL=msedge`

## 9. Create the scheduled task on the server

Run:

```powershell
.\scripts\register-daily-task.ps1 -TaskName "PBIRS Daily Capture" -StartTime "08:00"
```

Then in Task Scheduler:

- open the task properties
- change the task user to your service account
- select “Run whether user is logged on or not”
- enable “Run with highest privileges” if your environment requires it

If you prefer to create the task manually, point it to:

```text
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Apps\pbirs-daily-capture\scripts\run-daily-job.ps1"
```

## 10. First production test on the server

After the task is created:

1. Run `scripts\run-daily-job.ps1` manually in PowerShell.
2. Confirm screenshots are created.
3. Confirm email delivery works.
4. Run the scheduled task once from Task Scheduler.
5. Re-check `logs\pbirs_capture.log`.

## 11. Operations and maintenance

Recommended maintenance steps:

- review the log file after the first few runs
- clean old screenshots if retention is not needed
- keep the `.venv` and browser updated during planned maintenance windows
- re-test after PBIRS layout changes
- update `.env` if SMTP or recipients change

Optional retention cleanup can be handled by a separate scheduled PowerShell script if you want to keep only the last `N` days of screenshots.

## 12. Troubleshooting

### 401 Unauthorized

Usually means:

- the task is running under the wrong account
- `AUTH_MODE` is wrong
- the report server requires credentials you did not provide

### No tabs were found

Usually means:

- the report did not finish loading
- PBIRS layout changed
- authentication redirected the page somewhere else

Try increasing:

- `REPORT_RENDER_TIMEOUT_MS`
- `POST_TAB_CLICK_WAIT_MS`

### Email not sent

Check:

- `SMTP_HOST`
- `SMTP_PORT`
- `SMTP_USE_TLS`
- firewall rules between the server and the SMTP host
- whether your SMTP relay requires authentication

### Browser launch issues

If `msedge` causes startup issues, clear this value:

```env
BROWSER_CHANNEL=
```

Then reinstall the Playwright browser and test again.

## 13. Suggested server rollout

For a clean production rollout, use this order:

1. Deploy project files to the server.
2. Create `.venv`.
3. Install Python dependencies.
4. Install Playwright Chromium.
5. Create `.env`.
6. Test a manual run.
7. Create the scheduled task.
8. Switch the task to the service account.
9. Trigger the task once.
10. Confirm logs, screenshots, and email delivery.
