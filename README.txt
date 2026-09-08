# Trinity Sales Register — Automated Email Report
## Setup & Usage Guide

This project downloads the **Sales Register Mattresses** report from the
Ramco ERP and emails the exported `.xlsx` file **as-is** (no editing,
no breakup, no WhatsApp) to the configured recipients on a daily schedule
(07:00). It replaces the WhatsApp step of the original Sales Register
flow with email, using the OEE project's SMTP approach.

The scheduled job runs under the **Windows SYSTEM** account, so the
project is fully bundled (venv + Playwright browsers live inside the
project folder) — no reliance on any user's Python packages or profile.

---

### STEP 1 — Create the self-contained environment (once, as ADMIN)

Open a Command Prompt in this folder and run:

    C:\Python314\python.exe -m venv venv
    venv\Scripts\python.exe -m pip install -r requirements.txt
    set PLAYWRIGHT_BROWSERS_PATH=%~dp0.pw-browsers
    venv\Scripts\playwright.exe install chromium

This creates:
  - `venv\`         — isolated Python with `playwright`, `python-dotenv`, `openpyxl`
  - `.pw-browsers\` — the Chromium binaries Playwright uses

(Do this once after a fresh clone. `venv/` and `.pw-browsers/` are git-ignored.)

---

### STEP 2 — Verify config.env

Open `config.env` and confirm:

  ERP_URL          = https://peps.ramcoes.com/RVW/hub/index.html
  ERP_USERNAME     = Peps127              (RAMCO login id)
  ERP_PASSWORD     = peps@456             (RAMCO password)
  DOWNLOAD_DIR     = D:\Hari JR. DATA\Performance issues\Auto -reports\Trinity Sales Register\downloads

  SMTP_SERVER      = zimsmtp.logix.in
  SMTP_PORT        = 587                  (STARTTLS)
  SMTP_USER        = pepsit@pepsindia.com
  SMTP_PASSWORD    = Peps@151
  EMAIL_RECIPIENTS = trinity@pepsindia.com,hariit@pepsindia.com,
                     trinitysupermarketshubli@gmail.com           (To)
  EMAIL_CC         = sales@pepsindia.com,janaki@pepsindia.com,
                     itsupport@pepsindia.com                      (Cc)

  HEADLESS         = True   (unattended 7 AM runs; set False to watch)

  ERROR_EMAIL_RECIPIENTS = hariit@pepsindia.com   (error alert goes ONLY here)

> Note: `DOWNLOAD_DIR` is a project-local folder so the SYSTEM account
> can write the download into it (it could not write to
> `C:\Users\ADMIN\Downloads`). `config.env` is git-ignored — the real
> passwords never reach the repository.

Change nothing else unless you move to a different environment.

---

### STEP 3 — Run manually (first test)

    venv\Scripts\python.exe automate_report.py      # downloads the report
    venv\Scripts\python.exe email_sender.py         # emails the latest downloaded file

To send a **test email to one person only** (no CC leak):

    venv\Scripts\python.exe email_sender.py --to hariit@pepsindia.com --cc " "

---

### STEP 4 — Schedule daily at 7:00 AM (Windows Task Scheduler)

Right-click `register_task.bat` and choose **Run as administrator**.
This creates a daily 07:00 task named **"Trinity Sales Register Email"**
under the SYSTEM account, and then fixes the task action (full path +
working directory) because `schtasks /Create` splits paths that contain
spaces, which otherwise fails with error `0x80070002`.

Verify with:

    schtasks /Query /TN "Trinity Sales Register Email" /V

To remove later:

    schtasks /Delete /TN "Trinity Sales Register Email" /F

---

### STEP 5 — Schedule the watchdog at 7:40 AM (recommended)

Right-click `register_watchdog_task.bat` and choose **Run as
administrator**. This creates a daily 07:40 task named **"Trinity
Sales Register Watchdog"** under SYSTEM — ~40 minutes after the main
job, so it has time to finish all 6 download retries first.

The watchdog is an *independent* check: it verifies today's report was
actually downloaded and emailed, regardless of whether `run_task.bat`
itself ran, crashed, or was never triggered at all. If it wasn't, it
alerts `hariit` and then tries to download + email the report itself.
If that recovery also fails, it sends a second alert. See `watchdog.py`.

Verify with:

    schtasks /Query /TN "Trinity Sales Register Watchdog" /V

To remove later:

    schtasks /Delete /TN "Trinity Sales Register Watchdog" /F

---

### WHAT IT DOES

1. `run_task.bat` first syncs this folder to `origin/main` (`git fetch`
   + `git reset --hard`) so the scheduled run always executes the
   real, committed code — never a stale or accidentally-edited local
   copy. `config.env`/`downloads/`/`venv/` are gitignored and untouched.
2. It then runs `automate_report.py` (a project-local
   `venv\Scripts\python.exe`), retrying the download up to **6 times**
   (the ERP login is flaky), writing progress to `run_log.txt`.
3. `automate_report.py` logs into Ramco ERP with the configured
   credentials, opens the **Sales Register Mattresses** report, exports
   the current month to an `.xlsx` in `DOWNLOAD_DIR`, renames it to
   `Trinity Sales Register -DD.MM.YYYY.xlsx`, and fixes the ERP's broken
   xlsx relationship paths so Excel can open it.
   - Kept the fix that replaces `networkidle` waits with
     `domcontentloaded` (the earlier hang), and dismisses the
     "already logged in" session dialog.
4. `email_sender.py` refuses to send if the newest matching file isn't
   dated today (protects against ever re-sending a stale report), then
   attaches that file (unchanged) and sends it via SMTP to `trinity`
   and `hariit` (To) with `sales`, `janaki`, `itsupport` as Cc. It then
   archives exports older than 7 days into `downloads\archive\`, and
   records the send in `.last_sent` (used by the watchdog).
5. If **all** download attempts fail (e.g. ERP slow/unavailable), or if
   `email_sender.py` itself fails, an **error alert is emailed only to
   `hariit@pepsindia.com`** with a sentence stating the actual failure
   reason (from `last_error.txt`).
6. The watchdog (Step 5) independently double-checks the whole thing
   ~40 minutes later and self-heals if it didn't happen.

`python` is invoked with `-u` so the logs flush in real time.

---

### NOTES / KNOWN LIMITATIONS

- The exported file is emailed exactly as downloaded — no filtering or
  re-formatting (unlike the Sales Register WhatsApp breakup).
- **Recipients matter**: `hariit` is in the `To` field (not only `Cc`)
  because the mail server dropped the `Cc` copy in testing; `To`
  delivery was verified working.
- If Ramco updates its UI/selectors, the navigation code in
  `automate_report.py` may need adjustment.
- DURATION_LOG_FILE is optional; if set, the run time is also logged to
  the Sales Register duration workbook.
- Large monthly reports can be ~14 MB; the mail server accepted this
  size in testing, but very large months could approach attachment
  limits.
- The scheduled task has "disallow start on battery" enabled — keep the
  PC on AC power at 07:00.
