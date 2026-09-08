"""
Trinity Sales Register — Watchdog.

Runs ~40 minutes after the main scheduled job (run_task.bat, 7:00 AM).
Checks whether today's report actually got downloaded AND emailed. If
not — for ANY reason (run_task.bat itself crashing before reaching the
Python steps, the ERP download failing, the email failing, the machine
being off, etc.) — it alerts hariit and then tries to recover by
running the download + email itself.

This exists because run_task.bat can fail in ways that never reach the
Python-level error-alerting in email_sender.py (see the 08-Sep-2026
incident: a batch-syntax bug crashed run_task.bat before it ever
called automate_report.py, so nothing detected or reported the
failure). This is an independent, outside check.
"""

import os
import glob
import logging
import subprocess
from pathlib import Path
from datetime import date, datetime

from dotenv import load_dotenv

SCRIPT_DIR = Path(__file__).parent
load_dotenv(dotenv_path=SCRIPT_DIR / "config.env")

logging.basicConfig(level=logging.INFO, format="%(asctime)s [WATCHDOG] %(levelname)s %(message)s")
log = logging.getLogger("trinity_watchdog")

DOWNLOAD_DIR   = Path(os.getenv("DOWNLOAD_DIR", r"C:\Users\ADMIN\Downloads"))
LAST_SENT_FILE = SCRIPT_DIR / ".last_sent"
PYTHON         = SCRIPT_DIR / "venv" / "Scripts" / "python.exe"
if not PYTHON.exists():
    PYTHON = Path("python")

DOWNLOAD_RETRY_ATTEMPTS = 3


def _file_date(p) -> date:
    base = Path(p).stem
    part = base.split(" -")[-1] if " -" in base else ""
    try:
        return datetime.strptime(part, "%d.%m.%Y").date()
    except ValueError:
        return date.min


def today_was_downloaded() -> bool:
    matches = glob.glob(str(DOWNLOAD_DIR / "Trinity Sales Register -*.xlsx"))
    return any(_file_date(p) == date.today() for p in matches)


def today_was_emailed() -> bool:
    """email_sender.py writes today's date to .last_sent right after a
    successful smtplib.sendmail(). Checking that file rather than
    grepping run_log.txt means this works regardless of how
    email_sender.py was invoked (scheduled run, watchdog recovery, or
    an interactive/manual run whose output was never redirected into
    run_log.txt — grepping the log missed exactly that case once and
    caused a false-positive "failure" that triggered a real duplicate
    send)."""
    if not LAST_SENT_FILE.exists():
        return False
    try:
        text = LAST_SENT_FILE.read_text(encoding="utf-8").strip()
        return date.fromisoformat(text) == date.today()
    except (OSError, ValueError):
        return False


def run_python(*args) -> int:
    result = subprocess.run(
        [str(PYTHON), "-u", *args],
        cwd=str(SCRIPT_DIR),
        capture_output=True,
        text=True,
    )
    if result.stdout:
        log.info(result.stdout.strip())
    if result.stderr:
        log.warning(result.stderr.strip())
    return result.returncode


def attempt_recovery() -> bool:
    log.info("Attempting recovery: running automate_report.py + email_sender.py directly.")

    if today_was_downloaded():
        log.info("Today's file already exists on disk — skipping re-download, just emailing.")
    else:
        for attempt in range(1, DOWNLOAD_RETRY_ATTEMPTS + 1):
            log.info(f"Recovery download attempt {attempt} of {DOWNLOAD_RETRY_ATTEMPTS}...")
            code = run_python(str(SCRIPT_DIR / "automate_report.py"))
            if code == 0:
                log.info(f"Recovery download succeeded on attempt {attempt}.")
                break
            log.warning(f"Recovery download attempt {attempt} failed (exit {code}).")
        else:
            log.error("All recovery download attempts failed.")
            return False

    code = run_python(str(SCRIPT_DIR / "email_sender.py"))
    if code == 0:
        log.info("Recovery email sent successfully.")
        return True
    log.error(f"Recovery email failed (exit {code}).")
    return False


def send_watchdog_alert(message: str):
    # email_sender.py's --error path already sends to ERROR_EMAIL_RECIPIENTS
    # (hariit only) by itself — no --to needed here.
    code = run_python(str(SCRIPT_DIR / "email_sender.py"), "--error", message)
    if code != 0:
        log.error("Also failed to send the watchdog alert email itself.")


def main():
    if today_was_downloaded() and today_was_emailed():
        log.info("Today's report was downloaded and emailed successfully. No action needed.")
        return

    log.warning("Today's Trinity Sales Register report was NOT confirmed sent by the normal scheduled run.")
    send_watchdog_alert(
        "The watchdog did not find confirmation that today's Trinity Sales Register "
        "report was downloaded and emailed by the normal 7 AM scheduled run "
        "(run_task.bat). This usually means run_task.bat itself failed to run or "
        "crashed before reaching the point where it can alert on its own.\n\n"
        "Attempting automatic recovery now (downloading and emailing the report "
        "directly) — a follow-up email with the report will arrive shortly if this "
        "succeeds."
    )

    recovered = attempt_recovery()
    if recovered:
        log.info("Recovery succeeded — today's report has been sent.")
    else:
        log.error("Recovery failed — today's report was NOT sent. Manual attention needed.")
        send_watchdog_alert(
            "Automatic recovery ALSO failed after the watchdog detected today's Trinity "
            "Sales Register report was missing. The report has NOT been sent today. "
            "Please check the ERP connection / run_log.txt and download & send it "
            "manually."
        )


if __name__ == "__main__":
    main()
