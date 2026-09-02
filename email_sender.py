"""
Trinity Sales Register — Email sender.

Emails the latest downloaded Sales Register Mattresses export as an
attachment, WITHOUT any modification (no editing, no breakup, no
WhatsApp). Mirrors the OEE project's SMTP approach (config.py style).

Meant to run automatically right after automate_report.py finishes a
successful download (see run_task.bat).
"""

import os
import glob
import logging
import smtplib
import argparse
from pathlib import Path
from datetime import date, datetime

from dotenv import load_dotenv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

load_dotenv(dotenv_path=Path(__file__).parent / "config.env")

# CLI overrides (used for safe test runs, e.g. --to hariit@pepsindia.com --cc "")
parser = argparse.ArgumentParser()
parser.add_argument("--to", default=None, help="Override TO recipients (comma-separated)")
parser.add_argument("--cc", default=None, help="Override CC recipients (comma-separated)")
parser.add_argument("--error", default=None, help="Send an error-alert email (no attachment) with this message")
parser.add_argument("--error-file", default=None, help="Read the issue detail from this file and append it to the error message")
args = parser.parse_args()

logging.basicConfig(level=logging.INFO, format="%(asctime)s [EMAIL] %(levelname)s %(message)s")
log = logging.getLogger("trinity_email")

DOWNLOAD_DIR   = Path(os.getenv("DOWNLOAD_DIR", r"C:\Users\ADMIN\Downloads"))
SMTP_SERVER    = os.getenv("SMTP_SERVER", "zimsmtp.logix.in")
SMTP_PORT      = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER      = os.getenv("SMTP_USER", "")
SMTP_PASSWORD  = os.getenv("SMTP_PASSWORD", "")
FROM_NAME      = os.getenv("EMAIL_FROM_NAME", "PEPS India")
RECIPIENTS     = [n.strip() for n in (args.to if args.to is not None else os.getenv("EMAIL_RECIPIENTS", "")).split(",") if n.strip()]
CC             = [n.strip() for n in (args.cc if args.cc is not None else os.getenv("EMAIL_CC", "")).split(",") if n.strip()]


def latest_sales_file() -> Path:
    matches = list(glob.glob(str(DOWNLOAD_DIR / "Trinity Sales Register -*.xlsx")))
    if not matches:
        raise FileNotFoundError(f"No 'Trinity Sales Register -*.xlsx' found in {DOWNLOAD_DIR}")

    # Pick the TRUE newest by parsing the DD.MM.YYYY date from the
    # filename. String-sorting the names is wrong across month
    # boundaries (e.g. "02.09.2026" < "29.08.2026" alphabetically).
    def _key(p: str):
        base = Path(p).stem  # 'Trinity Sales Register -29.08.2026'
        part = base.split(" -")[-1] if " -" in base else ""
        try:
            return datetime.strptime(part, "%d.%m.%Y")
        except ValueError:
            return datetime.min

    return Path(max(matches, key=_key))


def _attach_file(msg: MIMEMultipart, path: Path):
    with open(path, "rb") as fh:
        part = MIMEBase(
            "application",
            "vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        part.set_payload(fh.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename={path.name}")
    msg.attach(part)


def send_latest():
    if not SMTP_USER or not SMTP_PASSWORD or SMTP_PASSWORD == "FILL_PASSWORD_LATER":
        raise RuntimeError("SMTP_USER / SMTP_PASSWORD not configured in config.env")
    if not RECIPIENTS:
        raise RuntimeError("No EMAIL_RECIPIENTS configured in config.env")

    file_path = latest_sales_file()
    today = date.today()
    subject = f"Trinity Sales Register - {today.strftime('%d.%m.%Y')}"
    body = (
        f"Please find attached the Trinity Sales Register report "
        f"for {today.strftime('%d-%b-%Y')}.\n\n"
        f"File: {file_path.name}\n\n"
        f"This is an automated email — please do not reply."
    )

    msg = MIMEMultipart()
    msg["Subject"] = subject
    msg["From"] = f"{FROM_NAME} <{SMTP_USER}>"
    msg["To"] = ", ".join(RECIPIENTS)
    if CC:
        msg["Cc"] = ", ".join(CC)
    msg.attach(MIMEText(body, "plain", "utf-8"))
    _attach_file(msg, file_path)

    all_addrs = RECIPIENTS + CC
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=600) as srv:
        srv.ehlo()
        srv.starttls()
        srv.login(SMTP_USER, SMTP_PASSWORD)
        srv.sendmail(SMTP_USER, all_addrs, msg.as_string())

    log.info(f"Email sent with attachment {file_path.name} → To: {RECIPIENTS}  CC: {CC}")


def send_error(message: str, recipients: list):
    """Send a plain-text error-alert email (no attachment) — used when the
    download retries are exhausted. Goes only to the given recipients
    (typically just hariit@pepsindia.com), never the normal CC list."""
    if not SMTP_USER or not SMTP_PASSWORD or SMTP_PASSWORD == "FILL_PASSWORD_LATER":
        raise RuntimeError("SMTP_USER / SMTP_PASSWORD not configured in config.env")
    if not recipients:
        raise RuntimeError("No error-email recipients configured.")

    subject = f"{FROM_NAME} — Trinity Sales Register: DOWNLOAD FAILED"
    body = (
        f"The automated Trinity Sales Register download FAILED.\n\n"
        f"{message}\n\n"
        f"Generated: {datetime.now().strftime('%d-%b-%Y %I:%M %p')}\n"
        f"This is an automated alert — please do not reply."
    )
    msg = MIMEMultipart()
    msg["Subject"] = subject
    msg["From"] = f"{FROM_NAME} <{SMTP_USER}>"
    msg["To"] = ", ".join(recipients)
    msg.attach(MIMEText(body, "plain", "utf-8"))

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=600) as srv:
        srv.ehlo()
        srv.starttls()
        srv.login(SMTP_USER, SMTP_PASSWORD)
        srv.sendmail(SMTP_USER, recipients, msg.as_string())

    log.info(f"Error alert email sent → {recipients}")


if __name__ == "__main__":
    if args.error is not None:
        err_to = [n.strip() for n in
                  os.getenv("ERROR_EMAIL_RECIPIENTS", "hariit@pepsindia.com").split(",")
                  if n.strip()]
        message = args.error
        if args.error_file:
            try:
                with open(args.error_file, encoding="utf-8") as _fh:
                    detail = _fh.read().strip()
                if detail:
                    message = f"{message}\n\nIssue: {detail}"
            except Exception as _e:
                message = f"{message}\n\n(Also, could not read error detail file: {_e})"
        send_error(message, err_to)
    else:
        send_latest()
