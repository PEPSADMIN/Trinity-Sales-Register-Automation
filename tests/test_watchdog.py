import sys
from datetime import date, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import watchdog as wd  # noqa: E402


def test_today_was_downloaded_true_when_todays_file_present(tmp_path, monkeypatch):
    monkeypatch.setattr(wd, "DOWNLOAD_DIR", tmp_path)
    today_name = f"Trinity Sales Register -{date.today().strftime('%d.%m.%Y')}.xlsx"
    (tmp_path / today_name).write_bytes(b"x")
    assert wd.today_was_downloaded() is True


def test_today_was_downloaded_false_when_only_old_file_present(tmp_path, monkeypatch):
    monkeypatch.setattr(wd, "DOWNLOAD_DIR", tmp_path)
    yesterday = date.today() - timedelta(days=1)
    old_name = f"Trinity Sales Register -{yesterday.strftime('%d.%m.%Y')}.xlsx"
    (tmp_path / old_name).write_bytes(b"x")
    assert wd.today_was_downloaded() is False


def test_today_was_emailed_true_when_state_file_has_todays_date(tmp_path, monkeypatch):
    state_file = tmp_path / ".last_sent"
    state_file.write_text(date.today().isoformat(), encoding="utf-8")
    monkeypatch.setattr(wd, "LAST_SENT_FILE", state_file)
    assert wd.today_was_emailed() is True


def test_today_was_emailed_false_when_state_file_is_stale(tmp_path, monkeypatch):
    state_file = tmp_path / ".last_sent"
    state_file.write_text((date.today() - timedelta(days=1)).isoformat(), encoding="utf-8")
    monkeypatch.setattr(wd, "LAST_SENT_FILE", state_file)
    assert wd.today_was_emailed() is False


def test_today_was_emailed_false_when_state_file_missing(tmp_path, monkeypatch):
    monkeypatch.setattr(wd, "LAST_SENT_FILE", tmp_path / ".last_sent")
    assert wd.today_was_emailed() is False
