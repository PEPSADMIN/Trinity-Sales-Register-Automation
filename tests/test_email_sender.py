import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import email_sender as es  # noqa: E402


def test_file_date_parses_valid_name():
    assert es._file_date("Trinity Sales Register -07.09.2026.xlsx") == date(2026, 9, 7)


def test_file_date_unrecognized_name_returns_min():
    assert es._file_date("something-else.xlsx") == date.min


def test_latest_sales_file_picks_true_newest_not_alphabetical(tmp_path, monkeypatch):
    # Regression test for the Sep-06/07-2026 incident: alphabetical
    # sorting picked "...-29.08.2026.xlsx" over "...-07.09.2026.xlsx"
    # because "29" > "07" as a string, even though 07-Sep is newer.
    monkeypatch.setattr(es, "DOWNLOAD_DIR", tmp_path)
    for name in (
        "Trinity Sales Register -29.08.2026.xlsx",
        "Trinity Sales Register -02.09.2026.xlsx",
        "Trinity Sales Register -07.09.2026.xlsx",
    ):
        (tmp_path / name).write_bytes(b"x")

    picked = es.latest_sales_file()

    assert picked.name == "Trinity Sales Register -07.09.2026.xlsx"


def test_latest_sales_file_raises_when_none_found(tmp_path, monkeypatch):
    monkeypatch.setattr(es, "DOWNLOAD_DIR", tmp_path)
    try:
        es.latest_sales_file()
        assert False, "expected FileNotFoundError"
    except FileNotFoundError:
        pass
