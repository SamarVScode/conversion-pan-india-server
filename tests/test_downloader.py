import io
import sys
from pathlib import Path

import openpyxl
import pytest
import requests
from fastapi import HTTPException

SERVER_ROOT = Path(__file__).resolve().parent.parent
if str(SERVER_ROOT) not in sys.path:
    sys.path.insert(0, str(SERVER_ROOT))

from core.downloader import download_from_url

SHEET_FILE_ID = "1uxo76wYKGFPqWeHsCTsIWyFVYvSczYJV"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


class _FakeResponse:
    """Minimal requests.Response stand-in for streaming download tests."""

    def __init__(self, status_code: int, body: bytes = b"", headers: dict | None = None):
        self.status_code = status_code
        self._body = body
        self.headers = headers or {}
        self.cookies = {}

    @property
    def text(self) -> str:
        return self._body.decode("utf-8", errors="ignore")

    def iter_content(self, chunk_size: int = 32768):
        for i in range(0, len(self._body), chunk_size):
            yield self._body[i : i + chunk_size]


def _make_xlsx_bytes() -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Agent_view"
    ws.append(["Date", "CasperFHRID", "AgentName"])
    ws.append(["26-11-2026", "FHR101", "Rahul"])
    buf = io.BytesIO()
    wb.save(buf)
    wb.close()
    return buf.getvalue()


def test_download_from_url_falls_back_to_sheets_export_on_403(tmp_path, monkeypatch):
    xlsx_bytes = _make_xlsx_bytes()
    calls = []

    def fake_session_get(self, url, **kwargs):
        calls.append((url, kwargs))
        if url.startswith("https://www.googleapis.com/drive/v3/files/") and "/export" in url:
            return _FakeResponse(
                403,
                body=b'{"error":{"errors":[{"reason":"exportSizeLimitExceeded"}]}}',
                headers={"Content-Type": "application/json"},
            )
        if url == f"https://docs.google.com/spreadsheets/d/{SHEET_FILE_ID}/export?format=xlsx":
            return _FakeResponse(
                200,
                body=xlsx_bytes,
                headers={"Content-Type": XLSX_MIME, "Content-Length": str(len(xlsx_bytes))},
            )
        raise AssertionError(f"Unexpected URL requested: {url}")

    monkeypatch.setattr(requests.Session, "get", fake_session_get)

    dest_path = tmp_path / "sheet.xlsx"
    source_url = f"https://docs.google.com/spreadsheets/d/{SHEET_FILE_ID}/edit?access_token=tok_abc_123"
    result_path = download_from_url(source_url, dest_path)

    assert result_path == dest_path
    assert dest_path.read_bytes() == xlsx_bytes

    assert len(calls) == 2
    assert calls[0][0] == (
        f"https://www.googleapis.com/drive/v3/files/{SHEET_FILE_ID}/export?mimeType={XLSX_MIME}"
    )
    assert calls[0][1]["headers"] == {"Authorization": "Bearer tok_abc_123"}
    assert calls[1][0] == f"https://docs.google.com/spreadsheets/d/{SHEET_FILE_ID}/export?format=xlsx"
    assert calls[1][1]["headers"] == {"Authorization": "Bearer tok_abc_123"}

    wb = openpyxl.load_workbook(dest_path, read_only=True)
    assert wb.sheetnames == ["Agent_view"]
    wb.close()


def test_download_from_url_sheets_export_without_token(tmp_path, monkeypatch):
    xlsx_bytes = _make_xlsx_bytes()
    calls = []

    def fake_session_get(self, url, **kwargs):
        calls.append((url, kwargs))
        if url == f"https://docs.google.com/spreadsheets/d/{SHEET_FILE_ID}/export?format=xlsx":
            return _FakeResponse(
                200,
                body=xlsx_bytes,
                headers={"Content-Type": XLSX_MIME, "Content-Length": str(len(xlsx_bytes))},
            )
        raise AssertionError(f"Unexpected URL requested: {url}")

    monkeypatch.setattr(requests.Session, "get", fake_session_get)

    dest_path = tmp_path / "public_sheet.xlsx"
    result_path = download_from_url(f"https://docs.google.com/spreadsheets/d/{SHEET_FILE_ID}/edit", dest_path)

    assert result_path == dest_path
    assert dest_path.read_bytes() == xlsx_bytes
    assert len(calls) == 1
    assert calls[0][0] == f"https://docs.google.com/spreadsheets/d/{SHEET_FILE_ID}/export?format=xlsx"
    assert calls[0][1]["headers"] == {}


def test_download_from_url_validates_sheets_export_and_falls_back_on_html(tmp_path, monkeypatch):
    xlsx_bytes = _make_xlsx_bytes()
    html_signin = b"<!doctype html><html><body>Sign in to accounts.google.com to continue</body></html>" + b" " * 100
    seen_urls = []

    def fake_session_get(self, url, **kwargs):
        seen_urls.append(url)
        if "/drive/v3/files/" in url and "/export" in url:
            return _FakeResponse(403, headers={"Content-Type": "application/json"})
        if url == f"https://docs.google.com/spreadsheets/d/{SHEET_FILE_ID}/export?format=xlsx":
            # Simulate HTML payload served without text/html Content-Type so _validate_downloaded_file catches it
            return _FakeResponse(200, body=html_signin, headers={"Content-Type": "application/octet-stream"})
        if url == f"https://www.googleapis.com/drive/v3/files/{SHEET_FILE_ID}?alt=media":
            return _FakeResponse(200, body=xlsx_bytes, headers={"Content-Type": XLSX_MIME})
        raise AssertionError(f"Unexpected URL requested: {url}")

    monkeypatch.setattr(requests.Session, "get", fake_session_get)

    dest_path = tmp_path / "fallback_binary.xlsx"
    result_path = download_from_url(
        f"https://drive.google.com/uc?export=download&id={SHEET_FILE_ID}&access_token=tok_xyz",
        dest_path,
    )

    assert result_path == dest_path
    assert dest_path.read_bytes() == xlsx_bytes
    assert seen_urls == [
        f"https://www.googleapis.com/drive/v3/files/{SHEET_FILE_ID}/export?mimeType={XLSX_MIME}",
        f"https://docs.google.com/spreadsheets/d/{SHEET_FILE_ID}/export?format=xlsx",
        f"https://www.googleapis.com/drive/v3/files/{SHEET_FILE_ID}?alt=media",
    ]


def test_download_from_url_cleans_up_invalid_file_when_all_methods_fail(tmp_path, monkeypatch):
    html_signin = b"<!doctype html><html><body>Sign in to accounts.google.com</body></html>" + b" " * 100

    def fake_session_get(self, url, **kwargs):
        return _FakeResponse(200, body=html_signin, headers={"Content-Type": "application/octet-stream"})

    monkeypatch.setattr(requests.Session, "get", fake_session_get)

    dest_path = tmp_path / "failed.xlsx"
    with pytest.raises(HTTPException) as exc_info:
        download_from_url(f"https://docs.google.com/spreadsheets/d/{SHEET_FILE_ID}/edit", dest_path)

    assert exc_info.value.status_code == 400
    assert not dest_path.exists()
