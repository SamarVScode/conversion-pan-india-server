import csv
import shutil
import sys
import time
from pathlib import Path

import openpyxl
import pytest
from fastapi.testclient import TestClient

SERVER_ROOT = Path(__file__).resolve().parent.parent
if str(SERVER_ROOT) not in sys.path:
    sys.path.insert(0, str(SERVER_ROOT))

from app import create_app
from config.settings import REQUIRED_API_KEY
from generators.l4d_baseline_generator import aggregate_l4d_baseline, normalize_iso_date


def _write_xlsx(path: Path, sheets: dict) -> Path:
    wb = openpyxl.Workbook()
    first = True
    for name, rows in sheets.items():
        if first:
            ws = wb.active
            ws.title = name
            first = False
        else:
            ws = wb.create_sheet(title=name)
        for r in rows:
            ws.append(r)
    wb.save(path)
    wb.close()
    return path


def _write_csv(path: Path, rows: list) -> Path:
    with open(path, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerows(rows)
    return path


def test_normalize_iso_date_variants():
    assert normalize_iso_date("26-11-2026") == "2026-11-26"
    assert normalize_iso_date("2026-11-25") == "2026-11-25"
    assert normalize_iso_date("24/11/2026") == "2026-11-24"
    assert normalize_iso_date("26-Nov-2026") == "2026-11-26"
    assert normalize_iso_date(None) is None
    assert normalize_iso_date("") is None


def test_multi_day_aggregation_and_cleanup(tmp_path):
    # Day 1: D-1 Agent_View (oldest)
    day1 = _write_xlsx(
        tmp_path / "day1.xlsx",
        {
            "D-1 DC_View": [["Source_DC", "OFP"], ["BLR", 10]],
            "D-1 Agent_View": [
                ["Date", "CasperFHRID", "AgentName", "Source_DC", "DC"],
                ["24-11-2026", 1001.0, "Rahul Old", "BLR", "BLR"],
                ["24-11-2026", "FHR200", "Priya", "MUM", "MUM"],
            ],
        },
    )

    # Day 2: Sameday Agent_View (middle) - FHR1001 not present, FHR200 present, FHR300 only DC col
    day2 = _write_xlsx(
        tmp_path / "day2.xlsx",
        {
            "Sameday Agent_View": [
                ["Date", "CasperID", "AgentName", "DC"],
                ["2026-11-25", "FHR200", "Priya", "HYD"],
                ["2026-11-25", "FHR300", "Amit", "CHE"],
            ],
        },
    )

    # Day 3: CSV (newest) - 1001 moved to car-khr -> normalized CAR
    day3 = _write_csv(
        tmp_path / "day3.csv",
        [
            ["Date", "CasperFHRID", "AgentName", "Source_DC", "DC"],
            ["26/11/2026", "1001", "Rahul New", "car-khr", "BLR"],
            ["26/11/2026", "FHR200", "Priya", "MUM", "MUM"],
        ],
    )

    result = aggregate_l4d_baseline([day1, day2, day3], delete_after_read=True)

    # Temp files should be deleted immediately after processing
    assert not day1.exists()
    assert not day2.exists()
    assert not day3.exists()

    assert result["success"] is True
    assert result["total_agents"] == 3
    assert result["dates_processed"] == ["2026-11-26", "2026-11-25", "2026-11-24"]

    by_id = {row[0]: row for row in result["rows"]}

    # 1001: active 26 & 24, latest Source_DC normalized from car-khr -> CAR, latest name Rahul New
    assert by_id["1001"] == [
        "1001",
        "Rahul New",
        "CAR",
        "2026-11-26",
        "2026-11-24",
        "2026-11-26,2026-11-24",
    ]

    # FHR200: active 26, 25, 24; latest Source_DC MUM
    assert by_id["FHR200"] == [
        "FHR200",
        "Priya",
        "MUM",
        "2026-11-26",
        "2026-11-25",
        "2026-11-26,2026-11-25,2026-11-24",
    ]

    # FHR300: active only on 25, prevActiveDate is ""
    assert by_id["FHR300"] == [
        "FHR300",
        "Amit",
        "CHE",
        "2026-11-25",
        "",
        "2026-11-25",
    ]


def test_real_d1_report_if_present():
    sample_path = Path(r"C:\Users\User\Desktop\card\Conversion_Summary_D-1_Report_26-11-2026.xlsx")
    if not sample_path.exists():
        pytest.skip("Sample D-1 report file not present on disk")

    result = aggregate_l4d_baseline([sample_path], delete_after_read=False)
    assert sample_path.exists()  # Must not delete fixture
    assert result["success"] is True
    assert result["total_agents"] > 0
    assert result["dates_processed"] == ["2026-11-26"]
    first_row = result["rows"][0]
    assert len(first_row) == 6
    assert first_row[3] == "2026-11-26"
    assert first_row[4] == ""
    assert first_row[5] == "2026-11-26"


def test_l4d_baseline_api_auth_and_end_to_end(tmp_path, monkeypatch):
    f1 = _write_xlsx(
        tmp_path / "f1.xlsx",
        {
            "Agent_view": [
                ["Date", "CasperFHRID", "AgentName", "Source_DC"],
                ["25-11-2026", "FHR101", "Rahul", "BLR"],
            ]
        },
    )
    f2 = _write_xlsx(
        tmp_path / "f2.xlsx",
        {
            "D-1 Agent_View": [
                ["Date", "CasperFHRID", "AgentName", "Source_DC"],
                ["26-11-2026", "FHR101", "Rahul", "car_khr"],
                ["26-11-2026", "FHR102", "Kiran", "HYD"],
            ]
        },
    )

    id1 = "1AbCdEfGhIjKlMnOpQrStUvWxYz012345"
    id2 = "2AbCdEfGhIjKlMnOpQrStUvWxYz012345"
    source_map = {id1: f1, id2: f2}
    seen_urls = []

    def fake_download(url: str, dest_path: Path, job_id=None):
        seen_urls.append(url)
        for fid, src in source_map.items():
            if fid in url:
                dest_path.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(src, dest_path)
                return dest_path
        raise AssertionError(f"Unexpected download URL: {url}")

    monkeypatch.setattr("core.jobs.download_from_url", fake_download)

    app = create_app()
    client = TestClient(app)

    # Auth enforcement (403 without key)
    unauth_resp = client.post(
        "/l4d-historical-baseline",
        json={"file_ids": [id1]},
    )
    assert unauth_resp.status_code == 403

    unauth_res = client.get("/l4d-job/nonexistent/result")
    assert unauth_res.status_code == 403

    headers = {"X-API-KEY": REQUIRED_API_KEY}

    # Start L4D baseline job
    post_resp = client.post(
        "/l4d-historical-baseline",
        headers=headers,
        json={
            "file_ids": [id1, f"https://drive.google.com/file/d/{id2}/view"],
            "access_token": "tok_test_123",
        },
    )
    assert post_resp.status_code == 200
    body = post_resp.json()
    job_id = body["job_id"]
    assert body["status"] == "processing"
    assert body["file_count"] == 2

    # Poll /job/{job_id}
    final_status = None
    for _ in range(40):
        status_resp = client.get(f"/job/{job_id}", headers=headers)
        assert status_resp.status_code == 200
        sdata = status_resp.json()
        assert sdata["report_type"] == "l4d_baseline"
        if sdata["status"] in ("done", "error"):
            final_status = sdata
            break
        time.sleep(0.15)

    assert final_status is not None
    assert final_status["status"] == "done", f"Job failed: {final_status}"

    # Verify /job/{job_id}/logs works
    logs_resp = client.get(f"/job/{job_id}/logs", headers=headers)
    assert logs_resp.status_code == 200
    assert isinstance(logs_resp.json()["logs"], list)
    assert len(logs_resp.json()["logs"]) > 0

    # Fetch /l4d-job/{job_id}/result
    res_resp = client.get(f"/l4d-job/{job_id}/result", headers=headers)
    assert res_resp.status_code == 200
    res_json = res_resp.json()
    assert res_json["success"] is True
    assert res_json["total_agents"] == 2
    assert res_json["dates_processed"] == ["2026-11-26", "2026-11-25"]

    by_id = {r[0]: r for r in res_json["rows"]}
    assert by_id["FHR101"] == ["FHR101", "Rahul", "CAR", "2026-11-26", "2026-11-25", "2026-11-26,2026-11-25"]
    assert by_id["FHR102"] == ["FHR102", "Kiran", "HYD", "2026-11-26", "", "2026-11-26"]

    # Verify access_token was passed to downloader
    assert len(seen_urls) == 2
    assert all("access_token=tok_test_123" in u for u in seen_urls)

    # Cleanup test artifacts from cache
    client.post("/clear-cache", headers=headers)
