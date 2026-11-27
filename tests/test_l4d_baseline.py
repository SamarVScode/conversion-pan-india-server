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
                ["24-11-2026", "FHR400", "Vikram Old", "ABT", "ABT-OLD"],
                ["24-11-2026", "N/A", "Invalid NA", "BLR", "BLR"],
                ["24-11-2026", "0", "Invalid Zero", "BLR", "BLR"],
                ["24-11-2026", 0, "Invalid Zero Num", "BLR", "BLR"],
            ],
        },
    )

    # Day 2: Sameday Agent_View (middle) - FHR200 lowercase/padded, FHR300 only DC col with car-khr
    day2 = _write_xlsx(
        tmp_path / "day2.xlsx",
        {
            "Sameday Agent_View": [
                ["Date", "CasperID", "AgentName", "DC"],
                ["2026-11-25", " fhr200 ", "Priya", "HYD"],
                ["2026-11-25", "FHR300", "Amit", "car-khr"],
                ["2026-11-25", "n/a", "Invalid Lower NA", "CHE"],
            ],
        },
    )

    # Day 3: CSV (newest) - 1001 moved to car-khr -> source_dc CAR, mini_dc CAR-KHR; fhr400 -> ABT / ABT-ABR
    day3 = _write_csv(
        tmp_path / "day3.csv",
        [
            ["Date", "CasperFHRID", "AgentName", "Source_DC", "DC"],
            ["26/11/2026", "1001", "Rahul New", "car-khr", "car-khr"],
            ["26/11/2026", "FHR200", "Priya", "MUM", "MUM"],
            ["26/11/2026", "fhr400", "Vikram New", "ABT", "abt-abr"],
        ],
    )

    result = aggregate_l4d_baseline([day1, day2, day3], delete_after_read=True)

    # Temp files should be deleted immediately after processing
    assert not day1.exists()
    assert not day2.exists()
    assert not day3.exists()

    assert result["success"] is True
    assert result["total_agents"] == 4
    assert result["dates_processed"] == ["2026-11-26", "2026-11-25", "2026-11-24"]

    by_id = {row[0]: row for row in result["rows"]}
    assert "N/A" not in by_id
    assert "0" not in by_id

    # 1001: active 26 & 24, latest source_dc normalized from car-khr -> CAR, mini_dc CAR-KHR
    assert by_id["1001"] == [
        "1001",
        "Rahul New",
        "CAR",
        "CAR-KHR",
        "2026-11-26",
        "2026-11-24",
        "2026-11-26,2026-11-24",
    ]

    # FHR200: active 26, 25, 24 (normalized from " fhr200 "); latest source_dc/mini_dc MUM
    assert by_id["FHR200"] == [
        "FHR200",
        "Priya",
        "MUM",
        "MUM",
        "2026-11-26",
        "2026-11-25",
        "2026-11-26,2026-11-25,2026-11-24",
    ]

    # FHR400: active 26 & 24, latest mini_dc ABT-ABR alongside source_dc ABT
    assert by_id["FHR400"] == [
        "FHR400",
        "Vikram New",
        "ABT",
        "ABT-ABR",
        "2026-11-26",
        "2026-11-24",
        "2026-11-26,2026-11-24",
    ]

    # FHR300: active only on 25 with only DC column = car-khr -> source_dc CAR, mini_dc CAR-KHR
    assert by_id["FHR300"] == [
        "FHR300",
        "Amit",
        "CAR",
        "CAR-KHR",
        "2026-11-25",
        "",
        "2026-11-25",
    ]


def test_mini_dc_header_variants_and_latest_date_tracking(tmp_path):
    # Pass newest date first and oldest second to verify latest_date wins regardless of file order
    newer = _write_xlsx(
        tmp_path / "newer.xlsx",
        {
            "Agent_view": [
                ["Date", "CasperFHRID", "AgentName", "Source_DC", "Mini DC"],
                ["26-11-2026", "FHR501", "Suresh Latest", "ABT", "ABT-ABR"],
                ["26-11-2026", "FHR502", "Meena", "car-khr", None],
            ]
        },
    )
    older = _write_xlsx(
        tmp_path / "older.xlsx",
        {
            "Agent_view": [
                ["Date", "CasperFHRID", "AgentName", "Source_DC"],
                ["24-11-2026", "fhr501", "Suresh Old", "BLR"],
                ["24-11-2026", "FHR503", "Rohan", "car-khr"],
            ]
        },
    )
    sub_dc_file = _write_xlsx(
        tmp_path / "subdc.xlsx",
        {
            "Agent_view": [
                ["Date", "CasperFHRID", "AgentName", "Sub DC", "Source_DC"],
                ["25-11-2026", "FHR504", "Divya", "ABT-KRN", "ABT"],
            ]
        },
    )

    result = aggregate_l4d_baseline([newer, older, sub_dc_file], delete_after_read=False)
    by_id = {row[0]: row for row in result["rows"]}

    # FHR501 keeps newer date metadata (ABT / ABT-ABR), not overwritten by older file (BLR)
    assert by_id["FHR501"] == [
        "FHR501",
        "Suresh Latest",
        "ABT",
        "ABT-ABR",
        "2026-11-26",
        "2026-11-24",
        "2026-11-26,2026-11-24",
    ]
    # FHR502 blank Mini DC falls back to raw Source_DC uppercased (CAR-KHR)
    assert by_id["FHR502"] == [
        "FHR502",
        "Meena",
        "CAR",
        "CAR-KHR",
        "2026-11-26",
        "",
        "2026-11-26",
    ]
    # FHR503 has only Source_DC=car-khr -> mini_dc_idx is distinct (-1) and falls back to CAR-KHR
    assert by_id["FHR503"] == [
        "FHR503",
        "Rohan",
        "CAR",
        "CAR-KHR",
        "2026-11-24",
        "",
        "2026-11-24",
    ]
    # FHR504 resolves Sub DC column even when placed before Source_DC
    assert by_id["FHR504"] == [
        "FHR504",
        "Divya",
        "ABT",
        "ABT-KRN",
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
    assert len(first_row) == 7
    assert first_row[4] == "2026-11-26"
    assert first_row[5] == ""
    assert first_row[6] == "2026-11-26"

    by_id = {row[0]: row for row in result["rows"]}
    assert by_id["15227"][2] == "ABT"
    assert by_id["15227"][3] == "ABT-ABR"
    assert any(row[2] == "ABT" and row[3] == "ABT-ABR" for row in result["rows"])


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
                ["Date", "CasperFHRID", "AgentName", "Source_DC", "DC"],
                ["26-11-2026", "FHR101", "Rahul", "car-khr", "CAR-KHR"],
                ["26-11-2026", "FHR102", "Kiran", "ABT", "ABT-ABR"],
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
    assert by_id["FHR101"] == ["FHR101", "Rahul", "CAR", "CAR-KHR", "2026-11-26", "2026-11-25", "2026-11-26,2026-11-25"]
    assert by_id["FHR102"] == ["FHR102", "Kiran", "ABT", "ABT-ABR", "2026-11-26", "", "2026-11-26"]

    # Verify access_token was passed to downloader
    assert len(seen_urls) == 2
    assert all("access_token=tok_test_123" in u for u in seen_urls)

    # Cleanup test artifacts from cache
    client.post("/clear-cache", headers=headers)


def test_l4d_baseline_single_flight_coalescing(tmp_path, monkeypatch):
    """Multiple concurrent / repeated requests for the same file_ids coalesce onto the first job_id."""
    import threading
    import core.jobs as jobs_mod

    f1 = _write_xlsx(
        tmp_path / "f1.xlsx",
        {
            "Agent_view": [
                ["Date", "CasperFHRID", "AgentName", "Source_DC"],
                ["26-11-2026", "FHR201", "Anita", "BLR"],
            ]
        },
    )
    id1 = "3AbCdEfGhIjKlMnOpQrStUvWxYz012345"
    id2 = "4AbCdEfGhIjKlMnOpQrStUvWxYz012345"
    download_calls = {"count": 0}
    gate = threading.Event()

    def slow_fake_download(url: str, dest_path: Path, job_id=None):
        download_calls["count"] += 1
        gate.wait(timeout=2.0)
        dest_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(f1, dest_path)
        return dest_path

    monkeypatch.setattr(jobs_mod, "download_from_url", slow_fake_download)

    app = create_app()
    client = TestClient(app)
    headers = {"X-API-KEY": REQUIRED_API_KEY}
    client.post("/clear-cache", headers=headers)

    try:
        # Request 1 starts the single server job
        r1 = client.post("/l4d-historical-baseline", headers=headers, json={"file_ids": [id1, id2]})
        assert r1.status_code == 200
        j1 = r1.json()
        assert j1.get("deduplicated") is None

        # Request 2 (same IDs in reversed order) coalesces onto j1 instead of spawning a 2nd job
        r2 = client.post("/l4d-historical-baseline", headers=headers, json={"file_ids": [id2, id1]})
        assert r2.status_code == 200
        j2 = r2.json()
        assert j2["job_id"] == j1["job_id"]
        assert j2.get("deduplicated") is True

        gate.set()
        for _ in range(40):
            st = client.get(f"/job/{j1['job_id']}", headers=headers).json()
            if st["status"] in ("done", "error"):
                break
            time.sleep(0.1)

        assert st["status"] == "done"
        # Request 3 after completion returns cached result with same job_id without re-downloading
        r3 = client.post("/l4d-historical-baseline", headers=headers, json={"file_ids": [id1, id2]})
        assert r3.status_code == 200
        assert r3.json()["job_id"] == j1["job_id"]
        assert r3.json()["deduplicated"] is True
        # Only the 2 files from Request 1 were downloaded (not 6)
        assert download_calls["count"] == 2
    finally:
        gate.set()
        client.post("/clear-cache", headers=headers)
