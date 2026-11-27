import io
import sys
import time
from pathlib import Path

import openpyxl
import pytest
from fastapi.testclient import TestClient

SERVER_ROOT = Path(__file__).resolve().parent.parent
if str(SERVER_ROOT) not in sys.path:
    sys.path.insert(0, str(SERVER_ROOT))

from config.settings import REQUIRED_API_KEY
from app import create_app
from api.routes import _default_conversion_tabs
from generators.conversion_report_generator import (
    _clean_pct_val,
    build_agent_view,
    build_dc_view,
    extract_report_date_from_agent_view,
    generate_conversion_report,
)


def _make_workbook_bytes(sheets: dict) -> bytes:
    """Helper to build an in-memory .xlsx workbook from {sheet_name: [row_lists]}."""
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
    buf = io.BytesIO()
    wb.save(buf)
    wb.close()
    return buf.getvalue()


def _write_workbook(path: Path, sheets: dict) -> Path:
    path.write_bytes(_make_workbook_bytes(sheets))
    return path


def _build_full_synthetic_sheets():
    e2e_dc = [
        [
            "Source_DC", "L4D", "#Stores", "OFP", "Picked-up", "OFD", "Success_Del",
            "Succ_pickup%", "Succ_del%", "Productivity", "OFD_COD", "OFD_PP",
            "cod_del_update", "pp_del_update", "COD_Succ_del%", "PP_Succ_del%",
        ],
        ["car-khr", 10, 2, 50, 45, 40, 38, "90%", "95%", 12, 20, 20, 19, 19, "95%", "95%"],
        ["CAR_KHR", 5, 1, 25, 20, 20, 18, "80%", "90%", 10, 10, 10, 9, 9, "90%", "90%"],
        ["CARKHR", 5, 1, 25, 20, 20, 14, "80%", "70%", 10, 10, 10, 7, 7, "70%", "70%"],
        ["CAR", 10, 2, 100, 95, 100, 92, 0.95, 0.92, 15, 50, 50, 46, 46, 0.92, 0.92],
        ["BLR", 8, 3, 100, 88, 100, 86, 0.88, 0.86, 14, 50, 50, 43, 43, 0.86, 0.86],
        ["MUM", 6, 2, 100, 92, 100, 91, 0.92, 0.91, 13, 50, 50, 46, 45, 0.92, 0.90],
        ["HYD", 4, 1, 100, 80, 100, 80, 0.80, 0.80, 11, 50, 50, 40, 40, 0.80, 0.80],
        ["CHE", 3, 1, 100, 95, 100, 94, 0.95, 0.94, 12, 50, 50, 47, 47, 0.94, 0.94],
    ]

    agent_view = [
        ["Date", "CasperFHRID", "AgentName", "Source_DC", "OFP", "OFD", "del_update", "Picked-up", "DC"],
        ["26-11-2026", "FHR101", "Rahul", "CAR", 10, 10, 9, 9, "CAR"],
        ["26-11-2026", "FHR999", "Priya", "BLR", 10, 10, 8, 8, "BLR"],
        ["26-11-2026", "FHR888", "Amit", "MUM", 10, 10, 9, 9, "MUM"],
        ["26-11-2026", "FHR777", "Suresh West", "HYD", 10, 10, 7, 8, "HYD"],
        ["26-11-2026", "FHR666", "Kiran", "CHE", 10, 10, 9, 9, "CHE"],
    ]

    return {
        "E2E_DC": e2e_dc,
        "Agent_view": agent_view,
    }


def test_clean_pct_val_all_inputs():
    assert _clean_pct_val(None) == 0.0
    assert _clean_pct_val("") == 0.0
    assert _clean_pct_val(float("nan")) == 0.0
    assert _clean_pct_val("85.5%") == pytest.approx(0.855)
    assert _clean_pct_val("0.855") == pytest.approx(0.855)
    assert _clean_pct_val(0.92) == pytest.approx(0.92)
    assert _clean_pct_val(92) == pytest.approx(0.92)
    assert _clean_pct_val(92.5) == pytest.approx(0.925)
    assert _clean_pct_val("abc") == 0.0


def test_all_dc_passthrough_and_alias_consolidation(tmp_path):
    xlsx_path = _write_workbook(tmp_path / "input.xlsx", _build_full_synthetic_sheets())

    df_dc = build_dc_view(xlsx_path)
    dcs = set(df_dc["Source_DC"].tolist())
    assert dcs == {"CAR", "BLR", "MUM", "HYD", "CHE"}
    assert "Zone" not in df_dc.columns

    car_row = df_dc[df_dc["Source_DC"] == "CAR"].iloc[0]
    assert car_row["OFP"] == 200
    assert car_row["Picked-up"] == 180
    assert car_row["OFD"] == 180
    assert car_row["Success_Del"] == 162
    assert car_row["Succ_pickup%"] == pytest.approx(0.90)
    assert car_row["Succ_del%"] == pytest.approx(0.90)
    assert car_row["COD_Succ_del%"] == pytest.approx(round(81 / 90, 4))
    assert car_row["PP_Succ_del%"] == pytest.approx(round(81 / 90, 4))

    df_agent = build_agent_view(xlsx_path)
    assert set(df_agent["Source_DC"].tolist()) == {"CAR", "BLR", "MUM", "HYD", "CHE"}
    assert "Zone" not in df_agent.columns
    by_id = df_agent.set_index("CasperFHRID")
    assert by_id.loc["FHR101", "AgentName"] == "Rahul"
    assert by_id.loc["FHR101", "REV %"] == pytest.approx(0.9)
    assert by_id.loc["FHR101", "FWD %"] == pytest.approx(0.9)
    assert by_id.loc["FHR777", "REV %"] == pytest.approx(0.8)
    assert by_id.loc["FHR777", "FWD %"] == pytest.approx(0.7)


def test_report_date_tab_prefix_and_styling(tmp_path):
    xlsx_path = _write_workbook(tmp_path / "styled.xlsx", _build_full_synthetic_sheets())
    assert extract_report_date_from_agent_view(xlsx_path) == "26-11-2026"

    out_d1 = tmp_path / "out_d1.xlsx"
    generate_conversion_report(xlsx_path, out_d1, sub_type="d1")

    wb = openpyxl.load_workbook(out_d1)
    assert wb.sheetnames == ["D-1 DC_View", "D-1 Agent_View"]

    ws_dc = wb["D-1 DC_View"]
    assert ws_dc["A1"].fill.start_color.rgb.endswith("1F4E79")
    assert ws_dc.auto_filter.ref is not None

    headers = [c.value for c in ws_dc[1]]
    succ_pickup_col = headers.index("Succ_pickup%") + 1
    source_dc_col = headers.index("Source_DC") + 1

    fill_by_dc = {}
    for row in range(2, ws_dc.max_row + 1):
        dc_val = ws_dc.cell(row=row, column=source_dc_col).value
        cell = ws_dc.cell(row=row, column=succ_pickup_col)
        assert cell.number_format == "0.0%"
        fill_by_dc[dc_val] = cell.fill.start_color.rgb[-6:]

    assert fill_by_dc["HYD"] == "FFC7CE"  # 80% < 85% -> Red
    assert fill_by_dc["BLR"] == "FFEB9C"  # 88% in [85%, 90%) -> Yellow
    assert fill_by_dc["CAR"] == "C6EFCE"  # 90% >= 90% -> Green
    wb.close()


def test_default_conversion_tabs_helper():
    assert _default_conversion_tabs("d1") == ["D-1 DC_View", "D-1 Agent_View"]
    assert _default_conversion_tabs("sameday") == ["Sameday DC_View", "Sameday Agent_View"]
    assert _default_conversion_tabs("Custom") == ["Custom DC_View", "Custom Agent_View"]
    assert _default_conversion_tabs(None) == ["DC_View", "Agent_View"]


def test_api_end_to_end_integration():
    app = create_app()
    client = TestClient(app)

    r_root = client.get("/")
    assert r_root.status_code == 200
    assert r_root.json()["status"] == "ready"

    r_test = client.get("/test")
    assert r_test.status_code == 200
    assert "Conversion Report Server" in r_test.text

    r_unauth = client.get("/health")
    assert r_unauth.status_code == 403

    headers = {"X-API-KEY": REQUIRED_API_KEY}
    r_health = client.get("/health", headers=headers)
    assert r_health.status_code == 200
    assert r_health.json()["status"] == "ok"

    wb_bytes = _make_workbook_bytes(_build_full_synthetic_sheets())
    r_upload = client.post(
        "/reports/conversion?sub_type=d1",
        headers=headers,
        files={
            "file": (
                "E2E_Tracker.xlsx",
                wb_bytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        },
    )
    assert r_upload.status_code == 200
    job_id = r_upload.json()["job_id"]

    job_data = None
    for _ in range(40):
        r_status = client.get(f"/job/{job_id}", headers=headers)
        assert r_status.status_code == 200
        job_data = r_status.json()
        if job_data["status"] in ("done", "error"):
            break
        time.sleep(0.15)

    assert job_data is not None
    assert job_data["status"] == "done", f"Job failed: {job_data.get('error')}"
    assert job_data["tabs"] == ["D-1 DC_View", "D-1 Agent_View"]
    assert job_data["report_date"] == "26-11-2026"

    r_result = client.get(f"/job/{job_id}/result", headers=headers)
    assert r_result.status_code == 200
    assert (
        r_result.headers["content-type"]
        == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    out_wb = openpyxl.load_workbook(io.BytesIO(r_result.content))
    assert out_wb.sheetnames == ["D-1 DC_View", "D-1 Agent_View"]
    out_wb.close()

    r_logs = client.get(f"/job/{job_id}/logs", headers=headers)
    assert r_logs.status_code == 200
    assert len(r_logs.json()["logs"]) > 0

    r_clear = client.post("/clear-cache", headers=headers)
    assert r_clear.status_code == 200
    assert r_clear.json()["status"] == "ok"
