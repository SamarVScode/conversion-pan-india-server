#!/usr/bin/env python3
"""
L4D Historical Baseline Generator
=================================
Streams only the 'Agent_view' (or 'D-1 Agent_View' / 'Sameday Agent_View')
sheet across a list of downloaded spreadsheets (.xlsx / .xlsb / .csv) in
O(1) memory, aggregates per-agent active dates, and produces compact baseline
rows for 10-15 day historical tracking.
"""

import gc
import logging
import math
import re
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Optional, Set, Union

from config.dc_config import normalize_dc_code
from core.stream_engine import ColumnFinder, get_sheet_names, open_stream_reader

log = logging.getLogger("conversion_report_server.l4d_baseline")

_DATE_FORMATS = (
    "%Y-%m-%d",
    "%d-%m-%Y",
    "%d/%m/%Y",
    "%m/%d/%Y",
    "%Y/%m/%d",
    "%d-%b-%Y",
    "%d-%B-%Y",
    "%Y-%m-%d %H:%M:%S",
    "%d-%m-%Y %H:%M:%S",
    "%d/%m/%Y %H:%M:%S",
)


def normalize_iso_date(val: Any) -> Optional[str]:
    """Normalize a date value (datetime, date, Excel serial, or string) to YYYY-MM-DD."""
    if val is None or isinstance(val, bool):
        return None
    if isinstance(val, datetime):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, date):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, (int, float)):
        if math.isnan(val) or math.isinf(val):
            return None
        if 20000 <= val <= 100000:
            return (datetime(1899, 12, 30) + timedelta(days=int(val))).strftime("%Y-%m-%d")
        return None

    val_str = str(val).strip()
    if not val_str or val_str.lower() in ("none", "nan", "nat", "null"):
        return None

    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(val_str, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass

    match = re.search(r"(\d{4}-\d{2}-\d{2})|(\d{2}-\d{2}-\d{4})|(\d{2}/\d{2}/\d{4})", val_str)
    if match:
        matched = match.group(0)
        for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y"):
            try:
                return datetime.strptime(matched, fmt).strftime("%Y-%m-%d")
            except ValueError:
                pass
    return None


def _normalize_casper_id(val: Any) -> str:
    """Normalize CasperFHRID / CasperID to a clean string without trailing '.0'."""
    if val is None or isinstance(val, bool):
        return ""
    if isinstance(val, int):
        return str(val)
    if isinstance(val, float):
        if math.isnan(val) or math.isinf(val):
            return ""
        return str(int(val)) if val.is_integer() else str(val)

    s = str(val).strip()
    if not s or s.lower() in ("none", "nan", "nat", "null"):
        return ""
    if re.fullmatch(r"\d+\.0+", s):
        return s.split(".", 1)[0]
    return s


def _find_agent_view_sheet(file_path: Path) -> Optional[str]:
    """Find the first sheet whose name contains 'agent_view' or 'agent view' (case-insensitive)."""
    if file_path.suffix.lower() in (".csv", ".tsv", ".txt"):
        return None
    for name in get_sheet_names(file_path):
        low = name.lower()
        if "agent_view" in low or "agent view" in low:
            return name
    return None


def aggregate_l4d_baseline(
    file_paths: Iterable[Union[str, Path]],
    delete_after_read: bool = True,
    progress_cb: Optional[Callable[[int, int, Path], None]] = None,
) -> Dict[str, Any]:
    """
    Stream Agent_view sheets across multiple spreadsheet files and aggregate
    per-agent active dates in low memory.

    Returns:
        {
            "success": True,
            "total_agents": int,
            "dates_processed": List[str],  # sorted YYYY-MM-DD descending
            "rows": List[List[str]],       # [casperId, agentName, sourceDC, lastActiveDate, prevActiveDate, activeDatesCSV]
        }
    """
    paths = [Path(p) for p in file_paths]
    total_files = len(paths)
    dates_processed: Set[str] = set()
    agents: Dict[str, Dict[str, Any]] = {}

    for idx, path in enumerate(paths, 1):
        try:
            if not path.exists():
                log.warning(f"Skipping missing file: {path}")
                continue

            is_text_table = path.suffix.lower() in (".csv", ".tsv", ".txt")
            target_sheet = _find_agent_view_sheet(path)
            if not is_text_table and target_sheet is None:
                log.warning(f"No Agent_view sheet found in {path.name}; skipping.")
                continue

            with open_stream_reader(path, sheet_name=target_sheet) as (headers, row_iter):
                if not headers:
                    log.warning(f"Empty headers in {path.name}; skipping.")
                    continue

                finder = ColumnFinder(headers)
                date_idx = finder.find(
                    ["Date", "Report_Date", "Report Date", "Txn_Date"],
                    default=-1,
                )
                casper_idx = finder.find(
                    ["CasperFHRID", "CasperID", "Casper_ID", "Casper_FHR_ID", "FHRID", "Agent_ID"],
                    default=-1,
                )
                name_idx = finder.find(
                    ["AgentName", "Agent_Name", "Agent Name", "Name"],
                    default=-1,
                )
                dc_idx = finder.find(
                    ["Source_DC", "Source DC", "SourceDC"],
                    fallback_keywords=["sourcedc"],
                    default=-1,
                )
                if dc_idx < 0:
                    dc_idx = finder.find(
                        ["DC", "DC_Code", "DC Code"],
                        fallback_keywords=["dc"],
                        default=-1,
                    )

                if date_idx < 0 or casper_idx < 0:
                    log.warning(
                        f"Required Date/CasperFHRID columns missing in {path.name} "
                        f"(headers={headers[:10]}); skipping."
                    )
                    continue

                for row in row_iter:
                    if not row:
                        continue
                    row_len = len(row)
                    raw_casper = row[casper_idx] if casper_idx < row_len else None
                    casper_id = _normalize_casper_id(raw_casper)
                    if not casper_id:
                        continue

                    raw_date = row[date_idx] if date_idx < row_len else None
                    norm_date = normalize_iso_date(raw_date)
                    if not norm_date:
                        continue

                    raw_name = row[name_idx] if (0 <= name_idx < row_len) else None
                    agent_name = str(raw_name).strip() if raw_name is not None else ""
                    if agent_name.lower() in ("none", "nan", "nat"):
                        agent_name = ""

                    raw_dc = row[dc_idx] if (0 <= dc_idx < row_len) else None
                    source_dc = normalize_dc_code(raw_dc)

                    dates_processed.add(norm_date)

                    entry = agents.get(casper_id)
                    if entry is None:
                        agents[casper_id] = {
                            "dates": {norm_date},
                            "latest_date": norm_date,
                            "agent_name": agent_name,
                            "source_dc": source_dc,
                        }
                    else:
                        entry["dates"].add(norm_date)
                        if norm_date > entry["latest_date"]:
                            entry["latest_date"] = norm_date
                            entry["agent_name"] = agent_name
                            entry["source_dc"] = source_dc
                        elif norm_date == entry["latest_date"]:
                            if agent_name:
                                entry["agent_name"] = agent_name
                            if source_dc:
                                entry["source_dc"] = source_dc

            if progress_cb:
                progress_cb(idx, total_files, path)
        finally:
            if delete_after_read and path.exists():
                try:
                    path.unlink(missing_ok=True)
                except Exception as del_err:
                    log.warning(f"Could not delete temp file {path.name}: {del_err}")
            gc.collect()

    rows: List[List[str]] = []
    for casper_id, entry in agents.items():
        sorted_dates = sorted(entry["dates"], reverse=True)
        last_active = sorted_dates[0]
        prev_active = sorted_dates[1] if len(sorted_dates) > 1 else ""
        active_csv = ",".join(sorted_dates)
        rows.append([
            casper_id,
            entry["agent_name"],
            entry["source_dc"],
            last_active,
            prev_active,
            active_csv,
        ])

    rows.sort(key=lambda r: (r[3], r[0]), reverse=False)
    rows.sort(key=lambda r: r[3], reverse=True)

    return {
        "success": True,
        "total_agents": len(rows),
        "dates_processed": sorted(dates_processed, reverse=True),
        "rows": rows,
    }
