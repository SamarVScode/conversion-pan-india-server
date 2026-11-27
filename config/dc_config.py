import json
import logging
import os
from typing import Any, Dict

log = logging.getLogger("conversion_report_server.dc_config")

# ---------------------------------------------------------------------------
# DC Alias Normalization (All DCs retained; no allowlist filter)
# ---------------------------------------------------------------------------

DEFAULT_DC_ALIASES: Dict[str, str] = {
    "CAR-KHR": "CAR",
    "CAR_KHR": "CAR",
    "CARKHR": "CAR",
}


def _load_dc_aliases() -> Dict[str, str]:
    aliases = dict(DEFAULT_DC_ALIASES)
    raw_json = os.getenv("DC_ALIASES_JSON", "").strip()
    if raw_json:
        try:
            parsed = json.loads(raw_json)
            if isinstance(parsed, dict):
                for k, v in parsed.items():
                    if k and v:
                        aliases[str(k)] = str(v)
        except Exception as e:
            log.warning(f"Invalid DC_ALIASES_JSON ignored: {e}")
    return aliases


DC_ALIASES: Dict[str, str] = _load_dc_aliases()
DC_ALIASES_MAP: Dict[str, str] = {
    k.strip().upper(): v.strip().upper() for k, v in DC_ALIASES.items()
}


def normalize_dc_code(raw_dc: Any) -> str:
    """Normalize a DC code using uppercase trimming and configured DC_ALIASES."""
    if raw_dc is None:
        return ""
    s = str(raw_dc).strip().upper()
    if not s or s in ("NONE", "NAN", "NAT"):
        return ""
    if s in DC_ALIASES_MAP:
        return DC_ALIASES_MAP[s]
    s_hyphen = s.replace("_", "-")
    if s_hyphen in DC_ALIASES_MAP:
        return DC_ALIASES_MAP[s_hyphen]
    s_under = s.replace("-", "_")
    if s_under in DC_ALIASES_MAP:
        return DC_ALIASES_MAP[s_under]
    return s
