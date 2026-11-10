# py: 3.12
# requirements:
# wmill
# requests

"""
NNAS RB Portal: correct the dates on ONE existing record  (one-off Windmill script)
==================================================================================

Uses the same RBP Upload API route and credentials as nnas_rb_portal_upload.py:
    POST /wp-json/rbp/v1/record/<application_id>   (_match_by_regulatory_body)

The record is addressed exactly the way the delivery script addressed it:
application_id + regulatory_body. Take both values, and the post_id, from the
"results" of the delivery run that created the record you want to keep.

Safety:
  * dry_run defaults to True: nothing is written, the request is only shown.
  * Step 1 sends a bind with NO changes and checks that the portal returns the
    expected post_id. If it returns a different post_id, or reports that it
    created a new row, the script stops before writing any dates.
  * Only the date fields you supply are sent. Blank inputs are not sent, so
    they cannot blank an existing value.
"""

import re
from typing import Any, Dict, Optional
from urllib.parse import quote

import wmill
import requests

RBP_API_BASE     = "https://rbportal.nnas.ca/wp-json/rbp/v1"
RBP_WP_USER_VAR  = "u/knguyen/RBP_WP_USER"
RBP_APP_PASS_VAR = "u/knguyen/RBP_APP_PASS"
HTTP_TIMEOUT     = 30


def _ymd(v: Optional[str]) -> str:
    """Accept YYYY-MM-DD (or YYYY/MM/DD); return '' for blank, raise on bad input."""
    v = (v or "").strip()
    if not v:
        return ""
    m = re.match(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$", v)
    if not m:
        raise ValueError(f"'{v}' is not a date in YYYY-MM-DD format")
    return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"


def _post(application_id: str, body: Dict[str, Any], auth) -> Dict[str, Any]:
    url = f"{RBP_API_BASE}/record/{quote(application_id, safe='')}"
    r = requests.post(url, json=body, auth=auth, timeout=HTTP_TIMEOUT)
    try:
        data = r.json()
    except ValueError:
        data = {"raw": r.text[:500]}
    if not r.ok:
        raise RuntimeError(f"RB portal returned {r.status_code}: {data}")
    return data if isinstance(data, dict) else {"raw": data}


def main(application_id: str,
         regulatory_body: str,
         expected_post_id: int,
         issued_date: str = "",
         order_date: str = "",
         order_expiration_date: str = "",
         dry_run: bool = True):
    application_id  = (application_id or "").strip()
    regulatory_body = (regulatory_body or "").strip()
    if not application_id or not regulatory_body or not expected_post_id:
        return {"status": "error",
                "message": "application_id, regulatory_body and expected_post_id are required."}

    try:
        dates = {
            "issued_date":           _ymd(issued_date),
            "order_date":            _ymd(order_date),
            "order_expiration_date": _ymd(order_expiration_date),
        }
    except ValueError as e:
        return {"status": "error", "message": str(e)}
    dates = {k: v for k, v in dates.items() if v}
    if not dates:
        return {"status": "error", "message": "No dates supplied; nothing to update."}

    match = {"regulatory_body": regulatory_body, "_match_by_regulatory_body": True}
    update_body = {**match, **dates}

    if dry_run:
        return {"status": "dry_run",
                "application_id": application_id,
                "expected_post_id": int(expected_post_id),
                "would_post": update_body,
                "note": "Set dry_run=false to apply."}

    auth = (wmill.get_variable(RBP_WP_USER_VAR), wmill.get_variable(RBP_APP_PASS_VAR))
    if not auth[0] or not auth[1]:
        return {"status": "error", "message": "RB portal credentials not resolved."}

    # Step 1: bind with no changes, confirm it is the expected row.
    probe = _post(application_id, dict(match), auth)
    if probe.get("created"):
        return {"status": "stopped",
                "message": ("No existing record matched this application_id + "
                            "regulatory_body, and the portal created a new one. "
                            "No dates were written. Remove the new row "
                            f"(post_id {probe.get('post_id')}) and check the "
                            "regulatory_body value."),
                "probe": probe}
    if str(probe.get("post_id")) != str(expected_post_id):
        return {"status": "stopped",
                "message": (f"The portal matched post_id {probe.get('post_id')}, "
                            f"not {expected_post_id}. No dates were written."),
                "probe": probe}

    # Step 2: write only the supplied dates to that row.
    result = _post(application_id, update_body, auth)
    return {"status": "ok" if str(result.get("post_id")) == str(expected_post_id) else "check",
            "application_id": application_id,
            "post_id": result.get("post_id"),
            "updated": dates,
            "response": result}