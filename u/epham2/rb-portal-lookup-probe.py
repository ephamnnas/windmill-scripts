# py: 3.12
# requirements:
# wmill
# requests

"""
RBP Upload API lookup probe  (READ-ONLY, run as a standalone Windmill script)
=============================================================================

Checks whether the RB Portal API can return an application's existing records,
and in what shape, before the upload step is changed to rely on it.

Makes GET requests only. Nothing is created, updated or uploaded.

Output is limited to record keys and date/status fields. Applicant details
(name, date of birth, NNAS ID, country) are never printed; for those fields
only the field NAME is listed, so the output is safe to share.
"""

from typing import Any, Dict, List
from urllib.parse import quote

import requests
import wmill

RBP_API_BASE     = "https://rbportal.nnas.ca/wp-json/rbp/v1"
RBP_WP_USER_VAR  = "u/knguyen/RBP_WP_USER"
RBP_APP_PASS_VAR = "u/knguyen/RBP_APP_PASS"

# Only these field values are printed. Every other field shows its name only.
SAFE_FIELDS = ("post_id", "application_id", "regulatory_body", "discipline",
               "status", "issued_date", "report_expiration_date",
               "order_date", "order_expiration_date", "ported_date",
               "all_recipients_raw", "report_file")


def _summarize_record(r: Dict[str, Any]) -> Dict[str, Any]:
    return {
        "values": {k: r.get(k) for k in SAFE_FIELDS if k in r},
        "all_field_names": sorted(r.keys()),
    }


def _summarize_body(body: Any) -> Dict[str, Any]:
    if isinstance(body, list):
        recs = [x for x in body if isinstance(x, dict)]
        return {"shape": "list", "record_count": len(recs),
                "records": [_summarize_record(x) for x in recs]}
    if isinstance(body, dict):
        for k in ("records", "items", "data"):
            if isinstance(body.get(k), list):
                recs = [x for x in body[k] if isinstance(x, dict)]
                return {"shape": f"object with list under '{k}'",
                        "top_level_keys": sorted(body.keys()),
                        "record_count": len(recs),
                        "records": [_summarize_record(x) for x in recs]}
        if "code" in body and "message" in body:
            return {"shape": "wordpress error",
                    "code": body.get("code"),
                    "message": str(body.get("message"))[:200]}
        if body.get("post_id"):
            return {"shape": "single record", "records": [_summarize_record(body)]}
        return {"shape": "object (unrecognized)", "top_level_keys": sorted(body.keys())}
    return {"shape": type(body).__name__}


def _get(session: requests.Session, url: str) -> Dict[str, Any]:
    try:
        resp = session.get(url, timeout=30)
    except Exception as e:
        return {"url": url, "error": f"{type(e).__name__}: {e}"}
    out: Dict[str, Any] = {"url": url, "http_status": resp.status_code,
                           "content_type": resp.headers.get("Content-Type", "")}
    try:
        out.update(_summarize_body(resp.json()))
    except ValueError:
        out["shape"] = "not JSON"
    return out


def main(application_id: str) -> Dict[str, Any]:
    """
    application_id: an application you know already has portal records,
    ideally one with at least one RB already at "Report Issued".
    """
    session = requests.Session()
    session.auth = (wmill.get_variable(RBP_WP_USER_VAR),
                    wmill.get_variable(RBP_APP_PASS_VAR))
    app = quote(str(application_id).strip(), safe="")

    # 1. The namespace index lists every route the plugin registers and the
    #    HTTP methods each accepts. This shows whether any read route exists.
    routes: List[Dict[str, Any]] = []
    index_error = None
    try:
        idx = session.get(RBP_API_BASE, timeout=30)
        data = idx.json() if idx.ok else {}
        for path, info in (data.get("routes") or {}).items():
            methods = sorted({m for ep in info.get("endpoints", [])
                              for m in ep.get("methods", [])})
            routes.append({"route": path, "methods": methods})
        if not idx.ok:
            index_error = f"HTTP {idx.status_code}"
    except Exception as e:
        index_error = f"{type(e).__name__}: {e}"

    read_routes = [r for r in routes if "GET" in r["methods"]]

    # 2. Try the likely read paths for this application (GET only).
    candidates = [
        f"{RBP_API_BASE}/record/{app}",
        f"{RBP_API_BASE}/records/{app}",
        f"{RBP_API_BASE}/records?application_id={app}",
        f"{RBP_API_BASE}/record?application_id={app}",
    ]
    attempts = [_get(session, u) for u in candidates]

    working = [a for a in attempts
               if a.get("http_status") == 200 and a.get("records")]

    return {
        "summary": {
            "routes_found": len(routes),
            "routes_accepting_GET": [r["route"] for r in read_routes],
            "lookup_works": bool(working),
            "working_url": working[0]["url"] if working else None,
            "dates_present": bool(working) and any(
                rec["values"].get("issued_date") or rec["values"].get("report_expiration_date")
                for rec in working[0]["records"]),
        },
        "namespace_index_error": index_error,
        "all_routes": routes,
        "attempts": attempts,
    }