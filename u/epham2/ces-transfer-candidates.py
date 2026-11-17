# py: 3.12
# requirements:
# wmill
# requests

"""
CES -> TM transfer candidates (read-only)
=========================================

Lists every CES application and decides, per application:

  scope        in_scope       unpaid, or paid AFTER paid_after_date
               out_of_scope   paid on or before paid_after_date, archived or
                              trashed, or already has a TM application
               needs_review   paid but no payment date found, or invoiced only

  eligibility  tm_eligible_now        meets the TM criteria as answered
               sanctions_connected    Yes to a sanctions question
               unsure_needs_review    Unsure to the jurisdiction question
               not_eligible_other     answered No but fails TM for another reason
               no_sanctions_answer / eligibility_incomplete / no_eligibility_record

  ready        in_scope AND tm_eligible_now: the clone script will create the
               TM application without any eligibility change.

Payment evidence, as in the clone script: a completed Expedited Service
Payment task (any variant), then the Payment date custom field, then an
Invoiced label.

GET requests only. Output holds IDs, references, statuses, dates and
categories; no names, emails or free-text answers.
"""

import json
import re
import threading
import time
from collections import Counter
from concurrent.futures import ThreadPoolExecutor
from typing import Any, Dict, List, Optional, Tuple

import requests
import wmill

BASE = "https://nnas.smapply.io/api"
TOKEN_PATH = "u/epham2/SMAPPLY_ACCESS_TOKEN"
CES_PROGRAM_ID = 295865
TM_PROGRAM_ID = 295185
PAYMENT_DATE_FIELD = "Payment date"
TASK_STATUS_COMPLETE = 3
STATE_NAMES = {0: "active", 1: "locked", 2: "hidden", 3: "archived", 4: "trashed"}

_TOKEN: Optional[str] = None
_LOCK = threading.Lock()


# --------------------------------------------------------------------------
# http
# --------------------------------------------------------------------------

def _get(path: str, params: Optional[dict] = None) -> Tuple[Optional[int], Any]:
    global _TOKEN
    with _LOCK:
        if _TOKEN is None:
            _TOKEN = wmill.get_variable(TOKEN_PATH)
    url = path if path.startswith("http") else f"{BASE}{path}"
    for attempt in range(5):
        try:
            r = requests.get(url, params=params, timeout=90,
                             headers={"Authorization": f"Bearer {_TOKEN}",
                                      "Accept": "application/json"})
        except Exception:
            time.sleep(2 * (attempt + 1))
            continue
        if r.status_code == 429 or r.status_code >= 500:
            time.sleep(int(r.headers.get("Retry-After", 2 ** attempt)))
            continue
        try:
            return r.status_code, r.json()
        except Exception:
            return r.status_code, None
    return None, None


def _paginate(path: str, params: Optional[dict] = None) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    url: Optional[str] = path
    while url:
        status, body = _get(url, params)
        if status != 200 or not isinstance(body, dict):
            raise RuntimeError(f"Listing {path} stopped at HTTP {status} after {len(rows)} rows.")
        rows.extend(r for r in body.get("results", []) if isinstance(r, dict))
        url, params = body.get("next"), None
    return rows


def _id(v: Any) -> Any:
    return v.get("id") if isinstance(v, dict) else v


def _norm(s: Any) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip().casefold()


# --------------------------------------------------------------------------
# payment
# --------------------------------------------------------------------------

def _iso_date(value: Any) -> str:
    """Normalise a date to YYYY-MM-DD so the cutoff comparison is reliable.
    Returns "" if the value is not a recognisable date."""
    from datetime import datetime
    text = str(value or "").strip()
    if not text:
        return ""
    if re.match(r"\d{4}-\d{2}-\d{2}", text):
        return text[:10]
    for fmt in ("%m/%d/%Y", "%d/%m/%Y", "%Y/%m/%d", "%b %d, %Y", "%d %b %Y", "%Y %b %d"):
        try:
            return datetime.strptime(text[:20].strip(), fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return ""


def payment_status(app: Dict[str, Any], tasks: List[Dict[str, Any]]) -> Dict[str, Any]:
    custom = {cf.get("name"): cf.get("value") for cf in (app.get("custom_fields") or [])
              if isinstance(cf, dict)}
    raw_field = str(custom.get(PAYMENT_DATE_FIELD) or "").strip()
    field_date = _iso_date(raw_field)
    pay = [t for t in tasks if "expedited service payment" in _norm(t.get("name"))]
    done = [t for t in pay if t.get("status") == TASK_STATUS_COMPLETE]
    if done:
        task_date = _iso_date(done[0].get("completed_at"))
        return {"paid": "paid", "payment_date": field_date or task_date,
                "basis": "completed payment task"}
    if field_date:
        return {"paid": "paid", "payment_date": field_date, "basis": "Payment date field"}
    if raw_field:
        return {"paid": "paid", "payment_date": "",
                "basis": "Payment date field in an unrecognised format"}
    if "invoiced" in json.dumps(app.get("labels") or []).lower():
        return {"paid": "invoiced", "payment_date": "", "basis": "Invoiced label"}
    return {"paid": "unpaid", "payment_date": "",
            "basis": "payment task not completed" if pay else "no payment task"}


# --------------------------------------------------------------------------
# eligibility
# --------------------------------------------------------------------------

QUESTIONS = {
    "prior_report": ("have you completed and receiv", ["Yes", "No"]),
    "ofac": ("are any of the documents, cred", ["Yes", "No"]),
    "jurisdiction": ("is any part of your application c", ["Yes", "No", "Unsure"]),
}


def eligibility(user_id: Any) -> Dict[str, Any]:
    status, body = _get(f"/users/{user_id}/eligibility/")
    if status != 200 or not isinstance(body, dict):
        return {"eligibility": "no_eligibility_record", "detail": f"HTTP {status}"}
    answers = {k: "(not asked)" for k in QUESTIONS}
    for item in (body.get("responses") or {}).values():
        if not isinstance(item, dict):
            continue
        label = _norm(item.get("label") or item.get("name"))
        for key, (prefix, choices) in QUESTIONS.items():
            if label.startswith(prefix):
                v = item.get("response")
                if isinstance(v, int) and not isinstance(v, bool):
                    answers[key] = choices[v] if 0 <= v < len(choices) else "(unanswered)"
    tm_ok = answers["jurisdiction"] == "No" or (
        answers["prior_report"] == "Yes" and answers["ofac"] == "No")
    if body.get("is_completed") is False:
        cat = "eligibility_incomplete"
    elif "Yes" in (answers["jurisdiction"], answers["ofac"]):
        cat = "sanctions_connected"
    elif answers["jurisdiction"] == "Unsure":
        cat = "unsure_needs_review"
    elif tm_ok:
        cat = "tm_eligible_now"
    elif "No" in (answers["jurisdiction"], answers["ofac"]):
        cat = "not_eligible_other"
    else:
        cat = "no_sanctions_answer"
    return {"eligibility": cat, "jurisdiction": answers["jurisdiction"],
            "ofac": answers["ofac"], "prior_report": answers["prior_report"]}


# --------------------------------------------------------------------------
# existing TM application
# --------------------------------------------------------------------------

def existing_tm(applicant_id: Any) -> Optional[str]:
    """Filter params are undocumented, so every row is re-checked."""
    status, body = _get("/applications/", {"program": TM_PROGRAM_ID,
                                           "applicant": applicant_id, "per_page": 100})
    if status != 200 or not isinstance(body, dict):
        return "lookup_failed"
    for row in body.get("results") or []:
        if (isinstance(row, dict) and str(_id(row.get("applicant"))) == str(applicant_id)
                and str(_id(row.get("program"))) == str(TM_PROGRAM_ID)):
            return row.get("reference_id") or str(row.get("id"))
    return None


# --------------------------------------------------------------------------
# per application
# --------------------------------------------------------------------------

def assess(app_id: int, paid_after_date: str) -> Dict[str, Any]:
    status, app = _get(f"/applications/{app_id}/")
    if status != 200 or not isinstance(app, dict):
        return {"application_id": app_id, "scope": "error", "detail": f"application HTTP {status}"}
    applicant_id = _id(app.get("applicant"))
    st = app.get("status")
    row: Dict[str, Any] = {
        "application_id": app_id,
        "reference_id": app.get("reference_id"),
        "applicant_id": applicant_id,
        "application_status": (st.get("name") if isinstance(st, dict) else st) or "",
        "state": STATE_NAMES.get(app.get("state"), app.get("state")),
        "submitted": bool(app.get("last_submitted_at")),
    }
    try:
        tasks = _paginate(f"/applications/{app_id}/tasks/")
    except RuntimeError as exc:
        row.update(scope="error", detail=str(exc)[:200])
        return row
    row.update(payment_status(app, tasks))
    row["existing_tm_application"] = existing_tm(applicant_id) or ""
    row.update(eligibility(applicant_id))

    if row["state"] in ("archived", "trashed"):
        scope, why = "out_of_scope", f"application {row['state']}"
    elif row["existing_tm_application"] and row["existing_tm_application"] != "lookup_failed":
        scope, why = "out_of_scope", "already has a TM application"
    elif row["paid"] == "unpaid":
        scope, why = "in_scope", "unpaid"
    elif row["paid"] == "invoiced":
        scope, why = "needs_review", "invoiced, no completed payment"
    elif not row["payment_date"]:
        scope, why = "needs_review", "paid, payment date not found"
    elif row["payment_date"] > paid_after_date:
        scope, why = "in_scope", f"paid {row['payment_date']}, after cutoff"
    else:
        scope, why = "out_of_scope", f"paid {row['payment_date']}, on or before cutoff"
    if row["existing_tm_application"] == "lookup_failed":
        why += "; TM lookup failed, check manually"
    row["scope"], row["scope_reason"] = scope, why
    row["ready"] = scope == "in_scope" and row.get("eligibility") == "tm_eligible_now"
    return row


def main(program_id: int = CES_PROGRAM_ID, paid_after_date: str = "2026-11-13",
         limit: int = 0, max_workers: int = 4, include_out_of_scope: bool = False) -> dict:
    """
    paid_after_date       Paid applications are in scope only if paid AFTER
                          this date (YYYY-MM-DD, exclusive).
    limit                 Assess only the first N applications (0 = all).
    max_workers           Parallel requests; keep low for rate limits.
    include_out_of_scope  Also list out-of-scope rows (counts are always shown).
    """
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", paid_after_date or ""):
        return {"status": "error", "message": "paid_after_date must be YYYY-MM-DD."}
    started = time.time()
    listing = _paginate(f"/programs/{program_id}/applications/", {"per_page": 100})
    ids = [r["id"] for r in listing if r.get("id")]
    if limit:
        ids = ids[:limit]

    with ThreadPoolExecutor(max_workers=max(1, max_workers)) as pool:
        rows = list(pool.map(lambda i: assess(i, paid_after_date), ids))

    order = {"in_scope": 0, "needs_review": 1, "error": 2, "out_of_scope": 3}
    rows.sort(key=lambda r: (order.get(r["scope"], 9), not r.get("ready"),
                             r.get("eligibility", ""), r.get("reference_id") or ""))

    matrix = Counter((r["scope"], r.get("eligibility", "-")) for r in rows)
    return {
        "status": "ok",
        "paid_after_date": paid_after_date,
        "applications_assessed": len(rows),
        "ready_to_transfer": sum(1 for r in rows if r.get("ready")),
        "by_scope": dict(Counter(r["scope"] for r in rows)),
        "by_scope_and_eligibility": [
            {"scope": s, "eligibility": e, "applications": n}
            for (s, e), n in sorted(matrix.items(), key=lambda kv: (order.get(kv[0][0], 9), kv[0][1]))],
        "rows": [r for r in rows if include_out_of_scope or r["scope"] != "out_of_scope"],
        "elapsed_s": round(time.time() - started, 1),
    }