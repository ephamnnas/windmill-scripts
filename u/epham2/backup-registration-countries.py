# py: 3.12
# requirements:
# requests
# wmill

"""
READ-ONLY backup of the registration country fields for every application.

Uses the same auth as the credential report script:
  * SM Apply:   Bearer token from Windmill variable u/epham2/SMAPPLY_ACCESS_TOKEN
  * SharePoint: Graph client-credentials from u/epham2/AZURE_TENANT_ID,
                u/epham2/SHAREPOINT_CLIENT_ID, u/epham2/SHAREPOINT_CLIENT_SECRET

Only GET requests are sent to SM Apply. Nothing in SM Apply is changed.

Outputs (uploaded to SharePoint, since the Windmill worker disk is temporary):
  country_fields_<ts>.csv   One row per country field: application ID, task,
                            field label, stored value and its type, dates.
                            No names or emails.
  raw_backup_<ts>.jsonl     Full JSON for every matching task (complete
                            backup; may contain personal data, keep restricted).

Run order:
  1. main(discover=True, sample_appid=<a known application ID>)
     Prints the task list and matching fields for one application so you can
     confirm TASK_NAME_FILTER and that the country answer is an integer index.
  2. main()   Full backup of programs 295865 and 295185 (see PROGRAM_IDS).
"""

import csv
import io
import json
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from typing import Any, Dict, List, Optional
from urllib.parse import quote

import requests
import wmill


# =========================================================
# CONFIG
# =========================================================

SMAPPLY_BASE = "https://nnas.smapply.io/api"
SMAPPLY_TOKEN_VAR = "u/epham2/SMAPPLY_ACCESS_TOKEN"

# Only applications in these programs are backed up.
PROGRAM_IDS = [295865, 295185]

# Only tasks whose name contains this text (case-insensitive) are backed up.
# Confirm the applicant-facing task name with discover mode first.
TASK_NAME_FILTER = "registration"
# Only fields whose label contains this text are written to the CSV.
FIELD_KEYWORD = "country"

MAX_WORKERS = 4        # parallel applications; kept low to respect rate limits
MAX_RETRIES = 5

# SharePoint destination for the backup files
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SHAREPOINT_HOSTNAME = "nnasca.sharepoint.com"
SITE_PATH = "/sites/NNAS-Portal-Sharepoint"
DOCUMENT_LIBRARY_NAME = "Applications"
BACKUP_FOLDER = "Country list backups"   # created inside the Applications library


# =========================================================
# SM APPLY (read-only)
# =========================================================

_SM_HEADERS: Optional[Dict[str, str]] = None


def _sm_headers() -> Dict[str, str]:
    global _SM_HEADERS
    if _SM_HEADERS is None:
        token = wmill.get_variable(SMAPPLY_TOKEN_VAR)
        _SM_HEADERS = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    return _SM_HEADERS


def _sm_get(url: str, params: Optional[dict] = None) -> Any:
    for attempt in range(1, MAX_RETRIES + 1):
        resp = requests.get(url, headers=_sm_headers(), params=params, timeout=120)
        if resp.ok:
            return resp.json()
        if resp.status_code in (429, 500, 502, 503, 504):
            time.sleep(int(resp.headers.get("Retry-After", 2 ** attempt)))
            continue
        raise RuntimeError(f"SMApply GET {url} failed {resp.status_code}: {resp.text[:1000]}")
    raise RuntimeError(f"SMApply GET {url} failed after {MAX_RETRIES} attempts")


def _sm_paginate(path: str, params: Optional[dict] = None) -> List[Dict[str, Any]]:
    """Follow DRF pagination via 'next', same as fetch_all_reviews()."""
    url: Optional[str] = f"{SMAPPLY_BASE}{path}"
    out: List[Dict[str, Any]] = []
    while url:
        data = _sm_get(url, params)
        params = None
        if isinstance(data, list):
            out.extend(data)
            break
        out.extend(data.get("results", []))
        url = data.get("next")
    return out


def _program_id(app: Dict[str, Any]) -> Optional[int]:
    prog = app.get("program")
    if isinstance(prog, dict):
        prog = prog.get("id")
    try:
        return int(prog)
    except (TypeError, ValueError):
        return None


def list_applications(program_ids: List[int]) -> List[Dict[str, Any]]:
    """Fetch applications program by program. Results are also checked
    client-side, so nothing outside these programs is included even if the
    API ignores the program filter."""
    wanted = {int(p) for p in program_ids}
    apps: Dict[int, Dict[str, Any]] = {}
    for pid in wanted:
        for app in _sm_paginate("/applications/", {"program": pid}):
            if _program_id(app) in wanted:
                apps[int(app["id"])] = app
    return list(apps.values())


def list_tasks(appid: int) -> List[Dict[str, Any]]:
    return _sm_paginate(f"/applications/{appid}/tasks/")


def get_task(appid: int, task_id: int) -> Dict[str, Any]:
    # Same endpoint clear_smapply_task_files() already reads successfully.
    return _sm_get(f"{SMAPPLY_BASE}/applications/{appid}/tasks/{task_id}/")


def _task_name(task: Dict[str, Any]) -> str:
    t = task.get("task") if isinstance(task.get("task"), dict) else {}
    return str(task.get("name") or task.get("title") or t.get("name") or "")


def find_country_fields(obj: Any, path: str = "") -> List[Dict[str, Any]]:
    """Find fields whose label mentions FIELD_KEYWORD. Handles the
    {label, response} shape used by the reviews endpoint, plus {label, value}
    and plain keys, so it works whichever shape the task detail returns."""
    found: List[Dict[str, Any]] = []
    if isinstance(obj, dict):
        label = str(obj.get("label") or obj.get("name") or obj.get("title") or "")
        for value_key in ("response", "value", "answer"):
            if value_key in obj and FIELD_KEYWORD in label.lower():
                found.append({"path": path, "label": label, "value": obj[value_key]})
                break
        for k, v in obj.items():
            sub = f"{path}.{k}" if path else k
            if FIELD_KEYWORD in k.lower() and not isinstance(v, (dict, list)):
                found.append({"path": sub, "label": k, "value": v})
            found.extend(find_country_fields(v, sub))
    elif isinstance(obj, list):
        for i, item in enumerate(obj):
            found.extend(find_country_fields(item, f"{path}[{i}]"))
    return found


# =========================================================
# SHAREPOINT UPLOAD (same Graph auth as the report script)
# =========================================================

def _graph_token() -> str:
    tenant = wmill.get_variable("u/epham2/AZURE_TENANT_ID").strip()
    resp = requests.post(
        f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token",
        data={
            "client_id": wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_ID").strip(),
            "client_secret": wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_SECRET").strip(),
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        },
        timeout=60,
    )
    if resp.status_code != 200:
        raise RuntimeError(f"Graph token request failed {resp.status_code}: {resp.text[:1000]}")
    return resp.json()["access_token"]


def upload_to_sharepoint(files: Dict[str, bytes]) -> Dict[str, str]:
    token = _graph_token()
    h = {"Authorization": f"Bearer {token}"}
    site = requests.get(f"{GRAPH_BASE}/sites/{SHAREPOINT_HOSTNAME}:{SITE_PATH}", headers=h, timeout=60).json()
    drives = requests.get(f"{GRAPH_BASE}/sites/{site['id']}/drives", headers=h, timeout=60).json()["value"]
    drive = next(d for d in drives if d.get("name") == DOCUMENT_LIBRARY_NAME)
    urls = {}
    for name, content in files.items():
        path = quote(f"{BACKUP_FOLDER}/{name}", safe="/")
        resp = requests.put(
            f"{GRAPH_BASE}/drives/{drive['id']}/root:/{path}:/content",
            headers={**h, "Content-Type": "application/octet-stream"},
            data=content, timeout=300,
        )
        if resp.status_code not in (200, 201):
            raise RuntimeError(f"SharePoint upload of {name} failed {resp.status_code}: {resp.text[:1000]}")
        urls[name] = resp.json().get("webUrl", "")
    return urls


# =========================================================
# BACKUP
# =========================================================

def backup_application(app: Dict[str, Any]) -> Dict[str, Any]:
    appid = int(app["id"])
    rows, raw = [], []
    for task in list_tasks(appid):
        name = _task_name(task)
        if TASK_NAME_FILTER.lower() not in name.lower():
            continue
        task_id = task.get("id")
        detail = get_task(appid, task_id) if task_id is not None else task
        raw.append({"application_id": appid, "task_id": task_id, "task_name": name, "detail": detail})
        for f in find_country_fields(detail):
            v = f["value"]
            rows.append({
                "application_id": appid,
                "program": _program_id(app),
                "task_id": task_id,
                "task_name": name,
                "field_label": f["label"],
                "field_path": f["path"],
                "stored_value": json.dumps(v) if isinstance(v, (dict, list)) else v,
                "value_type": type(v).__name__,
                "app_last_edited": app.get("last_edited") or app.get("updated_at") or app.get("modified"),
                "app_submitted": app.get("submitted_at") or app.get("date_submitted"),
                "task_last_edited": detail.get("last_edited") or detail.get("updated_at") or detail.get("modified"),
            })
    return {"rows": rows, "raw": raw}


def main(discover: bool = False,
         sample_appid: Optional[int] = None,
         upload: bool = True):

    if discover:
        if not sample_appid:
            raise ValueError("Provide sample_appid when discover=True.")
        tasks = list_tasks(int(sample_appid))
        app = _sm_get(f"{SMAPPLY_BASE}/applications/{int(sample_appid)}/")
        if _program_id(app) not in PROGRAM_IDS:
            raise ValueError(f"Application {sample_appid} is in program {_program_id(app)}, "
                             f"not one of {PROGRAM_IDS}.")
        result = backup_application(app)
        return {
            "program": _program_id(app),
            "all_task_names": [_task_name(t) for t in tasks],
            "task_name_filter": TASK_NAME_FILTER,
            "matched_tasks": [r["task_name"] for r in result["raw"]],
            "country_fields_found": result["rows"],
            "first_matched_task_detail": result["raw"][0]["detail"] if result["raw"] else None,
        }

    ts = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    apps = list_applications(PROGRAM_IDS)

    all_rows: List[Dict[str, Any]] = []
    raw_lines: List[str] = []
    failed: List[Dict[str, Any]] = []
    lock = threading.Lock()

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
        futures = {pool.submit(backup_application, a): a.get("id") for a in apps}
        for fut in as_completed(futures):
            appid = futures[fut]
            try:
                res = fut.result()
            except Exception as exc:
                failed.append({"application_id": appid, "error": str(exc)[:500]})
                continue
            with lock:
                all_rows.extend(res["rows"])
                raw_lines.extend(json.dumps(r) for r in res["raw"])

    all_rows.sort(key=lambda r: (r["application_id"], str(r["task_name"]), r["field_path"]))

    buf = io.StringIO()
    fields = ["application_id", "program", "task_id", "task_name", "field_label", "field_path",
              "stored_value", "value_type", "app_last_edited", "app_submitted", "task_last_edited"]
    writer = csv.DictWriter(buf, fieldnames=fields)
    writer.writeheader()
    writer.writerows(all_rows)

    files = {
        f"country_fields_{ts}.csv": buf.getvalue().encode("utf-8"),
        f"raw_backup_{ts}.jsonl": ("\n".join(raw_lines) + "\n").encode("utf-8"),
    }
    sharepoint_urls = upload_to_sharepoint(files) if upload else {}

    return {
        "status": "ok" if not failed else "completed_with_failures",
        "programs": PROGRAM_IDS,
        "applications_scanned": len(apps),
        "applications_per_program": {p: sum(1 for a in apps if _program_id(a) == p) for p in PROGRAM_IDS},
        "country_fields_backed_up": len(all_rows),
        "failed_applications": failed,
        "sharepoint_files": sharepoint_urls,
        "note": "Rerun for any failed applications before reverting the country lists.",
    }