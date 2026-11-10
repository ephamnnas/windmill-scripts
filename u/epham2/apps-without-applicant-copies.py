# py: 3.12
# requirements:
# requests
# wmill

"""
Scan SM Apply program 295185 for applications currently in the
"Credential Report Issued" stage whose task 2077121 is still empty.

Auth matches the credential report script: a bearer token read from the
Windmill variable u/epham2/SMAPPLY_ACCESS_TOKEN.

Output contains application IDs and task status only (no applicant PII).
"""

import time
from typing import Any, Dict, List, Optional, Set

import requests
import wmill


# =========================================================
# CONFIG
# =========================================================

SMAPPLY_BASE = "https://nnas.smapply.io/api"
SMAPPLY_TOKEN_VAR = "u/epham2/SMAPPLY_ACCESS_TOKEN"

DEFAULT_PROGRAM_ID = 295185
DEFAULT_STAGE_NAME = "Credential Report Issued"
DEFAULT_TASK_ID = 2077121

# Same completed status value used by mark_smapply_task_complete()
SMAPPLY_STATUS_COMPLETED = 3

# Same file-detection keys used by _extract_task_files() in the report script
_SMAPPLY_FILE_SIGNATURE_KEYS = {
    "file", "filename", "original_filename", "download_url", "url",
    "content_type", "mimetype", "size", "filesize", "extension",
}


# =========================================================
# HTTP HELPERS
# =========================================================

def _smapply_headers() -> Dict[str, str]:
    access_token = wmill.get_variable(SMAPPLY_TOKEN_VAR)
    if not access_token:
        raise RuntimeError(f"Missing Windmill variable: {SMAPPLY_TOKEN_VAR}")
    return {
        "Authorization": f"Bearer {access_token.strip()}",
        "Accept": "application/json",
    }


def _get(session: requests.Session, url: str, params: Optional[dict] = None) -> Optional[Any]:
    """GET with retry on rate limits. Returns None on 404."""
    if not url.startswith("http"):
        url = f"{SMAPPLY_BASE}{url}"
    for attempt in range(5):
        resp = session.get(url, params=params, timeout=120)
        if resp.status_code == 429:
            time.sleep(int(resp.headers.get("Retry-After", 2 ** attempt)))
            continue
        if resp.status_code == 404:
            return None
        if not resp.ok:
            raise RuntimeError(
                f"SMApply GET failed {resp.status_code} for {url}: {resp.text[:2000]}"
            )
        return resp.json()
    raise RuntimeError(f"SMApply GET retried too many times: {url}")


def _paginate(session: requests.Session, path: str, params: Optional[dict] = None):
    """Follow DRF pagination via `next`, same as fetch_all_reviews()."""
    url = path
    while url:
        data = _get(session, url, params=params)
        if not data:
            return
        for item in data.get("results", []):
            yield item
        url = data.get("next")   # absolute URL or None
        params = None            # `next` already carries the query


# =========================================================
# TASK EVALUATION
# =========================================================

def _extract_task_files(task_json: Any) -> List[Dict[str, Any]]:
    """Same lookup order as the report script: data.files, files, then a
    recursive search for file-like objects."""
    if not isinstance(task_json, dict):
        return []

    data = task_json.get("data")
    if isinstance(data, dict) and isinstance(data.get("files"), list):
        return [f for f in data["files"] if isinstance(f, dict) and f.get("id") is not None]
    if isinstance(task_json.get("files"), list):
        return [f for f in task_json["files"] if isinstance(f, dict) and f.get("id") is not None]

    found: List[Dict[str, Any]] = []
    seen: Set[Any] = set()

    def walk(obj):
        if isinstance(obj, dict):
            if "id" in obj and any(k in obj for k in _SMAPPLY_FILE_SIGNATURE_KEYS):
                fid = obj.get("id")
                if fid is not None and fid not in seen:
                    seen.add(fid)
                    found.append(obj)
                return
            for v in obj.values():
                walk(v)
        elif isinstance(obj, list):
            for v in obj:
                walk(v)

    walk(task_json.get("data"))
    return found


def _has_any_value(value: Any) -> bool:
    if value is None:
        return False
    if isinstance(value, str):
        return value.strip() != ""
    if isinstance(value, (list, tuple)):
        return any(_has_any_value(v) for v in value)
    if isinstance(value, dict):
        return any(_has_any_value(v) for v in value.values())
    return True


def _task_status_value(task: Dict[str, Any]) -> Any:
    status = task.get("status")
    if isinstance(status, dict):
        return status.get("id", status.get("value", status.get("name")))
    return status


def _is_completed(status: Any) -> bool:
    try:
        return int(status) == SMAPPLY_STATUS_COMPLETED
    except (TypeError, ValueError):
        return str(status or "").strip().lower() in ("complete", "completed")


def _evaluate_task(task: Dict[str, Any]) -> Dict[str, Any]:
    """A task counts as empty when it is not completed AND has no uploaded
    files AND no saved field answers."""
    status = _task_status_value(task)
    files = _extract_task_files(task)
    data = task.get("data")
    if isinstance(data, dict):
        data = {k: v for k, v in data.items() if k != "files"}
    has_answers = _has_any_value(data)

    return {
        "status": status,
        "completed": _is_completed(status),
        "file_count": len(files),
        "has_answers": has_answers,
        "is_empty": not _is_completed(status) and not files and not has_answers,
    }


def _stage_id(app: Dict[str, Any]) -> str:
    stage = app.get("current_stage") or {}
    if isinstance(stage, dict):
        return str(stage.get("id") or "").strip()
    if isinstance(stage, int):
        return str(stage)
    return ""


def _stage_title(app: Dict[str, Any]) -> str:
    stage = app.get("current_stage") or {}
    if isinstance(stage, dict):
        return (stage.get("title") or stage.get("name") or "").strip()
    return str(stage).strip()


# =========================================================
# ENTRYPOINT
# =========================================================

def main(program_id: int = DEFAULT_PROGRAM_ID,
         stage_name: str = DEFAULT_STAGE_NAME,
         task_id: int = DEFAULT_TASK_ID,
         inspect_appid: Optional[int] = None):
    """
    inspect_appid: pass one application ID to return the raw task JSON for
    that application instead of running the scan. Useful for confirming the
    status / data shape before trusting the results.
    """
    session = requests.Session()
    session.headers.update(_smapply_headers())

    if inspect_appid:
        return {
            "appid": inspect_appid,
            "task_id": task_id,
            "task": _get(session, f"/applications/{inspect_appid}/tasks/{task_id}/"),
        }

    # Match on the stage title only (case-insensitive).
    target_stage = str(stage_name).strip().lower()
    scanned = 0
    in_stage = 0
    empty: List[Dict[str, Any]] = []
    not_assigned: List[int] = []
    errors: List[Dict[str, Any]] = []
    stages_seen: Dict[str, int] = {}
    stage_ids: Dict[str, str] = {}

    for app in _paginate(session, "/applications/", params={"program": program_id}):
        scanned += 1
        title = _stage_title(app)
        sid = _stage_id(app)
        stages_seen[title] = stages_seen.get(title, 0) + 1
        if sid:
            stage_ids[title] = sid

        if title.lower() != target_stage:
            continue
        in_stage += 1

        app_id = app.get("id")
        try:
            task = _get(session, f"/applications/{app_id}/tasks/{task_id}/")
        except Exception as exc:
            errors.append({"application_id": app_id, "error": str(exc)[:500]})
            continue

        if task is None:
            not_assigned.append(app_id)
            continue

        result = _evaluate_task(task)
        if result["is_empty"]:
            empty.append({"application_id": app_id, "task_status": result["status"]})

    return {
        "program_id": program_id,
        "stage_name": stage_name,
        "task_id": task_id,
        "applications_scanned": scanned,
        "applications_in_stage": in_stage,
        "empty_task_count": len(empty),
        "empty_task_application_ids": [e["application_id"] for e in empty],
        "empty_task_details": empty,
        "task_not_assigned_application_ids": not_assigned,
        "errors": errors,
        "stage_counts": stages_seen,
        "stage_ids": stage_ids,
    }