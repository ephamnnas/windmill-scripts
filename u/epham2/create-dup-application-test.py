"""
Duplicate an SM Apply application within the same program.

Run as a Windmill script: main(source_id, dry_run).
Start with dry_run=True to preview, then set dry_run=False to create the duplicate.

Logging intentionally prints task IDs, field keys and counts only, never
response values, so applicant personal information is not written to logs.
"""

import time

import requests
import wmill

# =========================================================
# SMAPPLY CONFIG
# =========================================================

SMAPPLY_BASE = "https://nnas.smapply.io/api"
SMAPPLY_REVIEWS_TOKEN_VAR = "u/epham2/SMAPPLY_ACCESS_TOKEN"

# =========================================================
# DUPLICATION CONFIG
# =========================================================

# Status applied to each task that receives copied data.
# 2 = In progress (applicant reviews and submits), 3 = Completed.
COPIED_TASK_STATUS = 2

# Task types copied by this script. Type 1 = form. Type 2 (admin file tasks)
# and type 10 (payments) are always skipped.
COPYABLE_TASK_TYPES = {1}

# Whole tasks to skip, matched by task name (case-insensitive).
# Names are used rather than IDs because task IDs differ between programs.
EXCLUDED_TASK_NAMES = {
    "Terms and Conditions",   # applicant re-signs
    "Application Sign Off",   # applicant re-signs
    # "Select Nursing Discipline",  # uncomment if this should differ
}

# Individual fields to skip within otherwise copied tasks: {task_name: {field_keys}}
# Field keys may also differ between programs, so confirm them for each program.
EXCLUDED_FIELDS = {
    # "Personal Information": {"fieldKeyHere"},
}


def normalize(name):
    return (name or "").strip().lower()


_EXCLUDED_TASKS_NORMALIZED = {normalize(n) for n in EXCLUDED_TASK_NAMES}
_EXCLUDED_FIELDS_NORMALIZED = {normalize(k): v for k, v in EXCLUDED_FIELDS.items()}

# Custom field IDs to copy onto the new application (empty = copy none).
COPY_CUSTOM_FIELD_IDS = set()

COPY_COLLABORATORS = False

# =========================================================
# SMAPPLY API HELPERS
# =========================================================

def get_session():
    token = wmill.get_variable(SMAPPLY_REVIEWS_TOKEN_VAR)
    session = requests.Session()
    session.headers.update({
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    })
    # No "apply-automations" header is sent, so task automations do not fire.
    return session


def request(session, method, url, **kwargs):
    """Send a request with simple retry on rate limiting and server errors."""
    if url.startswith("/"):
        url = SMAPPLY_BASE + url
    for attempt in range(5):
        resp = session.request(method, url, timeout=60, **kwargs)
        if resp.status_code == 429 or resp.status_code >= 500:
            wait = int(resp.headers.get("Retry-After", 2 ** attempt))
            print(f"  {resp.status_code} received, retrying in {wait}s")
            time.sleep(wait)
            continue
        if not resp.ok:
            raise RuntimeError(f"{method} {url} failed: {resp.status_code} {resp.text[:300]}")
        return resp.json() if resp.content else None
    raise RuntimeError(f"{method} {url} failed after retries")


def get_all_tasks(session, application_id):
    """Fetch every task for an application, following pagination."""
    tasks, url = [], f"/applications/{application_id}/tasks/"
    while url:
        page = request(session, "GET", url)
        tasks.extend(page["results"])
        url = page.get("next")  # absolute URL, used as-is
    return tasks

# =========================================================
# DATA TRANSFORMATION
# =========================================================

def is_empty(value):
    """True for unanswered values, which are skipped rather than written."""
    if value is None or value == "" or value == -1 or value == []:
        return True
    if isinstance(value, dict):
        other = value.get("other")
        other_empty = not other or (isinstance(other, list) and not any(other))
        return value.get("value") in (None, -1, []) and other_empty
    return False


def is_signature(value):
    return isinstance(value, str) and value.startswith("/protected/media-file/")


def build_patch_data(task):
    """Convert GET task data into a PATCH data object. Returns (data, skipped_files)."""
    excluded = _EXCLUDED_FIELDS_NORMALIZED.get(normalize(task["name"]), set())
    data, skipped_files = {}, []
    for key, entry in (task.get("data") or {}).items():
        if key in excluded:
            continue
        # File fields come back as a raw list of file objects, not {label, response}.
        if not isinstance(entry, dict) or "response" not in entry:
            skipped_files.append(key)
            continue
        value = entry["response"]
        if is_empty(value) or is_signature(value):
            continue
        data[key] = value
    return data, skipped_files

# =========================================================
# MAIN
# =========================================================

def main(source_id: int, dry_run: bool = True):
    session = get_session()
    source = request(session, "GET", f"/applications/{source_id}/")

    payload = {
        "applicant": source["applicant"]["id"],
        "program": source["program"]["id"],
        "title": source.get("title") or "",
        "state": 0,
    }
    if source.get("organization"):
        payload["organization"] = source["organization"]["id"]
    if COPY_COLLABORATORS:
        payload["collaborators"] = [c["id"] for c in source.get("collaborators", [])]
    custom = [
        {"id": cf["id"], "value": cf["value"]}
        for cf in source.get("custom_fields", [])
        if cf["id"] in COPY_CUSTOM_FIELD_IDS
    ]
    if custom:
        payload["custom_fields"] = custom

    source_tasks = get_all_tasks(session, source_id)
    print(f"Source application {source_id}: {len(source_tasks)} tasks found")

    if dry_run:
        new_id, target_task_ids = None, {t["id"] for t in source_tasks}
        print("DRY RUN: no application will be created")
    else:
        new_id = request(session, "POST", "/applications/", json=payload)["id"]
        target_task_ids = {t["id"] for t in get_all_tasks(session, new_id)}
        print(f"Created application {new_id}")

    copied, skipped, errors, files_to_reupload = [], [], [], []
    for task in source_tasks:
        label = f"[{task['id']}] {task['name']}"
        if task["type"] not in COPYABLE_TASK_TYPES:
            skipped.append(f"{label}: task type {task['type']}")
            continue
        if normalize(task["name"]) in _EXCLUDED_TASKS_NORMALIZED:
            skipped.append(f"{label}: excluded")
            continue
        if task["id"] not in target_task_ids:
            errors.append(f"{label}: not present on new application")
            continue

        data, skipped_files = build_patch_data(task)
        files_to_reupload += [f"task {task['id']}, field {key}" for key in skipped_files]
        if not data:
            skipped.append(f"{label}: no answered fields")
            continue

        if not dry_run:
            try:
                request(session, "PATCH", f"/applications/{new_id}/tasks/{task['id']}/",
                        json={"status": COPIED_TASK_STATUS, "data": data})
            except RuntimeError as err:
                errors.append(f"{label}: {err}")
                continue
        copied.append(f"{label}: {len(data)} fields")

    for line in copied:
        print(f"  copy {line}")
    for line in skipped:
        print(f"  skip {line}")
    for line in errors:
        print(f"  ERROR {line}")

    # Returned as the Windmill job result (IDs and counts only, no response values).
    return {
        "source_application_id": source_id,
        "new_application_id": new_id,
        "dry_run": dry_run,
        "copied_tasks": copied,
        "skipped_tasks": skipped,
        "errors": errors,
        "files_to_reupload": files_to_reupload,
    }