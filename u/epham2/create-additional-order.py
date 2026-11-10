"""
Duplicate an SM Apply application within the same program, including
files uploaded to form questions (such as identification documents).

Run as a Windmill script: main(source_id, dry_run, require_request).
Works for both the TM and CES programs. By default the source application
must have a completed "Place a New Order" task (matched by name, or by the
per-program IDs below); set require_request=False for staff-initiated overrides.
Creates the duplicate by default; pass dry_run=True to preview without making changes.

Logging intentionally prints task IDs, field keys and counts only, never
response values or file names, so applicant personal information is not
written to logs or job results.
"""

import mimetypes
import time

import requests
import wmill

# =========================================================
# SMAPPLY CONFIG
# =========================================================

SMAPPLY_HOST = "https://nnas.smapply.io"
SMAPPLY_BASE = f"{SMAPPLY_HOST}/api"
SMAPPLY_TOKEN_VAR = "u/epham2/SMAPPLY_ACCESS_TOKEN"

# =========================================================
# DUPLICATION CONFIG
# =========================================================

# Status applied to each task that receives copied data.
# 2 = In progress (applicant reviews and submits), 3 = Completed.
COPIED_TASK_STATUS = 2

# Task types copied by this script. Type 1 = form. Type 2 (admin file tasks)
# and type 10 (payments) are always skipped.
COPYABLE_TASK_TYPES = {1}

# Copy files uploaded to form questions (for example, ID documents).
COPY_FILES = True

# The "Place a New Order" task in the Credential Report Issued stage.
# The applicant (or an admin on their behalf) completes it to request a duplicate.
# It is matched by name, so the same script works for TM and CES. Known task
# IDs are also matched, in case the task is renamed in one program.
REQUEST_TASK_NAMES = {"Place a New Order"}
REQUEST_TASK_IDS = {
    2116703,  # TM (Expedited Application (T), program 295185)
    # CES: add the CES "Place a New Order" task ID here once it exists
}
COMPLETED_STATUS = 3

# Whole tasks to skip, matched by task name (case-insensitive).
# Names are used rather than IDs because task IDs differ between programs.
EXCLUDED_TASK_NAMES = {
    "Terms and Conditions",   # applicant re-signs
    "Application Sign Off",   # applicant re-signs
    "Select Nursing Discipline",  # applicant chooses profession and regulatory bodies for the new order
    "Place a New Order",      # the request itself is never copied
}

# Individual fields to skip within otherwise copied tasks: {task_name: {field_keys}}
# Field keys may also differ between programs, so confirm them for each program.
EXCLUDED_FIELDS = {
    # "Personal Information": {"fieldKeyHere"},
}

# Custom field IDs to copy onto the new application (empty = copy none).
COPY_CUSTOM_FIELD_IDS = set()

COPY_COLLABORATORS = False


def normalize(name):
    return (name or "").strip().lower()


_EXCLUDED_TASKS_NORMALIZED = {normalize(n) for n in EXCLUDED_TASK_NAMES}
_EXCLUDED_FIELDS_NORMALIZED = {normalize(k): v for k, v in EXCLUDED_FIELDS.items()}
_REQUEST_TASK_NAMES_NORMALIZED = {normalize(n) for n in REQUEST_TASK_NAMES}


def is_request_task(task):
    return (task["id"] in REQUEST_TASK_IDS
            or normalize(task.get("name")) in _REQUEST_TASK_NAMES_NORMALIZED)

# =========================================================
# SMAPPLY API HELPERS
# =========================================================

def get_session():
    token = wmill.get_variable(SMAPPLY_TOKEN_VAR)
    session = requests.Session()
    # Content-Type is set per request: requests adds JSON or multipart headers itself.
    session.headers.update({"Authorization": f"Bearer {token}"})
    # No "apply-automations" header is sent, so task automations do not fire.
    return session


def summarize_error(resp):
    """Describe a failed response by status code and field keys only, never values."""
    try:
        body = resp.json()
    except ValueError:
        return str(resp.status_code)
    if isinstance(body, dict):
        keys = []
        for k, v in body.items():
            if isinstance(v, dict):
                keys.extend(f"{k}.{sub}" for sub in v.keys())
            else:
                keys.append(k)
        return f"{resp.status_code} (keys: {', '.join(keys[:20])})"
    return str(resp.status_code)


def request(session, method, url, **kwargs):
    """Send a request with simple retry on rate limiting and server errors."""
    if url.startswith("/"):
        url = SMAPPLY_BASE + url
    for attempt in range(5):
        resp = session.request(method, url, timeout=120, **kwargs)
        if resp.status_code == 429 or resp.status_code >= 500:
            wait = int(resp.headers.get("Retry-After", 2 ** attempt))
            print(f"  {resp.status_code} received, retrying in {wait}s")
            time.sleep(wait)
            continue
        if not resp.ok:
            raise RuntimeError(f"{method} failed: {summarize_error(resp)}")
        return resp
    raise RuntimeError(f"{method} failed after retries")


def request_json(session, method, url, **kwargs):
    resp = request(session, method, url, **kwargs)
    return resp.json() if resp.content else None


def get_all_tasks(session, application_id):
    """Fetch every task for an application, following pagination."""
    tasks, url = [], f"/applications/{application_id}/tasks/"
    while url:
        page = request_json(session, "GET", url)
        tasks.extend(page["results"])
        url = page.get("next")  # absolute URL, used as-is
    return tasks

# =========================================================
# FILE COPY HELPERS
# =========================================================

def download_file(session, file_url):
    """Download a file's content. File URLs from the API start with /api/..."""
    url = file_url if file_url.startswith("http") else SMAPPLY_HOST + file_url
    resp = request(session, "GET", url)
    if "application/json" in resp.headers.get("content-type", ""):
        raise RuntimeError("download returned JSON metadata instead of file content")
    return resp.content


def upload_file(session, application_id, task_id, field_key, filename, content):
    """Upload one file to a form question; the question ID is the multipart field name."""
    content_type = mimetypes.guess_type(filename)[0] or "application/octet-stream"
    request(
        session, "POST", f"/applications/{application_id}/tasks/{task_id}/files/",
        files={field_key: (filename, content, content_type)},
    )


def copy_task_files(session, new_id, task, file_fields):
    """Copy every file in the given file fields. Returns (copied, errors) summaries."""
    copied, errors = [], []
    for field_key, files in file_fields.items():
        for position, file_info in enumerate(files, start=1):
            ref = f"task {task['id']}, field {field_key}, file {position}"
            try:
                content = download_file(session, file_info["url"])
                upload_file(session, new_id, task["id"], field_key,
                            file_info.get("filename") or f"{field_key}_{position}", content)
                copied.append(ref)
            except (RuntimeError, KeyError) as err:
                errors.append(f"{ref}: {err}")
    return copied, errors

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


def is_file_list(entry):
    return isinstance(entry, list) and all(isinstance(f, dict) and "url" in f for f in entry)


def build_task_copy(task):
    """Split GET task data into a PATCH data object and a dict of file fields."""
    excluded = _EXCLUDED_FIELDS_NORMALIZED.get(normalize(task["name"]), set())
    data, file_fields, unknown = {}, {}, []
    for key, entry in (task.get("data") or {}).items():
        if key in excluded:
            continue
        # File fields come back as a raw list of file objects, not {label, response}.
        if is_file_list(entry):
            if entry:
                file_fields[key] = entry
            continue
        if not isinstance(entry, dict) or "response" not in entry:
            unknown.append(key)
            continue
        value = entry["response"]
        if is_empty(value) or is_signature(value):
            continue
        data[key] = value
    return data, file_fields, unknown

# =========================================================
# MAIN
# =========================================================

def main(source_id: int, dry_run: bool = False, require_request: bool = True):
    session = get_session()
    source = request_json(session, "GET", f"/applications/{source_id}/")

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

    request_tasks = [t for t in source_tasks if is_request_task(t)]
    request_task = next((t for t in request_tasks if t.get("status") == COMPLETED_STATUS),
                        request_tasks[0] if request_tasks else None)
    request_completed = bool(request_task) and request_task.get("status") == COMPLETED_STATUS
    if require_request and not request_completed:
        reason = "not found" if request_task is None else f"status {request_task.get('status')}"
        raise RuntimeError(
            f"Application {source_id} has no completed 'Place a New Order' request "
            f"({reason}). Nothing was created. "
            "Run with require_request=False to override."
        )

    if dry_run:
        new_id, target_task_ids = None, {t["id"] for t in source_tasks}
        print("DRY RUN: no application will be created")
    else:
        new_id = request_json(session, "POST", "/applications/", json=payload)["id"]
        target_task_ids = {t["id"] for t in get_all_tasks(session, new_id)}
        print(f"Created application {new_id}")

    copied, skipped, errors = [], [], []
    files_copied, file_errors, unrecognized_fields = [], [], []

    for task in source_tasks:
        label = f"[{task['id']}] {task['name']}"
        if task["type"] not in COPYABLE_TASK_TYPES:
            skipped.append(f"{label}: task type {task['type']}")
            continue
        if (normalize(task["name"]) in _EXCLUDED_TASKS_NORMALIZED
                or is_request_task(task)):
            skipped.append(f"{label}: excluded")
            continue
        if task["id"] not in target_task_ids:
            errors.append(f"{label}: not present on new application")
            continue

        data, file_fields, unknown = build_task_copy(task)
        unrecognized_fields += [f"task {task['id']}, field {key}" for key in unknown]
        if not COPY_FILES:
            file_fields = {}
        file_count = sum(len(f) for f in file_fields.values())

        if not data and not file_fields:
            skipped.append(f"{label}: no answered fields")
            continue

        if not dry_run:
            if data:
                try:
                    request(session, "PATCH", f"/applications/{new_id}/tasks/{task['id']}/",
                            json={"status": COPIED_TASK_STATUS, "data": data})
                except RuntimeError as err:
                    errors.append(f"{label}: {err}")
                    continue
            if file_fields:
                ok, failed = copy_task_files(session, new_id, task, file_fields)
                files_copied += ok
                file_errors += failed
        else:
            files_copied += [f"task {task['id']}, field {k}: {len(v)} file(s)"
                             for k, v in file_fields.items()]

        summary = f"{label}: {len(data)} fields"
        if file_count:
            summary += f", {file_count} file(s)"
        copied.append(summary)

    for line in copied:
        print(f"  copy {line}")
    for line in skipped:
        print(f"  skip {line}")
    for line in errors + file_errors:
        print(f"  ERROR {line}")

    # Returned as the Windmill job result (IDs and counts only, no response values).
    return {
        "source_application_id": source_id,
        "new_application_id": new_id,
        "dry_run": dry_run,
        "request_completed": request_completed,
        "copied_tasks": copied,
        "skipped_tasks": skipped,
        "files_copied": files_copied,
        "errors": errors,
        "file_errors": file_errors,
        "unrecognized_fields": unrecognized_fields,
    }