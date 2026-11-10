# requirements:
# requests
# wmill
# psycopg2-binary
#
# SMApply -> Postgres sync (incremental)
# --------------------------------------
# Lists ALL applications every run (cheap: ~46 pages), but only re-fetches
# tasks for applications whose updated_at moved past the stored cursor.
# Per changed application, tasks + task_fields are replaced atomically
# (delete + insert in one transaction; FK cascade cleans task_fields).
# Cursor is kept in Windmill script state via wmill.get_state/set_state.
#
# Run with full_resync=true for the first load (or to rebuild from scratch).

import json
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock

import psycopg2
import psycopg2.extras
import requests
import wmill

BASE_URL = "https://portal.nnas.ca/api"
SMAPPLY_TOKEN_PATH = "u/epham2/SMAPPLY_ACCESS_TOKEN"  # known-good from intake script
PG_RESOURCE_PATH = "u/knguyen/smapply_pgwmill.get_variable"          # <- adjust to your resource

# Filter applied to the applications list endpoint. NOTE: applications that
# leave this filter disappear from the list; the sync marks them listed=false
# rather than deleting, so historical data stays queryable.
LIST_PARAMS = {"state": 0}

TASK_TYPE_LABELS = {1: "form", 2: "file_storage", 5: "confirmation", 10: "payment"}
STATUS_LABELS = {1: "not_started", 2: "in_progress", 3: "completed"}

print_lock = Lock()


# ---------------------------------------------------------------- HTTP layer
def api_get(session, url, params=None, max_retries=6):
    for attempt in range(max_retries):
        resp = session.get(url, params=params, timeout=90)
        if resp.status_code == 429:
            wait = resp.headers.get("Retry-After")
            wait = int(wait) if wait and wait.isdigit() else min(2 ** attempt, 30)
            time.sleep(wait)
            continue
        if resp.status_code >= 400:
            raise RuntimeError(
                f"GET {resp.url} failed with {resp.status_code}: {resp.text[:2000]}")
        return resp.json()
    raise RuntimeError(f"GET {url} still rate-limited after {max_retries} retries")


def paged_get(session, url, per_page=25, extra_params=None):
    results, page = [], 1
    while True:
        params = {"page": page, "per_page": per_page}
        if extra_params:
            params.update(extra_params)
        payload = api_get(session, url, params=params)
        results.extend(payload.get("results", []))
        if page >= payload.get("num_pages", 1) or not payload.get("next"):
            return results
        page += 1


# ------------------------------------------------------------------ helpers
def ts(value):
    """Empty string / None -> None so Postgres gets a real NULL."""
    return value or None


def flatten_application(app):
    applicant = app.get("applicant") or {}
    stage = app.get("current_stage") or {}
    status = app.get("status") or {}
    program = app.get("program") or {}
    cf = {(c.get("name") or f"field_{c.get('id')}").strip(): c.get("value")
          for c in app.get("custom_fields") or []}
    return (
        app.get("id"), app.get("reference_id"), app.get("title"),
        applicant.get("id"), applicant.get("first_name"),
        applicant.get("last_name"), applicant.get("email"),
        program.get("id"), program.get("name"),
        stage.get("id"), stage.get("title"),
        status.get("name"), app.get("state"),
        "; ".join(l.get("name", "") for l in (app.get("labels") or [])),
        app.get("average_score"), app.get("overall_score"),
        app.get("weighted_score"), json.dumps(cf, ensure_ascii=False),
        ts(app.get("created_at")), ts(app.get("updated_at")),
        ts(app.get("last_submitted_at")),
    )


def flatten_task(app_id, task):
    stage = task.get("stage") or {}
    return (
        app_id, task.get("id"), task.get("name"),
        task.get("type"), TASK_TYPE_LABELS.get(task.get("type"), ""),
        task.get("status"), STATUS_LABELS.get(task.get("status"), ""),
        ts(task.get("created_at")), ts(task.get("updated_at")),
        ts(task.get("completed_at")),
        stage.get("id"), stage.get("title"), stage.get("type"),
    )


def explode_task_fields(app, task):
    rows = []
    base = (app.get("id"), task.get("id"), app.get("reference_id"),
            task.get("name"), task.get("status"), ts(task.get("completed_at")),
            (task.get("stage") or {}).get("title"))
    for key, value in (task.get("data") or {}).items():
        # files
        if isinstance(value, list) and value and isinstance(value[0], dict) \
                and "filename" in value[0]:
            for f in value:
                rows.append(base + (key, "", "file", f.get("filename"), "",
                                    json.dumps(f, ensure_ascii=False)))
            continue
        # form field with label/response
        if isinstance(value, dict) and "label" in value:
            label = (value.get("label") or "").strip()
            resp = value.get("response")
            if isinstance(resp, dict):
                rows.append(base + (key, label, "choice_with_other",
                                    "" if resp.get("value") is None else str(resp.get("value")),
                                    resp.get("other") or "",
                                    json.dumps(resp, ensure_ascii=False)))
            elif isinstance(resp, list):
                rows.append(base + (key, label, "multi_select",
                                    "; ".join(str(v) for v in resp), "",
                                    json.dumps(resp, ensure_ascii=False)))
            else:
                rows.append(base + (key, label, "scalar",
                                    "" if resp is None else str(resp), "", None))
            continue
        # flat payment / confirmation blobs
        if isinstance(value, (dict, list)):
            rows.append(base + (key, "", "raw", "", "",
                                json.dumps(value, ensure_ascii=False, default=str)))
        else:
            rows.append(base + (key, "", "raw",
                                "" if value is None else str(value), "", None))
    return rows


APP_UPSERT = """
INSERT INTO applications (application_id, reference_id, title, applicant_id,
    applicant_first_name, applicant_last_name, applicant_email, program_id,
    program_name, current_stage_id, current_stage_title, status_name, state,
    labels, average_score, overall_score, weighted_score, custom_fields,
    created_at, updated_at, last_submitted_at, synced_at)
VALUES %s
ON CONFLICT (application_id) DO UPDATE SET
    reference_id = EXCLUDED.reference_id, title = EXCLUDED.title,
    applicant_id = EXCLUDED.applicant_id,
    applicant_first_name = EXCLUDED.applicant_first_name,
    applicant_last_name = EXCLUDED.applicant_last_name,
    applicant_email = EXCLUDED.applicant_email,
    program_id = EXCLUDED.program_id, program_name = EXCLUDED.program_name,
    current_stage_id = EXCLUDED.current_stage_id,
    current_stage_title = EXCLUDED.current_stage_title,
    status_name = EXCLUDED.status_name, state = EXCLUDED.state,
    labels = EXCLUDED.labels, average_score = EXCLUDED.average_score,
    overall_score = EXCLUDED.overall_score,
    weighted_score = EXCLUDED.weighted_score,
    custom_fields = EXCLUDED.custom_fields, created_at = EXCLUDED.created_at,
    updated_at = EXCLUDED.updated_at,
    last_submitted_at = EXCLUDED.last_submitted_at, synced_at = now()
"""

TASK_INSERT = """
INSERT INTO tasks (application_id, task_id, task_name, task_type,
    task_type_label, task_status, task_status_label, created_at, updated_at,
    completed_at, stage_id, stage_title, stage_type)
VALUES %s
"""

FIELD_INSERT = """
INSERT INTO task_fields (application_id, task_id, reference_id, task_name,
    task_status, task_completed_at, stage_title, field_key, field_label,
    response_type, response_value, response_other, response_json)
VALUES %s
"""


# --------------------------------------------------------------------- main
def main(
    full_resync: bool = False,
    max_workers: int = 8,
    per_page: int = 25,
    app_limit: int = 0,   # 0 = all; set e.g. 10 for a test run
):
    full_resync = str(full_resync).lower() in ("true", "1", "yes")
    max_workers = int(max_workers)
    per_page = int(per_page)
    app_limit = int(app_limit)

    token = wmill.get_variable(SMAPPLY_TOKEN_PATH)
    session = requests.Session()
    session.headers.update({"Authorization": f"Bearer {token}",
                            "Accept": "application/json"})

    pg = wmill.get_resource(PG_RESOURCE_PATH)
    conn = psycopg2.connect(
        host=pg["host"], port=pg.get("port", 5432), dbname=pg["dbname"],
        user=pg["user"], password=pg["password"],
        sslmode=pg.get("sslmode", "prefer"))
    conn.autocommit = False

    state = wmill.get_state() or {}
    cursor_ts = None if full_resync else state.get("last_updated_at")

    # 1. Full application list every run (change detection needs it anyway)
    print(f"Fetching application list (filter: {LIST_PARAMS})...")
    apps = paged_get(session, f"{BASE_URL}/applications/", per_page,
                     extra_params=LIST_PARAMS)
    if app_limit:
        apps = apps[:app_limit]
    max_updated = max((a.get("updated_at") or "" for a in apps), default="")

    changed = [a for a in apps
               if cursor_ts is None or (a.get("updated_at") or "") >= cursor_ts]
    print(f"  {len(apps)} applications total, {len(changed)} changed "
          f"(cursor: {cursor_ts or 'none - full load'})")

    # 2. Upsert application rows (all of them - cheap, keeps stage/status fresh)
    with conn.cursor() as cur:
        psycopg2.extras.execute_values(
            cur, APP_UPSERT,
            [flatten_application(a) + ("now()",) for a in apps],
            template="(" + ",".join(["%s"] * 21) + ", now())")
        # Applications no longer returned by the filtered list (e.g. they left
        # state=0) stay in the DB for history but get flagged so queries can
        # distinguish live rows from ones the sync no longer refreshes.
        listed_ids = [a["id"] for a in apps]
        cur.execute("UPDATE applications SET listed = true "
                    "WHERE application_id = ANY(%s)", (listed_ids,))
        cur.execute("UPDATE applications SET listed = false "
                    "WHERE NOT (application_id = ANY(%s))", (listed_ids,))
    conn.commit()

    # 3. Re-fetch + replace tasks for changed applications only
    def fetch(app):
        return app, paged_get(
            session, f"{BASE_URL}/applications/{app['id']}/tasks/", per_page)

    errors, done = [], 0
    with ThreadPoolExecutor(max_workers=max_workers) as pool:
        futures = {pool.submit(fetch, a): a["id"] for a in changed}
        for fut in as_completed(futures):
            app_id = futures[fut]
            try:
                app, tasks = fut.result()
                task_rows = [flatten_task(app_id, t) for t in tasks]
                field_rows = []
                for t in tasks:
                    field_rows.extend(explode_task_fields(app, t))
                # DB writes serialized on the shared connection via a
                # per-application transaction (replace pattern; FK cascades
                # clear old task_fields)
                with print_lock:
                    with conn.cursor() as cur:
                        cur.execute(
                            "DELETE FROM tasks WHERE application_id = %s",
                            (app_id,))
                        if task_rows:
                            psycopg2.extras.execute_values(
                                cur, TASK_INSERT, task_rows)
                        if field_rows:
                            psycopg2.extras.execute_values(
                                cur, FIELD_INSERT, field_rows, page_size=500)
                    conn.commit()
            except Exception as exc:
                conn.rollback()
                errors.append({"application_id": app_id, "error": str(exc)})
                print(f"[error] app {app_id}: {exc}")
            done += 1
            if done % 100 == 0:
                print(f"  {done}/{len(changed)} synced")

    conn.close()

    # 4. Advance cursor only if nothing failed (failed apps retry next run)
    if not errors and max_updated:
        wmill.set_state({"last_updated_at": max_updated})

    summary = {"applications_listed": len(apps), "applications_synced": len(changed),
               "failed": errors, "cursor_advanced_to": max_updated if not errors else None}
    print(json.dumps(summary, indent=2, default=str))
    return summary