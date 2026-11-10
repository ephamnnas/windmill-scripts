#requirements:
#wmill
#requests

import json
import time

import requests
import wmill

# =========================================================
# NNAS TM PAYMENT AUDIT
# Paid expedited applications in program 295185, split into:
#   - currently submitted
#   - pushed back (submitted at least once, now not submitted)
#   - truly unsubmitted (never submitted)
#
# Reuses the auth, pagination and confirmed shapes from the revenue
# report generator (v7):
#   - payment tasks: status INT, 3 = completed; only the applicable one
#     of the five reaches 3.
#   - custom field 153648 = first submission date. Populated means the
#     application was submitted at least once.
#
# Output is application IDs only. No names or emails are printed or
# returned.
# =========================================================

SMAPPLY_TOKEN_VAR = "u/epham2/SMAPPLY_ACCESS_TOKEN"
SMAPPLY_BASE = "https://nnas.smapply.io/api"

PAYMENT_TASK_MARKER = "expedited service payment"  # matches all five tasks
STATUS_COMPLETED = 3
FIELD_FIRST_SUBMISSION = "153648"

# Current submission state on the application object. Confirm with
# probe_app_id before trusting the pushed back count; adjust if needed.
APP_SUBMITTED_BOOL_FIELDS = ("submitted", "is_submitted")
APP_STATUS_FIELDS = ("status", "state")
APP_SUBMITTED_STATUS_WORDS = ("submitted", "complete")

PII_KEYS = {"applicant", "applicant_first_name", "applicant_last_name",
            "applicant_email", "first_name", "last_name", "email", "name",
            "title", "owner", "phone", "address", "collaborators"}

_NOT_PRESENT = object()


# ---------------------------------------------------------
# SMAPPLY HTTP (same pattern as the revenue report)
# ---------------------------------------------------------

_TOKEN = None


def _sm_headers():
    global _TOKEN
    if _TOKEN is None:
        _TOKEN = wmill.get_variable(SMAPPLY_TOKEN_VAR)
    return {"Authorization": "Bearer %s" % _TOKEN,
            "Content-Type": "application/json"}


def _get(url, params=None, retries=3):
    for _ in range(retries):
        r = requests.get(url, headers=_sm_headers(), params=params, timeout=60)
        if r.status_code == 429:
            wait = int(r.headers.get("Retry-After", "5"))
            print("  429 rate limited, sleeping %ss" % wait)
            time.sleep(wait)
            continue
        r.raise_for_status()
        return r.json()
    raise RuntimeError("giving up after retries: %s" % url)


def _paginate(url, params=None):
    page = 1
    while True:
        body = _get(url, params={**(params or {}), "page": page})
        if isinstance(body, list):
            yield from body
            return
        results = body.get("results", [])
        if not results:
            return
        yield from results
        num_pages = body.get("num_pages")
        if num_pages is not None:
            if page >= num_pages:
                return
        elif body.get("next") in (None, ""):
            return
        page += 1


def iter_applications(program_id):
    return _paginate("%s/applications/" % SMAPPLY_BASE, {"program": program_id})


def get_tasks(appid):
    return list(_paginate("%s/applications/%s/tasks/" % (SMAPPLY_BASE, appid)))


# ---------------------------------------------------------
# CLASSIFICATION
# ---------------------------------------------------------

def find_cf(custom_fields, field_id):
    target = str(field_id)
    if isinstance(custom_fields, list):
        for cf in custom_fields:
            if str(cf.get("id")) == target:
                return cf.get("value")
    elif isinstance(custom_fields, dict):
        if target in custom_fields:
            v = custom_fields[target]
            return v.get("value") if isinstance(v, dict) else v
        for v in custom_fields.values():
            if isinstance(v, dict) and str(v.get("id")) == target:
                return v.get("value")
    return _NOT_PRESENT


def paid_payment_task(tasks):
    """Name of the completed payment task, or None."""
    for t in tasks:
        name = str(t.get("name") or t.get("title") or "")
        if PAYMENT_TASK_MARKER in name.lower() and t.get("status") == STATUS_COMPLETED:
            return name
    return None


def ever_submitted(item):
    v = find_cf(item.get("custom_fields"), FIELD_FIRST_SUBMISSION)
    return v not in (None, "", _NOT_PRESENT)


def currently_submitted(item):
    for f in APP_SUBMITTED_BOOL_FIELDS:
        if isinstance(item.get(f), bool):
            return item[f]
    for f in APP_STATUS_FIELDS:
        v = item.get(f)
        if isinstance(v, dict):
            v = v.get("name") or v.get("label")
        if isinstance(v, str) and v:
            s = v.lower()
            return any(w in s for w in APP_SUBMITTED_STATUS_WORDS)
    return None  # unknown; reported separately rather than guessed


# ---------------------------------------------------------
# PROBE (masked)
# ---------------------------------------------------------

def _mask(obj):
    if isinstance(obj, dict):
        return {k: ("***" if k.lower() in PII_KEYS else _mask(v))
                for k, v in obj.items()}
    if isinstance(obj, list):
        return [_mask(i) for i in obj]
    return obj


def probe(program_id, probe_app_id):
    item = next((a for a in iter_applications(program_id)
                 if a.get("id") == probe_app_id), None)
    if item is None:
        return "app %s not found in program %s" % (probe_app_id, program_id)
    print("=== application (masked, custom_fields omitted) ===")
    print(json.dumps(_mask({k: v for k, v in item.items()
                            if k != "custom_fields"}), indent=2, default=str))
    print("\n153648 (first submission): %r"
          % find_cf(item.get("custom_fields"), FIELD_FIRST_SUBMISSION))
    print("currently_submitted() -> %r" % currently_submitted(item))
    print("\n=== tasks ===")
    for t in get_tasks(probe_app_id):
        print(json.dumps({k: t.get(k) for k in
                          ("id", "name", "status", "completed_at")}, default=str))
    return "probe complete"


# ---------------------------------------------------------
# MAIN
# ---------------------------------------------------------

def main(program_id: int = 295185,
         probe_app_id: int = 0,
         task_fetch_sleep: float = 0.0):
    if probe_app_id:
        return probe(program_id, probe_app_id)

    seen = 0
    submitted, pushed_back, never_submitted, unknown = [], [], [], []
    by_task = {}

    for item in iter_applications(program_id):
        seen += 1
        appid = item.get("id")
        tasks = get_tasks(appid)
        if task_fetch_sleep:
            time.sleep(task_fetch_sleep)

        task_name = paid_payment_task(tasks)
        if not task_name:
            continue
        by_task[task_name] = by_task.get(task_name, 0) + 1

        now = currently_submitted(item)
        if now is True:
            submitted.append(appid)
        elif now is None:
            unknown.append(appid)
        elif ever_submitted(item):
            pushed_back.append(appid)
        else:
            never_submitted.append(appid)

        if seen % 100 == 0:
            print("  ...%d applications processed" % seen)

    paid_total = len(submitted) + len(pushed_back) + len(never_submitted) + len(unknown)

    print("\n--- TM program %s ---" % program_id)
    print("  Applications seen               : %d" % seen)
    print("  Payment task paid               : %d" % paid_total)
    for name, n in sorted(by_task.items()):
        print("      %-40s %d" % (name, n))
    print("  Paid and currently submitted    : %d" % len(submitted))
    print("  Paid, not submitted, pushed back: %d" % len(pushed_back))
    print("  Paid, never submitted           : %d" % len(never_submitted))
    if unknown:
        print("  Paid, submission state unknown  : %d  (run probe_app_id on "
              "one of these and adjust APP_* fields)" % len(unknown))

    return {
        "program_id": program_id,
        "counts": {"seen": seen, "paid": paid_total,
                   "submitted": len(submitted),
                   "pushed_back": len(pushed_back),
                   "never_submitted": len(never_submitted),
                   "unknown": len(unknown)},
        "paid_by_task": by_task,
        "pushed_back_ids": pushed_back,
        "never_submitted_ids": never_submitted,
        "unknown_ids": unknown,
    }