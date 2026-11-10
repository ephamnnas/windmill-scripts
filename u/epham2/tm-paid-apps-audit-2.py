#requirements:
#wmill
#requests

import json
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
import wmill

# =========================================================
# NNAS TM PAYMENT AUDIT (v5)
# Paid expedited applications in program 295185, split into:
#   - currently submitted / in process
#   - pushed back (submitted at least once, now returned to applicant)
#   - truly unsubmitted (never submitted)
#
# Confirmed by probe (app 35542766):
#   - payment tasks: status INT, 3 = completed; one of five reaches 3.
#   - application has last_submitted_at (platform-set), state (int),
#     status {id, name} (custom status), current_stage {id, title}.
#   - "Application Sign Off" task completes on submission.
#
# Classification: sign-off task incomplete = never submitted; sign-off
# complete but stage in PUSHED_BACK_STAGES = pushed back; else in process.
# The breakdown table is still returned for auditing.
#
# Privacy: output is application IDs plus workflow values (state, status,
# stage). Applicant fields and staff notes are never printed or returned.
# =========================================================

SMAPPLY_TOKEN_VAR = "u/epham2/SMAPPLY_ACCESS_TOKEN"
SMAPPLY_BASE = "https://nnas.smapply.io/api"

PAYMENT_TASK_MARKER = "expedited service payment"  # matches all five tasks
SIGNOFF_TASK_NAME = "application sign off"
STATUS_COMPLETED = 3
FIELD_FIRST_SUBMISSION = "153648"

# Calibrated from the v3 breakdown run:
#   - Sign-off task complete = submitted at least once (it stays complete
#     after a rejection; last_submitted_at / 153648 miss some issued apps).
#   - Pushed back = signed off but current stage is the applicant stage.
PUSHED_BACK_STAGES = {"Application Stage"}
REFUNDED_STATUS_WORDS = ("refunded",)

# Allowlist for the probe: only workflow fields, nothing applicant-level.
PROBE_SAFE_KEYS = ("id", "state", "status", "current_stage", "decision",
                   "labels", "created_at", "updated_at", "last_submitted_at")

_NOT_PRESENT = object()


# ---------------------------------------------------------
# SMAPPLY HTTP
# ---------------------------------------------------------

_TOKEN = None
_local = threading.local()


def _session():
    """One keep-alive session per worker thread, token fetched once."""
    global _TOKEN
    if _TOKEN is None:
        _TOKEN = wmill.get_variable(SMAPPLY_TOKEN_VAR)
    s = getattr(_local, "s", None)
    if s is None:
        s = requests.Session()
        s.headers.update({"Authorization": "Bearer %s" % _TOKEN,
                          "Content-Type": "application/json"})
        _local.s = s
    return s


def _get(url, params=None, retries=6):
    for _ in range(retries):
        r = _session().get(url, params=params, timeout=60)
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
# EXTRACTION
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


def task_name(t):
    return str(t.get("name") or t.get("title") or "")


def scan_tasks(tasks):
    """(paid payment task name or None, sign-off task status or None)."""
    paid, signoff = None, None
    for t in tasks:
        n = task_name(t).lower()
        if PAYMENT_TASK_MARKER in n and t.get("status") == STATUS_COMPLETED:
            paid = task_name(t)
        elif n == SIGNOFF_TASK_NAME:
            signoff = t.get("status")
    return paid, signoff


def status_name(item):
    s = item.get("status")
    return (s.get("name") if isinstance(s, dict) else s) or "(none)"


def stage_title(item):
    s = item.get("current_stage")
    return (s.get("title") if isinstance(s, dict) else s) or "(none)"


def ever_submitted(item):
    if item.get("last_submitted_at"):
        return True
    v = find_cf(item.get("custom_fields"), FIELD_FIRST_SUBMISSION)
    return v not in (None, "", _NOT_PRESENT)


def is_refunded(row):
    return any(w in row["status"].lower() for w in REFUNDED_STATUS_WORDS)


# ---------------------------------------------------------
# PROBE (allowlist only)
# ---------------------------------------------------------

def probe(program_id, probe_app_id):
    item = next((a for a in iter_applications(program_id)
                 if a.get("id") == probe_app_id), None)
    if item is None:
        return "app %s not found in program %s" % (probe_app_id, program_id)
    print("=== application (workflow fields only) ===")
    print(json.dumps({k: item.get(k) for k in PROBE_SAFE_KEYS},
                     indent=2, default=str))
    print("\n153648 (first submission): %r"
          % find_cf(item.get("custom_fields"), FIELD_FIRST_SUBMISSION))
    print("\n=== tasks ===")
    for t in get_tasks(probe_app_id):
        print(json.dumps({k: t.get(k) for k in
                          ("id", "name", "status", "completed_at")}, default=str))
    return "probe complete"


# ---------------------------------------------------------
# MAIN
# ---------------------------------------------------------

def _check(item):
    paid, signoff = scan_tasks(get_tasks(item.get("id")))
    return item, paid, signoff


def main(program_id: int = 295185,
         probe_app_id: int = 0,
         workers: int = 8,
         ui_submitted_refs: list = None):
    """
    ui_submitted_refs: optional list of reference IDs (e.g. "NNAS-T-0000000445")
        from the SMApply "Submitted Applications" export. When given, the
        result includes a reconciliation against this audit's groups.
    """
    if probe_app_id:
        return probe(program_id, probe_app_id)

    t0 = time.time()
    raw = list(iter_applications(program_id))
    items = list({it.get("id"): it for it in raw}.values())
    print("Listed %d applications (%d unique) in %.0fs"
          % (len(raw), len(items), time.time() - t0))

    paid_rows = []
    all_info = {}
    by_task = {}
    done = 0
    with ThreadPoolExecutor(max_workers=max(1, workers)) as pool:
        futures = [pool.submit(_check, it) for it in items]
        for fut in as_completed(futures):
            item, paid, signoff = fut.result()
            done += 1
            if done % 250 == 0:
                print("  ...%d / %d (%.0fs)" % (done, len(items), time.time() - t0))
            all_info[item.get("id")] = {
                "ref": item.get("reference_id"), "paid": bool(paid),
                "signoff": signoff, "stage": stage_title(item),
                "status": status_name(item), "state": item.get("state")}
            if not paid:
                continue
            by_task[paid] = by_task.get(paid, 0) + 1
            paid_rows.append({
                "id": item.get("id"),
                "state": item.get("state"),
                "status": status_name(item),
                "stage": stage_title(item),
                "signoff": signoff,
                "ever_submitted": ever_submitted(item),
                "item": item,
            })
    print("Finished in %.0fs" % (time.time() - t0))

    # Breakdown of real values among paid applications, for calibration.
    breakdown = {}
    for r in paid_rows:
        key = "state=%s | status=%s | signoff=%s | ever_submitted=%s" % (
            r["state"], r["status"], r["signoff"], r["ever_submitted"])
        entry = breakdown.setdefault(key, {"count": 0, "example_ids": []})
        entry["count"] += 1
        if len(entry["example_ids"]) < 3:
            entry["example_ids"].append(r["id"])
    breakdown = dict(sorted(breakdown.items(), key=lambda kv: -kv[1]["count"]))

    stages = {}
    for r in paid_rows:
        stages[r["stage"]] = stages.get(r["stage"], 0) + 1

    print("\n--- TM program %s ---" % program_id)
    print("  Applications seen : %d" % len(items))
    print("  Payment task paid : %d" % len(paid_rows))
    for name, n in sorted(by_task.items()):
        print("      %-40s %d" % (name, n))
    print("\n  Breakdown of paid applications:")
    for key, e in breakdown.items():
        print("    %4d  %s   e.g. %s" % (e["count"], key, e["example_ids"]))
    print("\n  Current stage of paid applications:")
    for st, n in sorted(stages.items(), key=lambda kv: -kv[1]):
        print("    %4d  %s" % (n, st))

    result = {"program_id": program_id,
              "seen": len(items),
              "paid": len(paid_rows),
              "paid_by_task": by_task,
              "breakdown": breakdown,
              "stages": stages}

    pushed_back, never_submitted, in_process, refunded = [], [], [], []
    pb_status = {}
    for r in paid_rows:
        if is_refunded(r):
            refunded.append(r["id"])
        if r["signoff"] != STATUS_COMPLETED:
            never_submitted.append(r["id"])
        elif r["stage"] in PUSHED_BACK_STAGES:
            pushed_back.append(r["id"])
            pb_status[r["status"]] = pb_status.get(r["status"], 0) + 1
        else:
            in_process.append(r["id"])

    print("\n  Paid, submitted / in process  : %d" % len(in_process))
    print("  Paid, pushed back             : %d" % len(pushed_back))
    for st, n in sorted(pb_status.items(), key=lambda kv: -kv[1]):
        print("      status %-35s %d" % (st, n))
    print("  Paid, never submitted         : %d" % len(never_submitted))
    if refunded:
        print("  (of all paid, refunded: %d  %s)" % (len(refunded), sorted(refunded)))

    result.update({"counts": {"in_process": len(in_process),
                              "pushed_back": len(pushed_back),
                              "never_submitted": len(never_submitted),
                              "refunded": len(refunded)},
                   "pushed_back_by_status": pb_status,
                   "pushed_back_ids": sorted(pushed_back),
                   "never_submitted_ids": sorted(never_submitted),
                   "refunded_ids": sorted(refunded)})

    if ui_submitted_refs:
        result["reconciliation"] = reconcile(
            ui_submitted_refs, all_info,
            set(in_process), set(pushed_back), set(never_submitted))
    return result


def reconcile(ui_refs, all_info, in_process, pushed_back, never_submitted):
    """Compare the UI's submitted list (reference IDs) with audit groups.
    Returns workflow values only, no applicant data."""
    ui = {str(r).strip() for r in ui_refs if str(r).strip()}
    by_ref = {v["ref"]: k for k, v in all_info.items() if v.get("ref")}

    def group(appid):
        if appid in in_process:
            return "in_process"
        if appid in pushed_back:
            return "pushed_back"
        if appid in never_submitted:
            return "never_submitted"
        return "unpaid"

    def describe(appid):
        v = all_info[appid]
        return {"id": appid, "ref": v["ref"], "stage": v["stage"],
                "status": v["status"], "state": v["state"],
                "signoff": v["signoff"], "paid": v["paid"]}

    ui_by_group, not_found = {}, []
    for ref in sorted(ui):
        appid = by_ref.get(ref)
        if appid is None:
            not_found.append(ref)
            continue
        ui_by_group.setdefault(group(appid), []).append(appid)

    ui_ids = {by_ref[r] for r in ui if r in by_ref}
    in_process_not_ui = sorted(in_process - ui_ids)

    summary = {
        "ui_submitted": len(ui),
        "ui_matched_to_api": len(ui_ids),
        "ui_not_found_in_api": len(not_found),
        "ui_by_audit_group": {g: len(v) for g, v in ui_by_group.items()},
        "audit_in_process_not_in_ui": len(in_process_not_ui),
    }
    print("\n--- Reconciliation with UI export ---")
    for k, v in summary.items():
        print("  %-28s %s" % (k, v))

    return {
        "summary": summary,
        "ui_not_found_in_api": not_found,
        # UI says submitted, audit disagrees: show workflow detail
        "ui_submitted_but_audit_pushed_back": [describe(i) for i in ui_by_group.get("pushed_back", [])],
        "ui_submitted_but_audit_never_submitted": [describe(i) for i in ui_by_group.get("never_submitted", [])],
        "ui_submitted_but_audit_unpaid": [describe(i) for i in ui_by_group.get("unpaid", [])],
        # Audit says in process, UI does not list as submitted
        "audit_in_process_not_in_ui": [describe(i) for i in in_process_not_ui],
    }