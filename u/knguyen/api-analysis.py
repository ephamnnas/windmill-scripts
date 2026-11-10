# requirements:
# requests
# wmill

"""
ACCOUNTS vs APPLICATIONS — the true top of the funnel.

Pulls every SMApply user account and every application (ALL programs, ALL
states by default, so an account with an application anywhere is never
mis-flagged), then reports accounts that have no application at all, by the
month the account was created.

Why all programs: the pipeline report scopes to T and C, but an account that
applied under a legacy or other program is not "never applied". Program
scoping is applied to the *comparison* output, not to the match.

Outputs (SharePoint Applications/zDATA/volume_report/ + inline summary):
  accounts_<ts>.csv          one row per account (anonymised): created month,
                              role fields, has_application, n_applications,
                              programs, first/last application dates
  accounts_monthly_<ts>.csv  month x: accounts created, with application,
                              without, % never applied, applications created
  eligibility_<ts>.csv       one row per account per eligibility question:
                              question id, question text, answer (anonymised
                              account id). Eligibility is answered BEFORE an
                              application exists, so this is the only
                              attribute data available for never-applied
                              accounts.
  eligibility_signal_<ts>.csv question x answer: accounts, how many went on
                              to create an application, % applied (k>=5).
  The run result also prints the raw field names and role values seen on the
  user objects, because the users endpoint schema is not fully documented --
  check that block on the first run to confirm applicants are being
  distinguished from staff/reviewers correctly.
"""

import csv
import hashlib
import hmac
import io
import json
import time
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock
from datetime import datetime, timezone
from urllib.parse import quote

import requests
import wmill

BASE_URL = "https://nnas.smapply.io/api"
SMAPPLY_TOKEN_PATH = "u/epham2/SMAPPLY_ACCESS_TOKEN"

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SP_HOST, SP_SITE, SP_LIBRARY = "nnasca.sharepoint.com", "/sites/NNAS-Portal-Sharepoint", "Applications"
SP_FOLDER = "zDATA/volume_report"

# Keys that may identify user type on the user object; reported, and used for
# the applicant filter when present.
ROLE_KEYS = ("roles", "role", "user_type", "type", "is_staff", "is_admin", "is_reviewer",
             "is_applicant", "groups", "teams")
DATE_KEYS = ("signed_up", "date_joined", "created_at", "created", "joined_at", "invited_at")
ACTIVE_KEYS = ("last_active", "last_login")


# ---------------------------------------------------------------- HTTP
def api_get(s, url, params=None, max_retries=6):
    for attempt in range(max_retries):
        r = s.get(url, params=params, timeout=90)
        if r.status_code == 429:
            w = r.headers.get("Retry-After")
            time.sleep(int(w) if w and w.isdigit() else min(2 ** attempt, 30))
            continue
        if r.status_code == 403:
            raise RuntimeError(f"403 on {r.url}: {r.text[:400]} -- the token needs "
                               f"'View users' (for /users/) and application read permissions.")
        if r.status_code >= 400:
            raise RuntimeError(f"GET {r.url} -> {r.status_code}: {r.text[:800]}")
        return r.json()
    raise RuntimeError(f"GET {url}: rate-limited after {max_retries} retries")


def paged(s, url, extra=None, per_page=50):
    out, page = [], 1
    while True:
        p = api_get(s, url, {"page": page, "per_page": per_page, **(extra or {})})
        if isinstance(p, list):
            return p
        out.extend(p.get("results", []))
        if page >= p.get("num_pages", 1) or not p.get("next"):
            return out
        page += 1


def to_date(v):
    v = (str(v) if v is not None else "").strip()
    return v[:10] if len(v) >= 10 and v[4] == "-" else None


def anon(v, salt):
    return "USR-" + hmac.new(salt.encode(), str(v).encode(), hashlib.sha256).hexdigest()[:12].upper()


def role_summary(u):
    """Flatten whatever role-ish fields exist into one string for reporting."""
    parts = []
    for k in ROLE_KEYS:
        if k in u and u[k] not in (None, "", [], {}):
            v = u[k]
            if isinstance(v, list):
                v = "|".join(str(x.get("name", x) if isinstance(x, dict) else x) for x in v)
            parts.append(f"{k}={v}")
    return "; ".join(parts) or None


def user_created(u):
    for k in DATE_KEYS:
        if u.get(k):
            return to_date(u[k])
    return None


def user_last_active(u):
    for k in ACTIVE_KEYS:
        if u.get(k):
            return to_date(u[k])
    return None



# ---------------------------------------------------------------- eligibility
RESP_LIST_KEYS = ("responses", "Response", "response", "answers", "questions", "results")
ANSWER_KEYS = ("response", "value", "answer", "selected", "choice", "choices", "text")
_elig_lock = Lock()
_elig_schema = {"payload_keys": set(), "item_keys": set(), "answer_key_used": Counter(), "samples": []}


def fetch_eligibility(s, uid):
    """Return parsed eligibility or None if the user has none (404)."""
    for attempt in range(6):
        r = s.get(f"{BASE_URL}/users/{uid}/eligibility/", timeout=90)
        if r.status_code == 429:
            w = r.headers.get("Retry-After")
            time.sleep(int(w) if w and w.isdigit() else min(2 ** attempt, 30))
            continue
        if r.status_code == 404:
            return None
        if r.status_code == 403:
            raise RuntimeError(f"403 on eligibility for user {uid}: token lacks 'View users' permission")
        if r.status_code >= 400:
            raise RuntimeError(f"eligibility {uid} -> {r.status_code}: {r.text[:300]}")
        return r.json()
    raise RuntimeError(f"eligibility {uid}: rate-limited")


def parse_eligibility(payload):
    """-> (header dict, list of (qid, question, answer))"""
    if not isinstance(payload, dict):
        return {"eligibility_present": False}, []
    with _elig_lock:
        _elig_schema["payload_keys"].update(payload.keys())
        if len(_elig_schema["samples"]) < 2:
            _elig_schema["samples"].append(json.dumps(payload, default=str)[:1500])
    raw = payload.get("responses")
    if raw is None:
        raw = next((v for v in payload.values() if isinstance(v, (dict, list))), {})
    rows = []
    def add(qid, qtext, ans):
        if isinstance(ans, (dict, list)):
            if isinstance(ans, list) and all(isinstance(x, dict) for x in ans):
                ans = "; ".join(str(x.get("label") or x.get("name") or x.get("value") or x) for x in ans)
            elif isinstance(ans, dict):
                ans = ans.get("label") or ans.get("name") or ans.get("value") or json.dumps(ans, default=str)
            else:
                ans = "; ".join(str(x) for x in ans)
        if ans is not None:
            ans = str(ans)
            # signatures / uploads come back as protected media paths: not an answer
            if ans.startswith("/protected/") or ans.startswith("http") or len(ans) > 300:
                ans = "[file]"
        rows.append((str(qid), str(qtext or ""), ans))
    if isinstance(raw, dict):              # {field_hash: {"label":..., "response":...}}
        for qid, it in raw.items():
            if isinstance(it, dict):
                with _elig_lock:
                    _elig_schema["item_keys"].update(it.keys())
                ans, akey = None, None
                for k in ANSWER_KEYS:
                    if k in it and it[k] not in (None, "", [], {}):
                        ans, akey = it[k], k
                        break
                if akey:
                    with _elig_lock:
                        _elig_schema["answer_key_used"][akey] += 1
                add(qid, it.get("label") or it.get("name") or it.get("question"), ans)
            else:
                add(qid, qid, it)
    elif isinstance(raw, list):            # [{id, name, response}]
        for it in raw:
            if not isinstance(it, dict):
                continue
            with _elig_lock:
                _elig_schema["item_keys"].update(it.keys())
            ans, akey = None, None
            for k in ANSWER_KEYS:
                if k in it and it[k] not in (None, "", [], {}):
                    ans, akey = it[k], k
                    break
            if akey:
                with _elig_lock:
                    _elig_schema["answer_key_used"][akey] += 1
            add(it.get("id"), it.get("name") or it.get("label") or it.get("question"), ans)
    header = {"eligibility_present": True,
              "eligibility_started": to_date(payload.get("created_at")),
              "eligibility_updated": to_date(payload.get("updated_at")),
              "eligibility_completed": payload.get("is_completed"),
              "eligibility_answers_n": sum(1 for r in rows if r[2] not in (None, "", "[file]"))}
    return header, rows

# ---------------------------------------------------------------- SharePoint (compact)
_tok = _drv = None
def graph_token():
    global _tok
    if _tok: return _tok
    r = requests.post(f"https://login.microsoftonline.com/{wmill.get_variable('u/epham2/AZURE_TENANT_ID').strip()}/oauth2/v2.0/token",
        data={"client_id": wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_ID").strip(),
              "client_secret": wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_SECRET").strip(),
              "scope": "https://graph.microsoft.com/.default", "grant_type": "client_credentials"}, timeout=60)
    if r.status_code != 200: raise RuntimeError(f"Graph token [{r.status_code}]: {r.text[:500]}")
    _tok = r.json()["access_token"]; return _tok
def graph(m, ep, ok=(200, 201), **kw):
    h = kw.pop("headers", {}); h["Authorization"] = f"Bearer {graph_token()}"
    r = requests.request(m, f"{GRAPH_BASE}{ep}", headers=h, timeout=120, **kw)
    if r.status_code not in ok: raise RuntimeError(f"Graph {m} {ep} [{r.status_code}]: {r.text[:600]}")
    return r.json() if r.content else None
def drive_id():
    global _drv
    if _drv: return _drv
    site = graph("GET", f"/sites/{SP_HOST}:{SP_SITE}")
    d = next(x for x in graph("GET", f"/sites/{site['id']}/drives")["value"] if x.get("name") == SP_LIBRARY)
    _drv = d["id"]; return _drv
def upload(name, data: bytes):
    path = quote(f"{SP_FOLDER}/{name}", safe="/")
    item = graph("PUT", f"/drives/{drive_id()}/root:/{path}:/content",
                 headers={"Content-Type": "application/octet-stream"}, data=data)
    return item.get("webUrl")
def csv_bytes(rows):
    if not rows: return b""
    keys = []
    for r in rows:
        for k in r.keys():
            if k not in keys: keys.append(k)
    buf = io.StringIO(); w = csv.DictWriter(buf, fieldnames=keys, extrasaction="ignore")
    w.writeheader(); w.writerows(rows); return buf.getvalue().encode("utf-8-sig")


# ---------------------------------------------------------------- main
def main(
    compare_programs: list[int] = [295185, 295865],  # programs that count as "has an expedited application"
    match_all_programs: bool = True,                 # match accounts against applications in ANY program
    created_since: str = "2026-06-01",               # account-creation lower bound for the monthly view ("" = all)
    applicant_filter: str = "",                      # substring that role_summary must contain, e.g. "applicant"; "" = no filter
    anonymize: bool = True,
    anon_salt: str = "nnas-volume-report",
    upload_to_sharepoint: bool = True,
    user_limit: int = 0,                             # smoke test
    pull_eligibility: bool = True,
    eligibility_scope: str = "applicants",           # applicants | never_applied | none
    max_workers: int = 8,
):
    pull_eligibility = str(pull_eligibility).lower() in ("true", "1", "yes") and eligibility_scope != "none"
    anonymize = str(anonymize).lower() in ("true", "1", "yes")
    upload_to_sharepoint = str(upload_to_sharepoint).lower() in ("true", "1", "yes")
    match_all_programs = str(match_all_programs).lower() in ("true", "1", "yes")
    compare_programs = {int(p) for p in compare_programs}

    s = requests.Session()
    s.headers.update({"Authorization": f"Bearer {wmill.get_variable(SMAPPLY_TOKEN_PATH)}",
                      "Accept": "application/json"})

    # ---- 1. users ----
    print("listing users...")
    users = paged(s, f"{BASE_URL}/users/")
    if int(user_limit):
        users = users[: int(user_limit)]
    print(f"  {len(users)} user accounts")
    schema = {"fields_seen": sorted({k for u in users[:200] for k in u.keys()}),
              "role_values": Counter(role_summary(u) for u in users).most_common(15),
              "date_field_used": next((k for k in DATE_KEYS if any(u.get(k) for u in users[:50])), None),
              "active_field_used": next((k for k in ACTIVE_KEYS if any(u.get(k) for u in users[:50])), None),
              "status_values": Counter(str(u.get("status")) for u in users).most_common(8),
              "email_verified_values": Counter(str(u.get("email_verified")) for u in users).most_common(4)}
    print("  user fields:", schema["fields_seen"])
    print("  role values:", schema["role_values"])

    # ---- 2. applications (all programs, all states) ----
    print("listing applications" + (" (all programs)" if match_all_programs else f" (programs {sorted(compare_programs)})") + "...")
    if match_all_programs:
        apps = paged(s, f"{BASE_URL}/applications/")
    else:
        apps = []
        for pid in compare_programs:
            apps.extend(paged(s, f"{BASE_URL}/applications/", {"program": pid}))
    print(f"  {len(apps)} applications")

    by_user = defaultdict(list)
    for a in apps:
        uid = (a.get("applicant") or {}).get("id")
        if uid is not None:
            by_user[uid].append(a)

    # ---- 3. join ----
    rows = []
    for u in users:
        rs = role_summary(u)
        if applicant_filter and (not rs or applicant_filter.lower() not in rs.lower()):
            continue
        ua = by_user.get(u.get("id"), [])
        exp = [a for a in ua if (a.get("program") or {}).get("id") in compare_programs]
        created = user_created(u)
        rows.append({
            "account_id": anon(u.get("id"), anon_salt) if anonymize else u.get("id"),
            "account_created": created, "account_month": created[:7] if created else None,
            "last_active": user_last_active(u),
            "account_status": u.get("status"),
            "email_verified": u.get("email_verified"),
            "invited_at": to_date(u.get("invited_at")),
            "role_summary": rs,
            "is_staff": bool(rs and any(x in rs.lower() for x in ("admin", "reviewer", "recommender"))),
            "has_any_application": bool(ua), "n_applications": len(ua),
            "has_expedited_application": bool(exp), "n_expedited": len(exp),
            "programs": "; ".join(sorted({str((a.get("program") or {}).get("name")) for a in ua})) or None,
            "first_application_created": min((to_date(a.get("created_at")) for a in ua if a.get("created_at")), default=None),
            "any_application_paid_hint": any((a.get("last_submitted_at")) for a in ua),  # weak hint only
        })
    rows.sort(key=lambda r: (r["account_created"] or "", r["account_id"]))

    # ---- 3b. eligibility (answered before any application exists) ----
    elig_rows, elig_errors = [], []
    if pull_eligibility:
        id_of = {}
        for u in users:
            id_of[anon(u.get("id"), anon_salt) if anonymize else u.get("id")] = u.get("id")
        targets = [r for r in rows if not r["is_staff"]
                   and (eligibility_scope != "never_applied" or not r["has_any_application"])]
        print(f"fetching eligibility for {len(targets)} accounts ({eligibility_scope})...")
        def job(r):
            return r, fetch_eligibility(s, id_of[r["account_id"]])
        with ThreadPoolExecutor(max_workers=int(max_workers)) as pool:
            futs = [pool.submit(job, r) for r in targets]
            done = 0
            for f in as_completed(futs):
                try:
                    r, payload = f.result()
                    header, qrows = parse_eligibility(payload)
                    r.update(header)
                    for qid, qtext, ans in qrows:
                        elig_rows.append({"account_id": r["account_id"], "has_any_application": r["has_any_application"],
                                          "has_expedited_application": r["has_expedited_application"],
                                          "account_month": r["account_month"],
                                          "question_id": qid, "question": qtext, "answer": ans})
                except Exception as exc:
                    elig_errors.append(str(exc)[:200])
                done += 1
                if done % 250 == 0:
                    print(f"  {done}/{len(targets)}")
        for r in rows:
            for k in ("eligibility_present", "eligibility_started", "eligibility_updated",
                      "eligibility_completed", "eligibility_answers_n"):
                r.setdefault(k, None)
        print(f"  eligibility: {sum(1 for r in rows if r.get('eligibility_present'))} present, "
              f"{sum(1 for r in rows if r.get('eligibility_completed'))} completed, {len(elig_errors)} errors")

    # question x answer -> applied rate (k>=5)
    sig = defaultdict(lambda: {"accounts": 0, "applied": 0, "expedited": 0})
    for e in elig_rows:
        k = (e["question_id"], e["question"][:90], e["answer"])
        d = sig[k]; d["accounts"] += 1
        d["applied"] += 1 if e["has_any_application"] else 0
        d["expedited"] += 1 if e["has_expedited_application"] else 0
    signal = [{"question_id": k[0], "question": k[1], "answer": k[2], **v,
               "pct_applied": round(100 * v["applied"] / v["accounts"], 1)}
              for k, v in sig.items() if v["accounts"] >= 5]
    signal.sort(key=lambda x: (x["question_id"], -x["accounts"]))

    # ---- 4. monthly ----
    scoped = [r for r in rows if r["account_month"] and not r["is_staff"]
              and (not created_since or r["account_created"] >= created_since)]
    app_created = Counter(to_date(a.get("created_at"))[:7] for a in apps
                          if to_date(a.get("created_at")) and (a.get("program") or {}).get("id") in compare_programs)
    months = sorted({r["account_month"] for r in scoped} | {m for m in app_created if not created_since or m >= created_since[:7]})
    monthly = []
    for m in months:
        ms = [r for r in scoped if r["account_month"] == m]
        n = len(ms); w_any = sum(r["has_any_application"] for r in ms); w_exp = sum(r["has_expedited_application"] for r in ms)
        never = [r for r in ms if not r["has_any_application"]]
        monthly.append({"month": m, "accounts_created": n,
                        "with_any_application": w_any, "with_expedited_application": w_exp,
                        "never_applied": n - w_any,
                        "pct_never_applied": round(100 * (n - w_any) / n, 1) if n else None,
                        "never_applied_email_unverified": sum(1 for r in never if str(r["email_verified"]).lower() in ("false", "0", "none")),
                        "never_applied_active_after_signup": sum(1 for r in never if r["last_active"] and r["account_created"] and r["last_active"] > r["account_created"]),
                        "expedited_applications_created": app_created.get(m, 0)})

    applicants = [r for r in rows if not r["is_staff"]]
    total = len(applicants); never = sum(1 for r in applicants if not r["has_any_application"])
    summary = {"applicant_accounts": total, "accounts_never_applied": never,
               "pct_never_applied": round(100 * never / total, 1) if total else None,
               "accounts_staff_excluded": sum(1 for r in rows if r["is_staff"]),
               "accounts_with_expedited": sum(1 for r in applicants if r["has_expedited_application"]),
               "accounts_with_other_program_only": sum(1 for r in applicants if r["has_any_application"] and not r["has_expedited_application"]),
               "never_applied_email_unverified": sum(1 for r in applicants if not r["has_any_application"] and str(r["email_verified"]).lower() in ("false","0","none")),
               "applications_total": len(apps), "applications_without_matching_user": sum(1 for a in apps if (a.get("applicant") or {}).get("id") not in {u.get("id") for u in users}),
               "scoped_since": created_since or "all"}
    print(json.dumps(summary, indent=1))

    ts = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    files = {f"accounts_{ts}.csv": csv_bytes(rows), f"accounts_monthly_{ts}.csv": csv_bytes(monthly),
             f"eligibility_{ts}.csv": csv_bytes(elig_rows),
             f"eligibility_signal_{ts}.csv": csv_bytes(signal),
             f"accounts_meta_{ts}.json": json.dumps({"summary": summary, "user_schema": schema,
                                                     "eligibility_schema": {k: (sorted(v) if isinstance(v, set) else (dict(v) if isinstance(v, Counter) else v)) for k, v in _elig_schema.items()},
                                                     "eligibility_errors": elig_errors[:20],
                                                     "generated_at": ts}, indent=1, default=str).encode()}
    urls = {}
    if upload_to_sharepoint:
        for n, d in files.items():
            try: urls[n] = upload(n, d)
            except Exception as exc: urls[n] = f"UPLOAD FAILED: {exc}"
    summary["eligibility_present"] = sum(1 for r in rows if r.get("eligibility_present"))
    summary["eligibility_completed"] = sum(1 for r in rows if r.get("eligibility_completed"))
    summary["eligibility_rows"] = len(elig_rows)
    return {"summary": summary, "user_schema": schema, "monthly": monthly,
            "eligibility_schema": {k: (sorted(v) if isinstance(v, set) else (dict(v) if isinstance(v, Counter) else v)) for k, v in _elig_schema.items()},
            "eligibility_signal_top": sorted(signal, key=lambda x: -x["accounts"])[:30],
            "files": urls or list(files.keys())}