# requirements:
# requests
# wmill

"""
ELIGIBILITY COHORT -- one run, one readable list.

Lists every applicant account whose answer to a chosen eligibility question is
in the wanted set (default: Yes or Unsure), with name, email, the option they
selected, when they completed eligibility, whether they have created an
application, and the application reference if so.

Output
  eligibility_cohort_<ts>.csv in Applications/zDATA/pipeline_pull/  (PII; internal)
  and the same rows returned inline for a quick read.

Answer codes
  SMApply returns eligibility answers as option codes, not text. The default
  mapping is 0 = Yes, 1 = No, 2 = Unsure, the same convention as the task
  forms. The run prints the code distribution so it can be checked against
  one known account; if the mapping is wrong, pass answer_labels with the
  corrected text and re-run.
"""
import csv, io, random, time
from datetime import datetime, timezone
from urllib.parse import quote
from concurrent.futures import ThreadPoolExecutor, as_completed
import requests, wmill

BASE_URL = "https://nnas.smapply.io/api"
SMAPPLY_TOKEN_PATH = "u/epham2/SMAPPLY_ACCESS_TOKEN"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SP_HOST, SP_SITE, SP_LIBRARY, SP_FOLDER = "nnasca.sharepoint.com", "/sites/NNAS-Portal-Sharepoint", "Applications", "zDATA/pipeline_pull"


# ------------------------------------------------------------ helpers
def get(s, url, params=None, ok404=False, retries=10):
    for attempt in range(retries):
        r = s.get(url, params=params, timeout=90)
        if r.status_code == 429:
            w = r.headers.get("Retry-After")
            time.sleep(int(w) if w and w.isdigit() else min(60, 2 ** attempt) + random.random())
            continue
        if r.status_code == 404 and ok404:
            return None
        if r.status_code >= 400:
            raise RuntimeError(f"GET {r.url} -> {r.status_code}: {r.text[:400]}")
        return r.json()
    raise RuntimeError(f"rate limited: {url}")


def paged(s, url, params=None):
    out, page = [], 1
    while True:
        p = get(s, url, {"page": page, "per_page": 50, **(params or {})})
        out.extend(p.get("results", []))
        if page >= p.get("num_pages", 1) or not p.get("next"):
            return out
        page += 1


def day(v):
    v = (str(v) if v is not None else "").strip()
    return v[:10] if len(v) >= 10 and v[4] == "-" else ""


def upload(name, data):
    tok = requests.post(f"https://login.microsoftonline.com/{wmill.get_variable('u/epham2/AZURE_TENANT_ID').strip()}/oauth2/v2.0/token",
        data={"client_id": wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_ID").strip(),
              "client_secret": wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_SECRET").strip(),
              "scope": "https://graph.microsoft.com/.default", "grant_type": "client_credentials"}, timeout=60).json()["access_token"]
    h = {"Authorization": f"Bearer {tok}"}
    site = requests.get(f"{GRAPH_BASE}/sites/{SP_HOST}:{SP_SITE}", headers=h, timeout=60).json()
    drive = next(x for x in requests.get(f"{GRAPH_BASE}/sites/{site['id']}/drives", headers=h, timeout=60).json()["value"] if x["name"] == SP_LIBRARY)["id"]
    r = requests.put(f"{GRAPH_BASE}/drives/{drive}/root:/{quote(SP_FOLDER + '/' + name, safe='/')}:/content",
                     headers={**h, "Content-Type": "application/octet-stream"}, data=data, timeout=300)
    if r.status_code not in (200, 201):
        raise RuntimeError(f"upload failed [{r.status_code}]: {r.text[:300]}")
    return r.json().get("webUrl")


# ------------------------------------------------------------ main
def main(
    question_label_starts_with: str = "Is any part of your application connected",
    answer_labels: dict = {"0": "Yes", "1": "No", "2": "Unsure"},
    wanted_answers: list = ["Yes", "Unsure"],
    eligibility_updated_since: str = "",     # e.g. "2026-10-03"; "" = no date filter
    only_without_application: bool = False,  # True = exclude people who already applied
    upload_to_sharepoint: bool = True,
    max_workers: int = 3,                    # eligibility endpoint is rate-limited; keep low
):
    only_without_application = str(only_without_application).lower() in ("true", "1", "yes")
    upload_to_sharepoint = str(upload_to_sharepoint).lower() in ("true", "1", "yes")
    wanted = {w.lower() for w in wanted_answers}
    s = requests.Session()
    s.headers.update({"Authorization": f"Bearer {wmill.get_variable(SMAPPLY_TOKEN_PATH)}", "Accept": "application/json"})

    # 1. accounts (applicants only)
    users = [u for u in paged(s, f"{BASE_URL}/users/")
             if not any(str(r.get("name", r) if isinstance(r, dict) else r).lower() in ("admin", "reviewer", "recommender")
                        for r in (u.get("roles") or []))]
    print(f"{len(users)} applicant accounts")

    # 2. applications -> references per applicant
    apps = {}
    for a in paged(s, f"{BASE_URL}/applications/"):
        uid = (a.get("applicant") or {}).get("id")
        if uid is not None:
            apps.setdefault(uid, []).append(a.get("reference_id") or str(a.get("id")))

    # 3. eligibility per account (parallel)
    def fetch(u):
        try:
            return u, get(s, f"{BASE_URL}/users/{u['id']}/eligibility/", ok404=True), None
        except Exception as exc:
            return u, None, str(exc)[:120]
    rows, code_counts, question_seen, failed = [], {}, None, []
    results = []
    with ThreadPoolExecutor(max_workers=int(max_workers)) as pool:
        futs = [pool.submit(fetch, u) for u in users]
        done = 0
        for f in as_completed(futs):
            u, e, err = f.result()
            done += 1
            if done % 500 == 0:
                print(f"  eligibility {done}/{len(users)}")
            if err:
                failed.append(u)
            else:
                results.append((u, e))
    # second, slow pass for anything that was throttled out
    if failed:
        print(f"  retrying {len(failed)} throttled accounts one at a time...")
        still = []
        for u in failed:
            time.sleep(1.0)
            u, e, err = fetch(u)
            (still if err else results).append(u if err else (u, e))
        failed = still
        if failed:
            print(f"  {len(failed)} accounts could not be fetched; listed in result as 'unfetched'")
    for u, e in results:
        if not e:
            continue
        upd = day(e.get("updated_at"))
        if eligibility_updated_since and upd < eligibility_updated_since:
            continue
        responses = e.get("responses") or {}
        ans_code = None
        for qid, item in responses.items():
            if isinstance(item, dict) and str(item.get("label", "")).lower().startswith(question_label_starts_with.lower()):
                question_seen = item.get("label")
                ans_code = item.get("response")
                break
        if ans_code is None or ans_code == "":
            continue
        code = str(ans_code)
        code_counts[code] = code_counts.get(code, 0) + 1
        label = answer_labels.get(code, f"code {code}")
        if label.lower() not in wanted:
            continue
        refs = apps.get(u["id"], [])
        if only_without_application and refs:
            continue
        rows.append({
            "name": f"{u.get('first_name') or ''} {u.get('last_name') or ''}".strip(),
            "email": u.get("email"),
            "answer": label,
            "eligibility_completed": "Yes" if e.get("is_completed") else "No",
            "eligibility_last_updated": upd,
            "signed_up": day(u.get("signed_up")),
            "last_active": day(u.get("last_active")),
            "has_application": "Yes" if refs else "No",
            "application_reference": "; ".join(refs),
            "user_id": u.get("id"),
        })
    rows.sort(key=lambda r: (r["answer"], r["has_application"], r["eligibility_last_updated"]), reverse=True)

    print(f"question matched: {question_seen!r}")
    print(f"answer code distribution (all applicants who answered): {code_counts}  -> labels {answer_labels}")
    print(f"{len(rows)} accounts answered {wanted_answers}"
          + (f" since {eligibility_updated_since}" if eligibility_updated_since else "")
          + (" and have no application" if only_without_application else ""))

    buf = io.StringIO()
    if rows:
        w = csv.DictWriter(buf, fieldnames=list(rows[0].keys())); w.writeheader(); w.writerows(rows)
    name = f"eligibility_cohort_{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S')}.csv"
    url = upload(name, buf.getvalue().encode("utf-8-sig")) if (upload_to_sharepoint and rows) else None
    return {"question": question_seen, "answer_codes_seen": code_counts, "labels_used": answer_labels,
            "count": len(rows), "file": url or name,
            "unfetched": [{"user_id": u.get("id"), "email": u.get("email")} for u in failed],
            "rows": rows}