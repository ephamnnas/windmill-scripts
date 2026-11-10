# py: 3.12
# requirements:
# requests
# wmill

"""
READ-ONLY: pull the question settings (label and dropdown choices) for the
registration country fields in programs 295865 and 295185.

The backup showed the fields are stored under these keys in the
"Registration/Licensing History" task (task 2044263 in program 295185,
task 2053917 in program 295865):

  GNviXZh7zt  0gqFv2wuRp  78nblFoLys  Je7rLX84r5  TqaiTTrVqd  pkmuzT9P75  pr11MJ2OXs

The SM Apply API documentation does not clearly name the endpoint that
returns form question settings, so this script tries a list of likely
endpoints, records which ones respond, and searches each response for the
field keys above. Wherever a field key is found next to a list of choices,
the choices are captured with their 0-based position.

Outputs:
  * The Windmill result: which endpoints worked, and a summary per field
    (label, number of choices, first and last few choices).
  * SharePoint (Applications/Country list backups):
      country_question_settings_<ts>.csv   one row per field and choice
      country_question_settings_<ts>.json  raw responses from every endpoint
                                           that returned data, for review

Run with upload=False first to check the endpoint results on screen.
"""

import csv
import io
import json
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

PROGRAM_TASKS = {295185: 2044263, 295865: 2053917}

# One application per program, used for application-level endpoints.
SAMPLE_APPS = {295185: 34643255, 295865: 34933184}

FIELD_KEYS = ["GNviXZh7zt", "78nblFoLys", "0gqFv2wuRp", "TqaiTTrVqd",
              "pkmuzT9P75", "Je7rLX84r5", "pr11MJ2OXs"]

# Candidate endpoints. {pid} = program ID, {tid} = task ID, {aid} = sample app ID.
CANDIDATE_PATHS = [
    "/tasks/{tid}/",
    "/tasks/{tid}/fields/",
    "/tasks/{tid}/form/",
    "/forms/{tid}/",
    "/programs/{pid}/",
    "/programs/{pid}/tasks/",
    "/programs/{pid}/tasks/{tid}/",
    "/programs/{pid}/fields/",
    "/applications/{aid}/tasks/{tid}/",
    "/applications/{aid}/tasks/{tid}/?expand=fields",
    "/fields/?task={tid}",
    "/fields/?program={pid}",
]

CHOICE_KEYS = ("choices", "options", "items", "values")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SHAREPOINT_HOSTNAME = "nnasca.sharepoint.com"
SITE_PATH = "/sites/NNAS-Portal-Sharepoint"
DOCUMENT_LIBRARY_NAME = "Applications"
BACKUP_FOLDER = "Country list backups"


# =========================================================
# SM APPLY (read-only)
# =========================================================

def _headers() -> Dict[str, str]:
    token = wmill.get_variable(SMAPPLY_TOKEN_VAR)
    return {"Authorization": f"Bearer {token}", "Accept": "application/json"}


def _try_get(path: str, headers: Dict[str, str]) -> Dict[str, Any]:
    url = f"{SMAPPLY_BASE}{path}"
    try:
        resp = requests.get(url, headers=headers, timeout=120)
    except requests.RequestException as exc:
        return {"path": path, "status": None, "error": str(exc)[:300]}
    out: Dict[str, Any] = {"path": path, "status": resp.status_code}
    if resp.ok:
        try:
            out["json"] = resp.json()
        except ValueError:
            out["error"] = "response was not JSON"
    else:
        out["error"] = resp.text[:300]
    return out


def _choice_text(c: Any) -> str:
    if isinstance(c, dict):
        for k in ("label", "text", "name", "value", "title"):
            if c.get(k) not in (None, ""):
                return str(c[k])
        return json.dumps(c)
    return str(c)


def find_field_settings(obj: Any, path: str = "") -> List[Dict[str, Any]]:
    """Find any object that references one of FIELD_KEYS (as its key, id,
    name or slug) and capture its label and choice list if present."""
    found: List[Dict[str, Any]] = []
    if isinstance(obj, dict):
        for k, v in obj.items():
            sub = f"{path}.{k}" if path else k
            ref = k if k in FIELD_KEYS else None
            if ref is None and isinstance(v, dict):
                for idk in ("key", "id", "name", "slug", "field_id", "uuid"):
                    if str(v.get(idk, "")) in FIELD_KEYS:
                        ref = str(v[idk])
                        break
            if ref and isinstance(v, dict):
                choices = next((v[c] for c in CHOICE_KEYS if isinstance(v.get(c), list)), None)
                found.append({
                    "field_key": ref,
                    "found_at": sub,
                    "label": v.get("label") or v.get("title") or v.get("name"),
                    "type": v.get("type") or v.get("field_type"),
                    "choices": [_choice_text(c) for c in choices] if choices is not None else None,
                    "other_settings": {kk: vv for kk, vv in v.items()
                                       if kk not in CHOICE_KEYS and not isinstance(vv, (dict, list))},
                })
            found.extend(find_field_settings(v, sub))
    elif isinstance(obj, list):
        for i, item in enumerate(obj):
            found.extend(find_field_settings(item, f"{path}[{i}]"))
    return found


# =========================================================
# SHAREPOINT UPLOAD (same Graph auth as the report script)
# =========================================================

def upload_to_sharepoint(files: Dict[str, bytes]) -> Dict[str, str]:
    tenant = wmill.get_variable("u/epham2/AZURE_TENANT_ID").strip()
    tok = requests.post(
        f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token",
        data={
            "client_id": wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_ID").strip(),
            "client_secret": wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_SECRET").strip(),
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        },
        timeout=60,
    )
    if tok.status_code != 200:
        raise RuntimeError(f"Graph token request failed {tok.status_code}: {tok.text[:1000]}")
    h = {"Authorization": f"Bearer {tok.json()['access_token']}"}
    site = requests.get(f"{GRAPH_BASE}/sites/{SHAREPOINT_HOSTNAME}:{SITE_PATH}", headers=h, timeout=60).json()
    drives = requests.get(f"{GRAPH_BASE}/sites/{site['id']}/drives", headers=h, timeout=60).json()["value"]
    drive = next(d for d in drives if d.get("name") == DOCUMENT_LIBRARY_NAME)
    urls = {}
    for name, content in files.items():
        path = quote(f"{BACKUP_FOLDER}/{name}", safe="/")
        resp = requests.put(f"{GRAPH_BASE}/drives/{drive['id']}/root:/{path}:/content",
                            headers={**h, "Content-Type": "application/octet-stream"},
                            data=content, timeout=300)
        if resp.status_code not in (200, 201):
            raise RuntimeError(f"Upload of {name} failed {resp.status_code}: {resp.text[:1000]}")
        urls[name] = resp.json().get("webUrl", "")
    return urls


# =========================================================
# MAIN
# =========================================================

def main(upload: bool = False):
    headers = _headers()
    endpoint_results: List[Dict[str, Any]] = []
    settings: List[Dict[str, Any]] = []

    for pid, tid in PROGRAM_TASKS.items():
        aid = SAMPLE_APPS.get(pid)
        for template in CANDIDATE_PATHS:
            if "{aid}" in template and not aid:
                continue
            path = template.format(pid=pid, tid=tid, aid=aid)
            res = _try_get(path, headers)
            res["program"] = pid
            if "json" in res:
                hits = find_field_settings(res["json"])
                for h in hits:
                    h.update({"program": pid, "task_id": tid, "endpoint": path})
                settings.extend(hits)
                res["field_keys_found"] = sorted({h["field_key"] for h in hits})
            endpoint_results.append(res)

    # Keep the most complete entry per program and field (one with choices wins).
    best: Dict[tuple, Dict[str, Any]] = {}
    for s in settings:
        key = (s["program"], s["field_key"])
        if key not in best or (s["choices"] and not best[key]["choices"]):
            best[key] = s

    summary = []
    for (pid, fk), s in sorted(best.items()):
        ch = s["choices"] or []
        summary.append({
            "program": pid, "field_key": fk, "label": s["label"], "type": s["type"],
            "endpoint": s["endpoint"], "choice_count": len(ch) if s["choices"] is not None else None,
            "first_choices": list(enumerate(ch[:6])),
            "last_choices": list(enumerate(ch))[-6:],
        })

    result = {
        "endpoints": [{"program": r["program"], "path": r["path"], "status": r["status"],
                       "field_keys_found": r.get("field_keys_found", []),
                       "error": r.get("error")} for r in endpoint_results],
        "fields": summary,
        "fields_missing": [f"{pid}:{fk}" for pid in PROGRAM_TASKS for fk in FIELD_KEYS
                           if (pid, fk) not in best],
    }

    if upload:
        ts = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
        buf = io.StringIO()
        w = csv.writer(buf)
        w.writerow(["program", "task_id", "field_key", "label", "type", "endpoint",
                    "choice_index_0_based", "choice"])
        for (pid, fk), s in sorted(best.items()):
            for i, c in enumerate(s["choices"] or [None]):
                w.writerow([pid, s["task_id"], fk, s["label"], s["type"], s["endpoint"],
                            i if c is not None else "", c if c is not None else "(no choices found)"])
        raw = [{"program": r["program"], "path": r["path"], "json": r["json"]}
               for r in endpoint_results if "json" in r]
        result["sharepoint_files"] = upload_to_sharepoint({
            f"country_question_settings_{ts}.csv": buf.getvalue().encode("utf-8"),
            f"country_question_settings_{ts}.json": json.dumps(raw, indent=2).encode("utf-8"),
        })

    return result