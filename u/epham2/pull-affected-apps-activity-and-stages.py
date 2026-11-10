# py: 3.12
# requirements:
# requests
# wmill

"""
READ-ONLY: pull the activity log, current stage and stage history for the
applications in post_switch_country_review.xlsx, and work out for each one
whether the Registration/Licensing History task had any activity before
Friday's switch (2026-10-30 21:20 UTC = 3:20 PM MDT).

Same auth as the other scripts (u/epham2/SMAPPLY_ACCESS_TOKEN, and the Graph
variables for the SharePoint upload). Only GET requests are sent to SM Apply.

The API documentation lists Activity schemas but no activity endpoint, so the
script tries likely endpoints and also checks the application record itself.
Run discover=True first: it reports which endpoints respond for one
application and shows the shape of the data, without saving anything.

Outputs (uploaded to SharePoint, Applications/Country list backups):
  activity_summary_<ts>.csv  One row per application: Group B answer
                             (Yes/No/Unknown), first and last registration
                             task events, current stage, stage history.
  activity_events_<ts>.csv   Every activity event: time (UTC and Mountain),
                             action, task, role. Person names are NOT written.
  stage_history_<ts>.csv     Every stage record for each application.
"""

import csv
import io
import json
import re
import time
from datetime import datetime, timezone, timedelta
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import quote

import requests
import wmill


# =========================================================
# CONFIG
# =========================================================

SMAPPLY_BASE = "https://nnas.smapply.io/api"
SMAPPLY_TOKEN_VAR = "u/epham2/SMAPPLY_ACCESS_TOKEN"

REGISTRATION_TASK_NAME = "Registration/Licensing History"
REGISTRATION_TASK_ID = 2044263
SWITCH_UTC = datetime(2026, 10, 30, 21, 20, tzinfo=timezone.utc)

APPS_BY_GROUP = {
    "Known": [35942863],
    "B": [35500617, 35589355, 35590885, 35596873, 35604818, 35625530, 35647178,
          35650362, 35671256, 35671271, 35690631, 35703034, 35716321, 35768338,
          35791245, 35801406, 35817358, 35853913, 35862552, 35864060, 35874317,
          35879414, 35879876, 35894301, 35896712, 35904275, 35908317, 35914759,
          35925873, 35934846, 35937949, 35945656, 35945662, 35949706, 35952334,
          35962741, 35968951, 35969748, 35969773, 35970115, 35971544],
    "A": [35975343, 35976737, 35978096, 35979241, 35980266, 35981909, 35982316,
          35984792, 35985498, 35986821, 35988943, 35989640, 35990640, 35990665,
          35991279, 35991913, 35992469, 35994869, 35996425, 35997088, 35997252,
          35997774, 35999153, 36000810, 36002836, 36003360, 36004071, 36007610,
          36008235, 36009809],
}

# Candidate activity endpoints; the first one that returns events is used.
ACTIVITY_PATHS = [
    "/applications/{aid}/activity/",
    "/applications/{aid}/activities/",
    "/activity/?application={aid}",
    "/activities/?application={aid}",
]
STAGES_PATH = "/applications/{aid}/stages/"
APPLICATION_PATH = "/applications/{aid}/"

# Keys that may hold a person's name or contact details; never written out.
PERSONAL_KEYS = {"first_name", "last_name", "name", "full_name", "email", "user",
                 "actor", "owner", "applicant", "username", "phone"}

REQUEST_DELAY = 0.2
MAX_RETRIES = 5

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SHAREPOINT_HOSTNAME = "nnasca.sharepoint.com"
SITE_PATH = "/sites/NNAS-Portal-Sharepoint"
DOCUMENT_LIBRARY_NAME = "Applications"
BACKUP_FOLDER = "Country list backups"


# =========================================================
# SM APPLY (read-only)
# =========================================================

_HEADERS: Optional[Dict[str, str]] = None


def _headers() -> Dict[str, str]:
    global _HEADERS
    if _HEADERS is None:
        token = wmill.get_variable(SMAPPLY_TOKEN_VAR)
        _HEADERS = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    return _HEADERS


def _get(url: str) -> Tuple[int, Any]:
    """Return (status, json or None). 404/403 are returned, not raised."""
    for attempt in range(1, MAX_RETRIES + 1):
        resp = requests.get(url, headers=_headers(), timeout=120)
        if resp.status_code in (429, 500, 502, 503, 504):
            time.sleep(int(resp.headers.get("Retry-After", 2 ** attempt)))
            continue
        time.sleep(REQUEST_DELAY)
        try:
            return resp.status_code, resp.json() if resp.ok else None
        except ValueError:
            return resp.status_code, None
    return 0, None


def _get_all(path: str) -> Tuple[int, List[Any]]:
    """Follow DRF pagination ('results' and 'next'); also accepts a bare list."""
    url: Optional[str] = f"{SMAPPLY_BASE}{path}"
    sep = "&" if "?" in path else "?"
    url = f"{url}{sep}per_page=100"
    out: List[Any] = []
    status = 0
    while url:
        status, data = _get(url)
        if status != 200 or data is None:
            return status, out
        if isinstance(data, list):
            out.extend(data)
            break
        out.extend(data.get("results", []))
        url = data.get("next")
    return status, out


# =========================================================
# PARSING HELPERS
# =========================================================

TIME_KEYS = ("created_at", "date", "timestamp", "time", "created", "occurred_at", "date_created")
TEXT_KEYS = ("description", "text", "message", "action", "verb", "summary", "event")
VERBS = ("began", "completed", "edited", "submitted", "created", "reopened", "uncompleted",
         "moved", "assigned", "deleted", "uploaded")


def _parse_time(v: Any) -> Optional[datetime]:
    if not v:
        return None
    try:
        dt = datetime.fromisoformat(str(v).replace("Z", "+00:00"))
    except ValueError:
        return None
    return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)   # API times are UTC


def _mountain(dt: datetime) -> str:
    # Mountain time: MDT (UTC-6) until 2026-11-01 08:00 UTC, then MST (UTC-7).
    if dt < datetime(2026, 11, 1, 8, 0, tzinfo=timezone.utc):
        return (dt - timedelta(hours=6)).strftime("%Y-%m-%d %H:%M MDT")
    return (dt - timedelta(hours=7)).strftime("%Y-%m-%d %H:%M MST")


def _first(d: Dict[str, Any], keys) -> Any:
    for k in keys:
        if d.get(k) not in (None, ""):
            return d[k]
    return None


def parse_event(ev: Dict[str, Any]) -> Dict[str, Any]:
    """Pull time, action, task and role from one activity event without
    keeping anyone's name."""
    when = _parse_time(_first(ev, TIME_KEYS))
    text = str(_first(ev, TEXT_KEYS) or "")
    task = ev.get("task")
    task_name = task.get("name") if isinstance(task, dict) else None
    if not task_name:
        m = re.search(r'task\s+"([^"]+)"', text)
        task_name = m.group(1) if m else ""
    role = ev.get("role")
    if isinstance(role, dict):
        role = role.get("name")
    if not role:
        m = re.search(r'as role\s+"([^"]+)"', text)
        role = m.group(1) if m else ""
    verb = str(ev.get("verb") or ev.get("action") or "")
    if not verb or len(verb) > 30:
        verb = next((v for v in VERBS if re.search(rf"\b{v}\b", text)), "")
    form_edit = bool(re.search(r"edited form", text))
    return {"when": when, "action": verb, "task": task_name or "",
            "role": role or "", "admin_form_edit": form_edit}


def find_event_list(obj: Any) -> List[Dict[str, Any]]:
    """Find a list of activity-like dicts inside an application record."""
    if isinstance(obj, dict):
        for k, v in obj.items():
            if "activit" in k.lower() and isinstance(v, list) and v and isinstance(v[0], dict):
                return v
            found = find_event_list(v)
            if found:
                return found
    elif isinstance(obj, list):
        for item in obj:
            found = find_event_list(item)
            if found:
                return found
    return []


def strip_personal(obj: Any) -> Any:
    if isinstance(obj, dict):
        return {k: ("[removed]" if k.lower() in PERSONAL_KEYS else strip_personal(v))
                for k, v in obj.items()}
    if isinstance(obj, list):
        return [strip_personal(v) for v in obj]
    return obj


def stage_row(rec: Dict[str, Any]) -> Dict[str, Any]:
    st = rec.get("stage") if isinstance(rec.get("stage"), dict) else {}
    return {
        "stage_record_id": rec.get("id"),
        "stage_id": st.get("id") or rec.get("stage"),
        "stage_name": st.get("title") or st.get("name") or rec.get("name") or rec.get("title"),
        "entered_utc": rec.get("created_at") or rec.get("date_entered") or rec.get("started_at"),
        "exited_utc": rec.get("date_exited") or rec.get("ended_at") or rec.get("completed_at"),
        "status": rec.get("status"),
        "decision": json.dumps(rec.get("decision")) if isinstance(rec.get("decision"), (dict, list)) else rec.get("decision"),
        "updated_utc": rec.get("updated_at"),
    }


# =========================================================
# PER-APPLICATION PULL
# =========================================================

_ACTIVITY_PATH: Optional[str] = None   # remembered once one works


def pull_activity(aid: int, app_record: Dict[str, Any]) -> Tuple[str, List[Dict[str, Any]]]:
    global _ACTIVITY_PATH
    paths = [_ACTIVITY_PATH] if _ACTIVITY_PATH else ACTIVITY_PATHS
    for p in paths:
        status, items = _get_all(p.format(aid=aid))
        if status == 200 and items:
            _ACTIVITY_PATH = p
            return p, items
    embedded = find_event_list(app_record)
    if embedded:
        return "application record", embedded
    return "not found", []


def pull_application(aid: int) -> Dict[str, Any]:
    status, app = _get(f"{SMAPPLY_BASE}{APPLICATION_PATH.format(aid=aid)}")
    app = app if status == 200 and isinstance(app, dict) else {}

    source, raw_events = pull_activity(aid, app)
    events = [parse_event(e) for e in raw_events if isinstance(e, dict)]
    events.sort(key=lambda e: e["when"] or datetime.max.replace(tzinfo=timezone.utc))

    reg = [e for e in events if e["task"].strip().lower() == REGISTRATION_TASK_NAME.lower()]
    reg_before = [e for e in reg if e["when"] and e["when"] < SWITCH_UTC]
    if source == "not found":
        answer = "Unknown"
    else:
        answer = "Yes" if reg_before else "No"

    s_status, stages = _get_all(STAGES_PATH.format(aid=aid))
    stage_rows = [stage_row(s) for s in stages if isinstance(s, dict)]
    cur = app.get("current_stage") or app.get("stage") or {}
    current_stage = (cur.get("title") or cur.get("name") or cur.get("id")) if isinstance(cur, dict) else cur

    return {
        "aid": aid,
        "activity_source": source,
        "events": events,
        "summary": {
            "activity_before_switch": answer,
            "registration_events": len(reg),
            "registration_events_before_switch": len(reg_before),
            "first_registration_event": _mountain(reg[0]["when"]) if reg and reg[0]["when"] else "",
            "last_registration_event": _mountain(reg[-1]["when"]) if reg and reg[-1]["when"] else "",
            "first_registration_actions_before_switch": "; ".join(
                f'{e["action"]} {_mountain(e["when"])}' for e in reg_before[:5]),
            "current_stage": current_stage,
            "stage_history": " > ".join(str(r["stage_name"] or r["stage_id"]) for r in
                                        sorted(stage_rows, key=lambda r: str(r["entered_utc"] or ""))),
            "stages_status": s_status,
        },
        "stages": stage_rows,
        "app_keys": sorted(app.keys()),
        "raw_sample": strip_personal(raw_events[:5]),
        "raw_stage_sample": strip_personal(stages[:3]),
    }


# =========================================================
# SHAREPOINT UPLOAD
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


def _csv(rows: List[Dict[str, Any]], fields: List[str]) -> bytes:
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=fields, extrasaction="ignore")
    w.writeheader()
    w.writerows(rows)
    return buf.getvalue().encode("utf-8")


# =========================================================
# MAIN
# =========================================================

def main(discover: bool = False, sample_appid: int = 35975343, upload: bool = True):

    if discover:
        # 35975343 is the application from the activity log example: its
        # registration task was begun after the switch, so the answer should be "No".
        r = pull_application(int(sample_appid))
        return {
            "activity_source": r["activity_source"],
            "events_found": len(r["events"]),
            "parsed_events": [{**e, "when": _mountain(e["when"]) if e["when"] else None}
                              for e in r["events"]][:40],
            "summary": r["summary"],
            "application_keys": r["app_keys"],
            "raw_activity_sample_names_removed": r["raw_sample"],
            "raw_stage_sample_names_removed": r["raw_stage_sample"],
        }

    summary_rows, event_rows, stage_rows = [], [], []
    for group, ids in APPS_BY_GROUP.items():
        for aid in ids:
            try:
                r = pull_application(aid)
            except Exception as exc:
                summary_rows.append({"application_id": aid, "group": group,
                                     "activity_before_switch": "Unknown",
                                     "activity_source": f"error: {str(exc)[:200]}"})
                continue
            summary_rows.append({"application_id": aid, "group": group,
                                 "activity_source": r["activity_source"], **r["summary"]})
            for e in r["events"]:
                event_rows.append({
                    "application_id": aid, "group": group,
                    "time_utc": e["when"].strftime("%Y-%m-%d %H:%M:%S") if e["when"] else "",
                    "time_mountain": _mountain(e["when"]) if e["when"] else "",
                    "before_switch": "Yes" if e["when"] and e["when"] < SWITCH_UTC else "No",
                    "action": e["action"], "task": e["task"], "role": e["role"],
                })
            for s in r["stages"]:
                stage_rows.append({"application_id": aid, "group": group, **s})

    counts: Dict[str, Dict[str, int]] = {}
    for s in summary_rows:
        g = counts.setdefault(s["group"], {})
        g[s["activity_before_switch"]] = g.get(s["activity_before_switch"], 0) + 1

    result: Dict[str, Any] = {
        "applications": len(summary_rows),
        "activity_before_switch_by_group": counts,
        "activity_sources": sorted({s.get("activity_source", "") for s in summary_rows}),
        "note": "Group A should all be 'No'. Any 'Yes' in Group A or 'Unknown' anywhere needs a manual look.",
    }

    if upload:
        ts = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
        result["sharepoint_files"] = upload_to_sharepoint({
            f"activity_summary_{ts}.csv": _csv(summary_rows, [
                "application_id", "group", "activity_before_switch", "registration_events",
                "registration_events_before_switch", "first_registration_event",
                "last_registration_event", "first_registration_actions_before_switch",
                "current_stage", "stage_history", "activity_source", "stages_status"]),
            f"activity_events_{ts}.csv": _csv(event_rows, [
                "application_id", "group", "time_utc", "time_mountain", "before_switch",
                "action", "task", "role"]),
            f"stage_history_{ts}.csv": _csv(stage_rows, [
                "application_id", "group", "stage_record_id", "stage_id", "stage_name",
                "entered_utc", "exited_utc", "status", "decision", "updated_utc"]),
        })
    return result