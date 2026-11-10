# py: 3.12
# requirements:
# wmill
# requests

"""
NNAS Applicant-Copy Indexer  (self-contained Windmill script, provider-aware)
=============================================================================

For one application, find every regulator's redacted applicant copy and push
it to the provider's applicant-copy SMApply task. Specifically:

  1. Resolve the program (TM or CES) from the webhook's `award` value (or an
     explicit `program_id` input), cross-checked against
     additional_content.application_info.provider when present. The program
     selects both the SharePoint base folder and the SMApply task:
         TM  (295185) -> Applications/TM/<reference_id>,  task 2077121
         CES (295865) -> Applications/CES/<reference_id>, task 2077784
     An unknown or conflicting program is a hard error, so a copy can never be
     uploaded to the wrong program's task.
  2. Walk every regulator subfolder under "credential reports", and inside each
     one read the "Applicant Copy" subfolder:
         credential reports/<RB>/Applicant Copy/*.pdf
  3. Drop regulators the applicant is no longer registered under, and keep only
     the newest copy per (regulator, designation).
  4. Download EVERY selected copy first. If any download fails, stop before
     touching SMApply, so the task is never left emptied with nothing to
     replace it.
  5. Clear any files already on the task for this application.
  6. Upload each copy as  <reference_id>_<RB>_Applicant_report.pdf
     (a sequence number is inserted when one regulator has more than one copy,
     e.g. an RN and an LPN designation -> _1_ / _2_).
  7. Mark the task complete, only if every upload succeeded.

This is read-only against SharePoint; the only writes are the SMApply file
deletions (step 5), uploads (step 6) and task completion (step 7).
"""

import os
import re
import json
import shutil
from collections import Counter
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional, Set, Tuple
from urllib.parse import quote

import wmill
import requests


# =========================================================
# CONFIG
# =========================================================

TENANT_ID = wmill.get_variable("u/epham2/AZURE_TENANT_ID")
CLIENT_ID = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_ID")
CLIENT_SECRET = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_SECRET")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SHAREPOINT_HOSTNAME = "nnasca.sharepoint.com"
SITE_PATH = "/sites/NNAS-Portal-Sharepoint"
DOCUMENT_LIBRARY_NAME = "Applications"

CREDENTIAL_REPORTS_SUBFOLDER = "credential reports"
APPLICANT_COPY_SUBFOLDER     = "Applicant Copy"

# SMApply
BASE              = "https://nnas.smapply.io"
SMAPPLY_TOKEN_VAR = "u/epham2/SMAPPLY_ACCESS_TOKEN"

# ---------------------------------------------------------
# Program -> provider, SharePoint base folder, applicant-copy task.
# Keyed by SMApply program id (the webhook's `award` value). Paths are
# relative to the Applications document library root.
# ---------------------------------------------------------
PROGRAM_CONFIG: Dict[int, Dict[str, str]] = {
    295185: {"provider": "tm",  "base": "TM/{app_id}",  "task_id": "2077121"},
    295865: {"provider": "ces", "base": "CES/{app_id}", "task_id": "2077784"},
}

# =========================================================
# !!! HARDCODED FOR THE NNAS-ID CORRECTION REPOST !!!
# =========================================================
# Constant, not a flow input, on purpose: this step needs to actually run
# (clear the existing task files and upload the corrected applicant copies)
# every time it's used for this fix, so a mis-wired flow input can't
# silently turn it into a no-op. Change this line directly (and redeploy)
# if the intended behavior changes.
DRY_RUN = False

_ILLEGAL = re.compile(r'[\\/:*?"<>|]+')

_GRAPH_TOKEN = None
_SP_SITE_ID = None
_SP_DRIVE_ID = None


# =========================================================
# REGULATORY BODY -> SFTP/FOLDER ABBREVIATION
# =========================================================
# Same mapping used by the SFTP package builder, needed here to translate
# additional_content's regulatory_body_rn/lpn (full college names) into the
# abbreviation each "credential reports/<ABBREV>/" folder is named after, so a
# regulator the applicant is no longer registered under can be filtered out.
RB_ABBREV = {
    # --- RN ---
    "British Columbia College of Nurses and Midwives":                   "bccnm",
    "College Of Nurses Of Ontario":                                      "cno",
    "College of Registered Nurses Of Alberta":                           "crna",
    "College Of Registered Nurses Of Manitoba":                          "crnm",
    "College of Registered Nurses of Saskatchewan":                      "crns",
    "College of Nursing of New Brunswick":                               "cnnb",
    "College of Registered Nurses of Newfoundland and Labrador":         "nlcn",
    "College of Registered Nurses and Midwives of Prince Edward Island": "peicnm",
    "Nova Scotia College of Nursing - RN":                               "nscn",
    # --- LPN ---
    "Association Of New Brunswick Licensed Practical Nurses":            "anblpn",
    "College Of Licensed Practical Nurses Of Alberta":                   "clha",
    "College Of Licensed Practical Nurses Of Manitoba":                  "clpnm",
    "College of Licensed Practical Nurses of Saskatchewan":              "clpns",
    "College Of Licensed Practical Nurses Of Newfoundland And Labrador": "nlcn",
    "College of Licensed Practical Nurses of Prince Edward Island":      "peicnm",
    "Nova Scotia College of Nursing - LPN":                              "nscn",
    # --- Registered Psychiatric Nurses ---
    "College of Registered Psychiatric Nurses of Alberta":               "crpna",
    "College of Registered Psychiatric Nurses of Manitoba":              "crpnm",
    "College of Registered Psychiatric Nurses of Saskatchewan":          "crpns",
    # --- portal display-name aliases (names as they arrive in additional_content) ---
    "NL College of Nurses":                                              "nlcn",
    "BC College of Nurses and Midwives":                                 "bccnm",
    "Nova Scotia College of Nursing":                                    "nscn",
    "PEI College of Nursing & Midwifery":                                "peicnm",
    "College of Licensed Practical Nurses and Health Care Aides of Alberta": "clha",
}


def abbrev_for(rb_name: str) -> str:
    """Abbreviation for a regulatory body name (exact, then case-insensitive)."""
    if rb_name in RB_ABBREV:
        return RB_ABBREV[rb_name]
    low = {k.lower(): v for k, v in RB_ABBREV.items()}
    return low.get((rb_name or "").strip().lower(), "")


_TAG_RE = re.compile(r"<[^>]+>")


def _strip_html(s: Any) -> str:
    if not s:
        return ""
    return _TAG_RE.sub("", str(s)).strip()


def _safe_str(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, dict):
        for k in ("label", "value", "name", "text"):
            if k in v and v[k] is not None:
                return _strip_html(v[k])
        return ""
    return _strip_html(v)


def parse_additional_content(additional_content: Optional[str]) -> Dict[str, Any]:
    if not additional_content:
        return {}
    if isinstance(additional_content, dict):
        return additional_content
    s = str(additional_content).strip()
    try:
        raw = json.loads(s)
        if isinstance(raw, str):
            raw = json.loads(raw)
    except json.JSONDecodeError:
        if (s.startswith('"') and s.endswith('"')) or (s.startswith("'") and s.endswith("'")):
            try:
                raw = json.loads(s[1:-1])
                if isinstance(raw, str):
                    raw = json.loads(raw)
            except Exception:
                return {}
        else:
            return {}
    return raw if isinstance(raw, dict) else {}


def current_regulator_abbrevs(payload: Dict[str, Any]) -> Tuple[Set[str], List[str]]:
    """
    (abbrev_set, rb_names) for the applicant's CURRENT regulator selection,
    read from application_info.regulatory_body_rn/lpn. An empty set means
    "unknown selection, don't filter" rather than "selected nothing".
    """
    ai = payload.get("application_info") or {}
    rb_names: List[str] = []
    seen = set()
    for field in ("regulatory_body_rn", "regulatory_body_lpn"):
        for rb in _safe_str(ai.get(field)).split(","):
            rb = rb.strip()
            if rb and rb.lower() not in seen:
                seen.add(rb.lower())
                rb_names.append(rb)
    abbrevs = {abbrev_for(rb).upper() for rb in rb_names if abbrev_for(rb)}
    return abbrevs, rb_names


# =========================================================
# PROGRAM / PROVIDER RESOLUTION
# =========================================================

def resolve_program(program_id: Any, award: Any,
                    payload: Dict[str, Any]) -> Tuple[int, Dict[str, str]]:
    """
    Return (program_id, config). Precedence: explicit program_id input, then
    the webhook's `award`, then additional_content.application_info.provider.
    Raises RuntimeError when the program can't be resolved, isn't in
    PROGRAM_CONFIG, or the sources disagree with each other.
    """
    ids: List[int] = []
    for v in (program_id, award):
        if v is None or str(v).strip() == "":
            continue
        try:
            ids.append(int(str(v).strip()))
        except ValueError:
            raise RuntimeError(f"Program id is not numeric: {v!r}")

    provider = str((payload.get("application_info") or {}).get("provider") or "").strip().lower()

    if ids:
        if len(set(ids)) > 1:
            raise RuntimeError(f"Conflicting program ids supplied: {sorted(set(ids))}")
        pid = ids[0]
        cfg = PROGRAM_CONFIG.get(pid)
        if not cfg:
            raise RuntimeError(
                f"Unknown program id {pid}. Expected one of: "
                f"{', '.join(str(k) for k in sorted(PROGRAM_CONFIG))}."
            )
        if provider and provider != cfg["provider"]:
            raise RuntimeError(
                f"Program {pid} is {cfg['provider'].upper()} but additional_content "
                f"says provider '{provider.upper()}'. Refusing to guess."
            )
        return pid, cfg

    if provider:
        for pid, cfg in PROGRAM_CONFIG.items():
            if cfg["provider"] == provider:
                return pid, cfg
        raise RuntimeError(f"Unknown provider '{provider}' in additional_content.")

    raise RuntimeError(
        "Cannot resolve the program: no program_id / award input and no "
        "application_info.provider in additional_content."
    )


# =========================================================
# MICROSOFT GRAPH / SHAREPOINT HELPERS
# =========================================================

def _library_relative_path(path: str) -> str:
    p = (path or "").strip().strip("/")
    if p.lower().startswith("sites/nnas-portal-sharepoint/"):
        p = p[len("sites/NNAS-Portal-Sharepoint/"):]
    if p.lower().startswith("applications/"):
        p = p[len("Applications/"):]
    return p.strip("/")


def _encode_graph_path(path: str) -> str:
    return quote(path.strip("/"), safe="/")


def _get_graph_token() -> str:
    global _GRAPH_TOKEN

    if _GRAPH_TOKEN:
        return _GRAPH_TOKEN

    if not TENANT_ID:
        raise RuntimeError("Missing Windmill variable: u/epham2/AZURE_TENANT_ID")
    if not CLIENT_ID:
        raise RuntimeError("Missing Windmill variable: u/epham2/SHAREPOINT_CLIENT_ID")
    if not CLIENT_SECRET:
        raise RuntimeError("Missing Windmill variable: u/epham2/SHAREPOINT_CLIENT_SECRET")

    token_response = requests.post(
        f"https://login.microsoftonline.com/{TENANT_ID.strip()}/oauth2/v2.0/token",
        data={
            "client_id": CLIENT_ID.strip(),
            "client_secret": CLIENT_SECRET.strip(),
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        },
        timeout=60,
    )

    if token_response.status_code != 200:
        raise RuntimeError(
            "Microsoft Graph token request failed.\n"
            f"Status: {token_response.status_code}\n"
            f"Client ID suffix: {CLIENT_ID[-6:] if CLIENT_ID else None}\n"
            f"Secret present: {bool(CLIENT_SECRET)}\n"
            f"Secret length: {len(CLIENT_SECRET.strip()) if CLIENT_SECRET else 0}\n"
            f"Response body:\n{token_response.text[:3000]}"
        )

    _GRAPH_TOKEN = token_response.json()["access_token"]
    return _GRAPH_TOKEN


def _graph_request(method: str, endpoint: str, expected_statuses=(200,), **kwargs):
    token = _get_graph_token()

    headers = kwargs.pop("headers", {})
    headers["Authorization"] = f"Bearer {token}"

    response = requests.request(
        method,
        f"{GRAPH_BASE}{endpoint}",
        headers=headers,
        timeout=120,
        **kwargs,
    )

    if response.status_code not in expected_statuses:
        raise RuntimeError(
            f"Microsoft Graph request failed.\n"
            f"Method: {method}\n"
            f"Endpoint: {endpoint}\n"
            f"Status: {response.status_code}\n"
            f"Response body:\n{response.text[:5000]}"
        )

    if response.status_code == 204:
        return None

    if response.content:
        return response.json()

    return None


def _graph_get_or_none(endpoint: str):
    response = requests.get(
        f"{GRAPH_BASE}{endpoint}",
        headers={"Authorization": f"Bearer {_get_graph_token()}"},
        timeout=120,
    )

    if response.status_code == 404:
        return None

    if response.status_code != 200:
        raise RuntimeError(
            f"Microsoft Graph GET failed.\n"
            f"Endpoint: {endpoint}\n"
            f"Status: {response.status_code}\n"
            f"Response body:\n{response.text[:5000]}"
        )

    return response.json()


def _init_sharepoint_graph() -> None:
    global _SP_SITE_ID, _SP_DRIVE_ID

    if _SP_SITE_ID and _SP_DRIVE_ID:
        return

    site = _graph_request(
        "GET",
        f"/sites/{SHAREPOINT_HOSTNAME}:{SITE_PATH}",
        expected_statuses=(200,),
    )
    _SP_SITE_ID = site["id"]

    drives = _graph_request(
        "GET",
        f"/sites/{_SP_SITE_ID}/drives",
        expected_statuses=(200,),
    )["value"]

    drive = next((d for d in drives if d.get("name") == DOCUMENT_LIBRARY_NAME), None)
    if not drive:
        available = [d.get("name") for d in drives]
        raise RuntimeError(
            f"Could not find document library drive '{DOCUMENT_LIBRARY_NAME}'. "
            f"Available drives: {available}"
        )

    _SP_DRIVE_ID = drive["id"]


def _get_drive_item_by_path(drive_relative_path: str):
    _init_sharepoint_graph()

    clean_path = _library_relative_path(drive_relative_path)

    if not clean_path:
        return _graph_request(
            "GET",
            f"/drives/{_SP_DRIVE_ID}/root",
            expected_statuses=(200,),
        )

    return _graph_get_or_none(
        f"/drives/{_SP_DRIVE_ID}/root:/{_encode_graph_path(clean_path)}"
    )


def folder_exists(folder_rel: str) -> bool:
    """True if the SharePoint folder exists (does not create it)."""
    item = _get_drive_item_by_path(folder_rel)
    return bool(item and item.get("folder") is not None)


def list_folder_files(folder_rel: str, recursive: bool = True) -> List[Dict[str, Any]]:
    """Return files (name + drive-relative path) under a folder, optionally recursing."""
    results: List[Dict[str, Any]] = []
    folder_rel = _library_relative_path(folder_rel)

    folder_item = _get_drive_item_by_path(folder_rel)
    if not folder_item or folder_item.get("folder") is None:
        return results

    def walk_folder(folder_item_id: str, current_rel: str):
        endpoint = f"/drives/{_SP_DRIVE_ID}/items/{folder_item_id}/children?$top=999"

        while endpoint:
            data = _graph_request("GET", endpoint, expected_statuses=(200,))
            for child in data.get("value", []):
                name = child.get("name", "")
                child_rel = f"{current_rel.rstrip('/')}/{name}" if current_rel else name

                if child.get("file") is not None:
                    results.append({
                        "name": name,
                        "url": child_rel,
                        "id": child.get("id", ""),
                    })

                elif recursive and child.get("folder") is not None and name != "Forms":
                    walk_folder(child["id"], child_rel)

            next_link = data.get("@odata.nextLink")
            if next_link:
                if next_link.startswith(GRAPH_BASE):
                    endpoint = next_link[len(GRAPH_BASE):]
                else:
                    raise RuntimeError(f"Unexpected Microsoft Graph nextLink: {next_link}")
            else:
                endpoint = None

    walk_folder(folder_item["id"], folder_rel)
    return results


def download_file(drive_relative_path: str, local_path: str) -> None:
    _init_sharepoint_graph()

    item = _get_drive_item_by_path(drive_relative_path)
    if not item:
        raise RuntimeError(f"SharePoint file not found: {drive_relative_path}")

    response = requests.get(
        f"{GRAPH_BASE}/drives/{_SP_DRIVE_ID}/items/{item['id']}/content",
        headers={"Authorization": f"Bearer {_get_graph_token()}"},
        timeout=120,
        allow_redirects=True,
    )

    if response.status_code != 200:
        raise RuntimeError(
            f"Failed downloading SharePoint file: {drive_relative_path}\n"
            f"Status: {response.status_code}\n"
            f"Response body:\n{response.text[:2000]}"
        )

    if not response.content:
        raise RuntimeError(f"Downloaded file is empty: {drive_relative_path}")

    os.makedirs(os.path.dirname(local_path), exist_ok=True)
    with open(local_path, "wb") as fh:
        fh.write(response.content)


def first_nonempty(*vals: Any) -> str:
    for v in vals:
        if v is not None and str(v).strip():
            return str(v).strip()
    return ""


# =========================================================
# SMAPPLY HELPERS
# =========================================================

_SMAPPLY_FILE_SIGNATURE_KEYS = {
    "file", "filename", "original_filename", "download_url", "url",
    "content_type", "mimetype", "size", "filesize", "extension",
}


def _smapply_headers() -> Dict[str, str]:
    access_token = wmill.get_variable(SMAPPLY_TOKEN_VAR)
    return {"Authorization": f"Bearer {access_token}"}


def _extract_task_files(task_json: Any) -> List[Dict[str, Any]]:
    """Locate uploaded-file objects in a task detail payload. Tries the known
    'data.files' location first, then a top-level 'files', then a recursive
    search for file-like objects. An EMPTY files list is a valid answer (no
    files on the task) and is returned as-is; the recursive search only runs
    when neither known location exists, so it can't mistake the task's stage
    object for a file."""
    if not isinstance(task_json, dict):
        return []

    data = task_json.get("data")
    if isinstance(data, dict) and isinstance(data.get("files"), list):
        candidates = data["files"]
    elif isinstance(task_json.get("files"), list):
        candidates = task_json["files"]
    else:
        candidates = None

    if candidates is not None:
        return [f for f in candidates if isinstance(f, dict) and f.get("id") is not None]

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

    walk(task_json)
    return found


def clear_smapply_task_files(smapply_appid: str, task_id: str) -> Dict[str, Any]:
    """Delete every file currently on the applicant-copy task for this
    application, so a re-run replaces rather than accumulates."""
    headers = _smapply_headers()
    task_url  = f"{BASE}/api/applications/{smapply_appid}/tasks/{task_id}/"
    files_url = f"{BASE}/api/applications/{smapply_appid}/tasks/{task_id}/files/"

    task_resp = requests.get(task_url, headers=headers, timeout=120)
    if not task_resp.ok:
        raise RuntimeError(
            f"SMApply get task failed {task_resp.status_code} "
            f"appid={smapply_appid} task={task_id}: {task_resp.text[:2000]}"
        )
    existing = _extract_task_files(task_resp.json())

    deleted = []
    for f in existing:
        file_id = f.get("id")
        if file_id is None:
            continue
        del_resp = requests.delete(f"{files_url}{file_id}/", headers=headers, timeout=120)
        # A 404 means the file is already gone; treat that as success.
        if del_resp.status_code not in (200, 202, 204, 404):
            raise RuntimeError(
                f"SMApply delete file {file_id} failed {del_resp.status_code} "
                f"appid={smapply_appid} task={task_id}: {del_resp.text[:2000]}"
            )
        deleted.append({"id": file_id, "filename": f.get("filename"),
                        "status": del_resp.status_code})

    return {"task_id": task_id, "found": len(existing), "deleted": deleted}


def upload_applicant_copy_to_task(smapply_appid: str, task_id: str,
                                  pdf_filepath: str, filename: str) -> Dict[str, Any]:
    """Post one PDF to the given SMApply task on the given application."""
    url = f"{BASE}/api/applications/{smapply_appid}/tasks/{task_id}/files/"
    with open(pdf_filepath, "rb") as f:
        file_data = f.read()
    response = requests.post(
        url,
        files={"file": (filename, file_data, "application/pdf")},
        headers=_smapply_headers(),
        timeout=300,
    )
    if not response.ok:
        raise RuntimeError(
            f"SMApply upload failed {response.status_code} task {task_id}, "
            f"file '{filename}': {response.text[:2000]}"
        )
    try:
        return response.json() if response.content else {}
    except ValueError:
        return {"raw": response.text}


def complete_smapply_task(smapply_appid: str, task_id: str,
                          completed_at: Optional[str] = None,
                          send_completed_at: bool = False) -> Dict[str, Any]:
    """
    Mark the SMApply task as complete (status 3). `apply-automations: full` is
    sent so SMApply fires the "Task has been marked as complete" automations.
    NOTE: if the task is already complete, SMApply may not treat this as a new
    completion, so those automations may not fire again.
    """
    url = f"{BASE}/api/applications/{smapply_appid}/tasks/{task_id}/"
    payload: Dict[str, Any] = {"status": 3}
    if send_completed_at:
        payload["completed_at"] = completed_at or datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    response = requests.patch(
        url,
        json=payload,
        headers={
            **_smapply_headers(),
            "Content-Type": "application/json",
            "apply-automations": "full",
        },
        timeout=120,
    )
    if not response.ok:
        try:
            detail = response.json()
        except ValueError:
            detail = response.text
        raise RuntimeError(f"{response.status_code} {response.reason}: {detail}")
    try:
        return response.json() if response.content else {}
    except ValueError:
        return {"raw": response.text}


# =========================================================
# APPLICANT COPY SELECTION
# =========================================================

_DESIG_RE = re.compile(r'(?:^|[_\-])(RN|LPN|RPN)(?:[_\-.]|$)', re.IGNORECASE)
_TS_RE = re.compile(r'(\d{8}_\d{6})')


def _parse_copy(name: str) -> Tuple[str, str]:
    """(designation, timestamp) parsed from an applicant-copy filename."""
    m = _DESIG_RE.search(name)
    desig = m.group(1).upper() if m else ""
    m2 = _TS_RE.search(name)
    ts = m2.group(1) if m2 else ""
    return desig, ts


def find_applicant_copies(reports_rel: str) -> List[Dict[str, str]]:
    """Every PDF sitting in credential reports/<RB>/Applicant Copy/."""
    found: List[Dict[str, str]] = []
    for entry in list_folder_files(reports_rel, recursive=True):
        url = entry.get("url") or ""
        name = entry.get("name") or ""
        if not url or not name.lower().endswith(".pdf"):
            continue
        rel = url[len(reports_rel):].lstrip("/")     # "<RB>/Applicant Copy/<file>.pdf"
        parts = rel.split("/")
        if len(parts) < 3:
            continue
        rb_folder, sub = parts[0], parts[1]
        if sub.strip().lower() != APPLICANT_COPY_SUBFOLDER.lower():
            continue
        found.append({"rb": rb_folder, "name": name, "url": url})
    return found


def keep_newest_per_group(found: List[Dict[str, str]]) -> Tuple[List[Dict[str, str]], List[str]]:
    """
    Keep only the newest copy per (regulator, designation), by the report
    timestamp in the filename. A copy with no parsable timestamp is treated as
    oldest. Returns (kept, skipped_descriptions).
    """
    latest_by_group: Dict[Tuple[str, str], Dict[str, str]] = {}
    skipped: List[str] = []
    for item in found:
        desig, ts = _parse_copy(item["name"])
        key = (item["rb"], desig)
        current = latest_by_group.get(key)
        if current is None or ts > _parse_copy(current["name"])[1]:
            if current is not None:
                skipped.append(f"{current['rb']}/{current['name']} (superseded by {item['name']})")
            latest_by_group[key] = item
        else:
            skipped.append(f"{item['rb']}/{item['name']} (older than {current['name']})")
    return list(latest_by_group.values()), skipped


# =========================================================
# ENTRYPOINT
# =========================================================

def main(reference_id: Optional[str] = None,
         appid: Any = None,
         additional_content: Optional[str] = None,
         program_id: Any = None,
         **kwargs: Any):
    """
    Field mapping (same webhook shape as the package builder):
        reference_id (or title) -> NNAS reference; names the SharePoint folder
                                   and prefixes each uploaded filename.
        appid (or webhook 'id') -> the NUMERIC SMApply application id used in
                                   the API URLs. Must be numeric; the old
                                   fallback to the NNAS reference was removed.
        program_id (or 'award') -> SMApply program id; selects TM or CES
                                   (folder + task). See PROGRAM_CONFIG.
        additional_content      -> used to cross-check the provider and to
                                   filter out regulators the applicant is no
                                   longer registered under.
    """
    title             = kwargs.get("title")
    application_title = kwargs.get("application_title")
    award             = kwargs.get("award")

    payload = parse_additional_content(additional_content)

    # ---- reference: names the folder and the uploaded files ----
    reference = first_nonempty(reference_id, title, application_title)
    if not reference:
        return {"status": "error",
                "message": "No reference_id/title to locate the SharePoint folder."}

    # ---- numeric SMApply application id ----
    smapply_appid = first_nonempty(
        (str(appid) if appid is not None else ""),
        (str(kwargs.get("id")) if kwargs.get("id") is not None else ""),
    )
    if not smapply_appid.isdigit():
        return {"status": "error",
                "message": (f"SMApply application id must be numeric; got "
                            f"'{smapply_appid or '(empty)'}'. Map appid to the webhook 'id'.")}

    # ---- program -> folder + task ----
    try:
        resolved_program_id, cfg = resolve_program(program_id, award, payload)
    except RuntimeError as e:
        return {"status": "error", "message": str(e),
                "debug": {"program_id_input": program_id, "award": award}}

    task_id      = cfg["task_id"]
    app_base_rel = _library_relative_path(cfg["base"].format(app_id=reference))
    reports_rel  = f"{app_base_rel}/{CREDENTIAL_REPORTS_SUBFOLDER}"

    debug: Dict[str, Any] = {
        "reference_id":  reference,
        "smapply_appid": smapply_appid,
        "program_id":    resolved_program_id,
        "provider":      cfg["provider"].upper(),
        "task_id":       task_id,
        "base_path":     app_base_rel,
        "reports_path":  reports_rel,
        "dry_run":       DRY_RUN,
    }

    if not folder_exists(app_base_rel):
        return {"status": "error",
                "message": f"Application folder not found: {app_base_rel}.",
                "debug": debug}
    if not folder_exists(reports_rel):
        return {"status": "no_reports_folder",
                "message": f"No '{CREDENTIAL_REPORTS_SUBFOLDER}' folder under {app_base_rel}.",
                "debug": debug}

    # ---- locate applicant copies ----
    found = find_applicant_copies(reports_rel)
    debug["applicant_copies_found"] = [f"{i['rb']}/{i['name']}" for i in found]

    # ---- drop regulators the applicant is no longer registered under ----
    current_abbrevs, current_rb_names = current_regulator_abbrevs(payload)
    debug["current_regulators_from_payload"] = current_rb_names
    debug["current_regulator_abbrevs"] = sorted(current_abbrevs)

    stale_regulator_skipped: List[str] = []
    if current_abbrevs:
        kept = []
        for item in found:
            if item["rb"].strip().upper() in current_abbrevs:
                kept.append(item)
            else:
                stale_regulator_skipped.append(f"{item['rb']}/{item['name']}")
        found = kept
    debug["stale_regulator_folders_skipped"] = stale_regulator_skipped

    # ---- newest copy per (regulator, designation) ----
    found, stale_skipped = keep_newest_per_group(found)
    debug["applicant_copies_used"] = [f"{i['rb']}/{i['name']}" for i in found]
    debug["applicant_copies_stale_skipped"] = stale_skipped

    if not found:
        return {"status": "no_applicant_copies",
                "message": "No applicant-copy PDFs found under any regulator folder.",
                "debug": debug}

    # ---- name each file: <reference>_<RB>_Applicant_report.pdf ----
    found.sort(key=lambda i: (i["rb"], i["name"]))
    counts = Counter(i["rb"] for i in found)
    running: Dict[str, int] = {}

    workdir = os.path.join("tmp", f"appcopy_{reference}_{datetime.utcnow():%Y%m%d_%H%M%S}")
    os.makedirs(workdir, exist_ok=True)

    try:
        # ---- phase 1: download EVERYTHING before touching SMApply ----
        prepared: List[Dict[str, str]] = []
        download_errors: List[str] = []
        for item in found:
            rb = item["rb"]
            running[rb] = running.get(rb, 0) + 1
            rb_safe = _ILLEGAL.sub(" ", rb).strip()
            if counts[rb] > 1:
                out_name = f"{reference}_{rb_safe}_{running[rb]}_Applicant_report.pdf"
            else:
                out_name = f"{reference}_{rb_safe}_Applicant_report.pdf"

            local = os.path.join(workdir, f"{rb_safe}_{running[rb]}.pdf")
            try:
                download_file(item["url"], local)
            except Exception as e:
                download_errors.append(f"{rb}/{item['name']}: {e}")
                continue
            prepared.append({"regulator": rb, "source": item["url"],
                             "local": local, "uploaded_as": out_name})

        if download_errors:
            debug["download_errors"] = download_errors
            return {"status": "error",
                    "message": ("One or more applicant copies failed to download; "
                                "SMApply was not modified."),
                    "debug": debug}

        # ---- phase 2: clear the task (real runs only) ----
        task_clear: Dict[str, Any] = {"attempted": False}
        if DRY_RUN:
            task_clear["skipped_reason"] = "dry run"
        else:
            task_clear["attempted"] = True
            try:
                task_clear["result"] = clear_smapply_task_files(smapply_appid, task_id)
            except Exception as e:
                task_clear["error"] = str(e)
                debug["task_clear"] = task_clear
                return {"status": "error",
                        "message": f"Failed to clear existing task files before re-upload: {e}",
                        "debug": debug}
        debug["task_clear"] = task_clear

        # ---- phase 3: upload ----
        uploaded: List[Dict[str, Any]] = []
        errors: List[str] = []
        for p in prepared:
            record = {"regulator": p["regulator"], "source": p["source"],
                      "uploaded_as": p["uploaded_as"], "task_id": task_id}
            if DRY_RUN:
                record["uploaded"] = False
                uploaded.append(record)
                continue
            try:
                record["smapply_response"] = upload_applicant_copy_to_task(
                    smapply_appid, task_id, p["local"], p["uploaded_as"])
                record["uploaded"] = True
            except Exception as e:
                record["uploaded"] = False
                errors.append(f"upload failed: {p['uploaded_as']}: {e}")
            uploaded.append(record)
    finally:
        shutil.rmtree(workdir, ignore_errors=True)

    uploaded_count = sum(1 for u in uploaded if u.get("uploaded"))

    # ---- mark the task complete once every copy is attached ----
    task_completion: Dict[str, Any] = {"attempted": False, "completed": False}
    if DRY_RUN:
        task_completion["skipped_reason"] = "dry run"
    elif errors:
        task_completion["skipped_reason"] = "upload errors; task left incomplete"
    elif uploaded_count == 0:
        task_completion["skipped_reason"] = "no files uploaded"
    else:
        task_completion["attempted"] = True
        try:
            task_completion["smapply_response"] = complete_smapply_task(smapply_appid, task_id)
            task_completion["completed"] = True
        except Exception as e:
            task_completion["error"] = str(e)
            errors.append(f"task completion failed: {e}")

    debug["errors"] = errors
    debug["task_completion"] = task_completion
    return {
        "status":         "ok" if not errors else "partial",
        "reference_id":   reference,
        "smapply_appid":  smapply_appid,
        "provider":       cfg["provider"].upper(),
        "task_id":        task_id,
        "dry_run":        DRY_RUN,
        "found_count":    len(found),
        "uploaded_count": uploaded_count,
        "task_cleared":   task_clear.get("result"),
        "task_completed": task_completion["completed"],
        "files":          uploaded,
        "debug":          debug,
    }