# py: 3.12
# requirements:
# wmill
# requests

"""
NNAS Applicant Copy Upload  (Windmill script)
=============================================

For one application:
  1. Locates the SharePoint application folder from the reference_id alone:
     NNAS-T-... is a TM application (tm/<id>), NNAS-C-... is CES (ces/<id>).
  2. Finds the applicant-copy reports that nnas_sftp_package.py writes to
     "credential reports/<RB>/Applicant Copy/*_applicant.pdf".
  3. Downloads them and uploads each one to the SMApply task named
     "Applicant Credential Report Copies".

By default only the NEWEST applicant copy per (regulator, designation) is
uploaded, judged by the YYYYMMDD_HHMMSS timestamp in the filename, so older
superseded copies are not sent to the applicant. Set include_all=True to upload
every applicant copy found.

!!! CONFIRM BEFORE FIRST LIVE RUN !!!
The SMApply endpoint used to attach a file to a task (SMAPPLY_UPLOAD_PATH below)
has NOT been verified against the SMApply API documentation. dry_run defaults to
True so you can check everything else first. See the notes beside
SMAPPLY_UPLOAD_PATH and smapply_upload_file().
"""

import os
import re
import shutil
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import quote

import wmill
import requests


# =========================================================
# SMAPPLY CONFIG
# =========================================================
SMAPPLY_API_BASE      = "https://portal.nnas.ca/api"
SMAPPLY_API_TOKEN_VAR = "u/epham2/SMAPPLY_ACCESS_TOKEN"
TARGET_TASK_NAME      = "Applicant Credential Report Copies"

# UNVERIFIED: the route that accepts a file for a task. {app_id} = SMApply
# internal application id, {task_id} = the task's id from the tasks listing.
# To confirm: open https://portal.nnas.ca/api/applications/<id>/tasks/<task_id>/
# in a browser while signed in (the API is browsable) and check the allowed
# methods and any upload/files link, or check the SMApply API documentation.
SMAPPLY_UPLOAD_PATH      = "/applications/{app_id}/tasks/{task_id}/upload/"
SMAPPLY_UPLOAD_FIELD     = "file"     # multipart field name for the file
HTTP_TIMEOUT             = 60


# =========================================================
# SHAREPOINT / MICROSOFT GRAPH CONFIG  (matches nnas_sftp_package.py)
# =========================================================
TENANT_ID     = wmill.get_variable("u/epham2/AZURE_TENANT_ID")
CLIENT_ID     = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_ID")
CLIENT_SECRET = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_SECRET")

GRAPH_BASE            = "https://graph.microsoft.com/v1.0"
SHAREPOINT_HOSTNAME   = "nnasca.sharepoint.com"
SITE_PATH             = "/sites/NNAS-Portal-Sharepoint"
DOCUMENT_LIBRARY_NAME = "Applications"

PROVIDER_FOLDER  = {"tm": "tm", "ces": "ces"}

CREDENTIAL_REPORTS_SUBFOLDER = "credential reports"
APPLICANT_COPY_SUBFOLDER     = "Applicant Copy"
APPLICANT_COPY_SUFFIX        = "_applicant.pdf"

REPORT_DESIGNATIONS = ("RN", "LPN", "RPN")
TS_RE = re.compile(r"(\d{8})_(\d{6})")

_TOKEN: Optional[str] = None
_DRIVE_ID: Optional[str] = None


# =========================================================
# SMALL HELPERS
# =========================================================

def safe_str(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, dict):
        for k in ("label", "value", "name", "text"):
            if v.get(k) is not None:
                return str(v[k]).strip()
        return ""
    return str(v).strip()


# Reference ID prefix -> provider folder: NNAS-T-... = TM, NNAS-C-... = CES.
REF_PROVIDER_RE = re.compile(r"^NNAS-([A-Z])-", re.IGNORECASE)
REF_PROVIDER_MAP = {"T": "tm", "C": "ces"}


def provider_from_reference(reference_id: str) -> Tuple[str, Optional[str]]:
    """
    (provider, error). Reads the letter after "NNAS-" in the reference ID.
    An ID without a letter (older format) or with an unknown letter returns an
    error rather than guessing a folder.
    """
    m = REF_PROVIDER_RE.match((reference_id or "").strip())
    if not m:
        return "", (f"Cannot tell TM or CES from reference ID '{reference_id}'. "
                    f"Expected NNAS-T-... or NNAS-C-..., or pass provider explicitly.")
    letter = m.group(1).upper()
    if letter not in REF_PROVIDER_MAP:
        return "", (f"Unknown program letter '{letter}' in reference ID '{reference_id}'. "
                    f"Expected T (TM) or C (CES), or pass provider explicitly.")
    return REF_PROVIDER_MAP[letter], None


def report_key(filename: str) -> Tuple[str, str]:
    """(regulator token, designation) taken from the filename.
    e.g. NNAS-C-..._CLHA_LPN_20261020_225822_applicant.pdf -> ('CLHA', 'LPN')."""
    stem = os.path.splitext(filename)[0].upper()
    tokens = [t for t in re.split(r"[-_.\s]+", stem) if t]
    desig = next((t for t in tokens if t in REPORT_DESIGNATIONS), "")
    # the regulator token is the one just before the designation, when present
    rb = ""
    if desig and desig in tokens:
        i = tokens.index(desig)
        if i > 0:
            rb = tokens[i - 1]
    return rb, desig


def report_timestamp(filename: str) -> str:
    m = TS_RE.search(filename)
    return f"{m.group(1)}{m.group(2)}" if m else ""


# =========================================================
# MICROSOFT GRAPH
# =========================================================

def graph_token() -> str:
    global _TOKEN
    if _TOKEN:
        return _TOKEN
    if not (TENANT_ID and CLIENT_ID and CLIENT_SECRET):
        raise RuntimeError("Missing SharePoint/Azure Windmill variables.")
    r = requests.post(
        f"https://login.microsoftonline.com/{TENANT_ID.strip()}/oauth2/v2.0/token",
        data={"client_id": CLIENT_ID.strip(), "client_secret": CLIENT_SECRET.strip(),
              "scope": "https://graph.microsoft.com/.default",
              "grant_type": "client_credentials"},
        timeout=60)
    if r.status_code != 200:
        raise RuntimeError(f"Graph token request failed [{r.status_code}]: {r.text[:500]}")
    _TOKEN = r.json()["access_token"]
    return _TOKEN


def graph_get(url_or_path: str, allow_404: bool = False) -> Optional[Dict[str, Any]]:
    url = url_or_path if url_or_path.startswith("http") else f"{GRAPH_BASE}{url_or_path}"
    r = requests.get(url, headers={"Authorization": f"Bearer {graph_token()}"}, timeout=120)
    if allow_404 and r.status_code == 404:
        return None
    if r.status_code != 200:
        raise RuntimeError(f"Graph GET failed [{r.status_code}] {url}: {r.text[:500]}")
    return r.json()


def drive_id() -> str:
    global _DRIVE_ID
    if _DRIVE_ID:
        return _DRIVE_ID
    site = graph_get(f"/sites/{SHAREPOINT_HOSTNAME}:{SITE_PATH}")
    drives = graph_get(f"/sites/{site['id']}/drives")["value"]
    d = next((x for x in drives if x.get("name") == DOCUMENT_LIBRARY_NAME), None)
    if not d:
        raise RuntimeError(f"Document library '{DOCUMENT_LIBRARY_NAME}' not found.")
    _DRIVE_ID = d["id"]
    return _DRIVE_ID


def item_by_path(path: str) -> Optional[Dict[str, Any]]:
    p = quote(path.strip("/"), safe="/")
    return graph_get(f"/drives/{drive_id()}/root:/{p}", allow_404=True)


def walk_files(folder_path: str) -> List[Dict[str, Any]]:
    """All files under folder_path (recursive), with their library path."""
    out: List[Dict[str, Any]] = []
    root = item_by_path(folder_path)
    if not root or root.get("folder") is None:
        return out

    def walk(item_id: str, path: str):
        url = f"/drives/{drive_id()}/items/{item_id}/children?$top=999"
        while url:
            data = graph_get(url)
            for c in data.get("value", []):
                cpath = f"{path}/{c.get('name', '')}"
                if c.get("file") is not None:
                    out.append({"name": c["name"], "id": c["id"], "path": cpath,
                                "parent": path})
                elif c.get("folder") is not None:
                    walk(c["id"], cpath)
            url = data.get("@odata.nextLink")

    walk(root["id"], folder_path.strip("/"))
    return out


def download(item_id: str, local_path: str) -> None:
    r = requests.get(f"{GRAPH_BASE}/drives/{drive_id()}/items/{item_id}/content",
                     headers={"Authorization": f"Bearer {graph_token()}"},
                     timeout=300, allow_redirects=True)
    if r.status_code != 200:
        raise RuntimeError(f"Download failed [{r.status_code}]: {r.text[:300]}")
    if not r.content:
        raise RuntimeError("Downloaded file is empty.")
    os.makedirs(os.path.dirname(local_path), exist_ok=True)
    with open(local_path, "wb") as fh:
        fh.write(r.content)


# =========================================================
# SMAPPLY
# =========================================================

def sm_headers(token: str) -> Dict[str, str]:
    return {"Authorization": f"Bearer {token}", "Accept": "application/json"}


def smapply_find_application_id(reference_id: str, token: str) -> Optional[str]:
    """
    Best-effort lookup of the SMApply internal id from the reference_id.
    UNVERIFIED filter parameter: tries ?search=; every hit is checked for an
    exact title/reference match, so a loose filter can never pick the wrong
    application. Pass `id` from the webhook to skip this entirely.
    """
    try:
        r = requests.get(f"{SMAPPLY_API_BASE}/applications/",
                         headers=sm_headers(token),
                         params={"search": reference_id}, timeout=HTTP_TIMEOUT)
        if not r.ok:
            return None
        for app in (r.json() or {}).get("results", []):
            if reference_id in (safe_str(app.get("title")), safe_str(app.get("reference_id"))):
                return str(app.get("id"))
    except Exception:
        return None
    return None


def smapply_find_task(app_id: str, token: str) -> Optional[Dict[str, Any]]:
    """Find the target task by name (case-insensitive) on the application."""
    url = f"{SMAPPLY_API_BASE}/applications/{app_id}/tasks/"
    for _ in range(20):
        r = requests.get(url, headers=sm_headers(token), timeout=HTTP_TIMEOUT)
        r.raise_for_status()
        data = r.json()
        for t in data.get("results", []):
            if safe_str(t.get("name")).lower() == TARGET_TASK_NAME.lower():
                return t
        nxt = data.get("next")
        if not nxt:
            break
        url = nxt.replace("http://", "https://", 1)
    return None


def smapply_upload_file(app_id: str, task_id: str, local_path: str, token: str) -> Dict[str, Any]:
    """
    Attach one file to the task. UNVERIFIED route and field name: see
    SMAPPLY_UPLOAD_PATH. The file is read into memory so a failure never sends
    a half-read stream.
    """
    with open(local_path, "rb") as fh:
        content = fh.read()
    url = SMAPPLY_API_BASE + SMAPPLY_UPLOAD_PATH.format(app_id=app_id, task_id=task_id)
    r = requests.post(url, headers=sm_headers(token),
                      files={SMAPPLY_UPLOAD_FIELD: (os.path.basename(local_path),
                                                    content, "application/pdf")},
                      timeout=HTTP_TIMEOUT)
    try:
        body = r.json()
    except ValueError:
        body = {"raw": r.text[:500]}
    if not r.ok:
        raise RuntimeError(f"SMApply upload failed [{r.status_code}] {url}: {body}")
    return body if isinstance(body, dict) else {"raw": body}


# =========================================================
# ENTRYPOINT
# =========================================================

def main(reference_id: Optional[str] = None,
         id: Any = None,
         include_all: bool = False,
         dry_run: bool = True,
         provider: Optional[str] = None,
         **kwargs: Any):
    """
    reference_id -> ApplicationID / SharePoint folder (falls back to title).
                    The program is read from it: NNAS-T-... = TM, NNAS-C-... = CES.
    id           -> SMApply INTERNAL application id (recommended; looked up from
                    reference_id if not given)
    include_all  -> upload every applicant copy, not just the newest per
                    regulator + designation
    dry_run      -> list what would be uploaded; no SMApply writes
    provider     -> optional override ("tm" or "ces"), only needed for reference
                    IDs without a T/C letter
    """
    ref = (reference_id or kwargs.get("title") or "").strip()
    if not ref:
        return {"status": "error", "message": "Missing reference_id."}

    override = (provider or "").strip().lower()
    if override:
        if override not in PROVIDER_FOLDER:
            return {"status": "error",
                    "message": f"Unknown provider '{provider}'. Expected tm or ces."}
        provider = override
    else:
        provider, perr = provider_from_reference(ref)
        if perr:
            return {"status": "error", "message": perr}

    base = f"{PROVIDER_FOLDER[provider]}/{ref}"
    if not item_by_path(base):
        return {"status": "error",
                "message": f"Application folder not found: {DOCUMENT_LIBRARY_NAME}/{base}"}

    # ---- collect applicant copies ----
    all_files = walk_files(f"{base}/{CREDENTIAL_REPORTS_SUBFOLDER}")
    copies = [f for f in all_files
              if f["name"].lower().endswith(APPLICANT_COPY_SUFFIX)
              and os.path.basename(f["parent"]).lower() == APPLICANT_COPY_SUBFOLDER.lower()]

    if include_all:
        selected = sorted(copies, key=lambda f: f["name"])
        superseded: List[str] = []
    else:
        newest: Dict[Tuple[str, str], Dict[str, Any]] = {}
        for f in copies:
            k = report_key(f["name"])
            if k not in newest or report_timestamp(f["name"]) > report_timestamp(newest[k]["name"]):
                newest[k] = f
        selected = sorted(newest.values(), key=lambda f: f["name"])
        chosen = {f["id"] for f in selected}
        superseded = sorted(f["name"] for f in copies if f["id"] not in chosen)

    debug = {"provider": provider.upper(), "base_path": base,
             "applicant_copies_found": [f["name"] for f in copies],
             "superseded_not_uploaded": superseded,
             "upload_route": SMAPPLY_UPLOAD_PATH}

    if not selected:
        return {"status": "no_files",
                "message": "No applicant copies found under credential reports/*/Applicant Copy/.",
                "debug": debug}

    # ---- resolve SMApply application + task ----
    token = wmill.get_variable(SMAPPLY_API_TOKEN_VAR)
    if not token:
        return {"status": "error", "message": f"Missing Windmill variable {SMAPPLY_API_TOKEN_VAR}."}

    app_id = str(id) if id not in (None, "") else smapply_find_application_id(ref, token)
    if not app_id:
        return {"status": "error",
                "message": (f"Could not find the SMApply application for '{ref}'. "
                            f"Pass the internal SMApply id as `id`."),
                "debug": debug}

    task = smapply_find_task(app_id, token)
    if not task:
        return {"status": "error",
                "message": f"Task '{TARGET_TASK_NAME}' not found on SMApply application {app_id}.",
                "debug": debug}
    task_id = str(task.get("id"))
    debug.update({"smapply_application_id": app_id, "task_id": task_id})

    if dry_run:
        return {"status": "dry_run", "reference_id": ref,
                "would_upload": [f["name"] for f in selected],
                "to_task": f"{TARGET_TASK_NAME} (task {task_id})",
                "debug": debug}

    # ---- download + upload ----
    workdir = os.path.join("tmp", f"ac_{ref}_{datetime.utcnow():%Y%m%d_%H%M%S}")
    uploaded, failed = [], []
    try:
        for f in selected:
            local = os.path.join(workdir, f["name"])
            try:
                download(f["id"], local)
                resp = smapply_upload_file(app_id, task_id, local, token)
                uploaded.append({"file": f["name"], "response": resp})
            except Exception as e:
                failed.append({"file": f["name"], "error": str(e)})
    finally:
        shutil.rmtree(workdir, ignore_errors=True)

    return {
        "status": "ok" if not failed else ("partial" if uploaded else "error"),
        "reference_id": ref,
        "uploaded_count": len(uploaded),
        "failed_count": len(failed),
        "uploaded": uploaded,
        "failed": failed,
        "debug": debug,
    }