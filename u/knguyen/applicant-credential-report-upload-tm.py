# py: 3.12
# requirements:
# wmill
# requests

"""
NNAS Applicant-Copy Indexer  (self-contained Windmill script)
=============================================================

For one application, find every regulator's redacted applicant copy and push
it to a single SMApply task. Specifically:

  1. Resolve the SharePoint application folder  applications/tm/<reference_id>.
  2. Walk every regulator subfolder under "credential reports", and inside each
     one read the "Applicant Copy" subfolder:
         credential reports/<RB>/Applicant Copy/*.pdf
  3. Download each applicant-copy PDF.
  4. Upload each to SMApply task 2077121 on the SAME application, naming the
     file:  <reference_id>_<RB>_Applicant_report.pdf
     (a sequence number is inserted when one regulator has more than one copy,
     e.g. an RN and an LPN designation -> _1_ / _2_).
  5. Mark that task complete (PATCH {"status": 3}) once every applicant copy
     uploaded cleanly.

This is read-only against SharePoint; the only writes are the SMApply uploads
and the completion PATCH.
"""

import os
import re
import shutil
from collections import Counter
from datetime import datetime
from typing import Any, Dict, List, Optional
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

# Application folder template. CES applications live under applications/ces/;
# change this template (or parameterize) if this script ever needs to run on a
# non-TM provider.
APP_BASE_TEMPLATE            = "/tm/{app_id}"
CREDENTIAL_REPORTS_SUBFOLDER = "credential reports"
APPLICANT_COPY_SUBFOLDER     = "Applicant Copy"

# SMApply
BASE              = "https://nnas.smapply.io"
SMAPPLY_TASK_ID   = "2077121"
SMAPPLY_TOKEN_VAR = "u/epham2/SMAPPLY_ACCESS_TOKEN"

# SMApply task status integer for "complete".
SMAPPLY_STATUS_COMPLETE = 3

_ILLEGAL = re.compile(r'[\\/:*?"<>|]+')

_GRAPH_TOKEN = None
_SP_SITE_ID = None
_SP_DRIVE_ID = None


# =========================================================
# MICROSOFT GRAPH / SHAREPOINT HELPERS
# =========================================================

def get_graph_token() -> str:
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


def graph_request(method: str, endpoint: str, expected_statuses=(200,), **kwargs):
    token = get_graph_token()

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


def graph_get_or_none(endpoint: str):
    token = get_graph_token()

    response = requests.get(
        f"{GRAPH_BASE}{endpoint}",
        headers={"Authorization": f"Bearer {token}"},
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


def encode_graph_path(path: str) -> str:
    return quote(path.strip("/"), safe="/")


def init_sharepoint_graph() -> None:
    global _SP_SITE_ID, _SP_DRIVE_ID

    if _SP_SITE_ID and _SP_DRIVE_ID:
        return

    site = graph_request(
        "GET",
        f"/sites/{SHAREPOINT_HOSTNAME}:{SITE_PATH}",
        expected_statuses=(200,),
    )
    _SP_SITE_ID = site["id"]

    drives = graph_request(
        "GET",
        f"/sites/{_SP_SITE_ID}/drives",
        expected_statuses=(200,),
    )["value"]

    drive = next(
        (d for d in drives if d.get("name") == DOCUMENT_LIBRARY_NAME),
        None,
    )

    if not drive:
        available = [d.get("name") for d in drives]
        raise RuntimeError(
            f"Could not find document library drive '{DOCUMENT_LIBRARY_NAME}'. "
            f"Available drives: {available}"
        )

    _SP_DRIVE_ID = drive["id"]


def get_drive_item_by_path(drive_relative_path: str):
    init_sharepoint_graph()

    clean_path = drive_relative_path.strip("/")

    if not clean_path:
        return graph_request(
            "GET",
            f"/drives/{_SP_DRIVE_ID}/root",
            expected_statuses=(200,),
        )

    return graph_get_or_none(
        f"/drives/{_SP_DRIVE_ID}/root:/{encode_graph_path(clean_path)}"
    )


def folder_exists(folder_rel: str) -> bool:
    item = get_drive_item_by_path(folder_rel)
    return bool(item and item.get("folder") is not None)


def list_folder_files(folder_rel: str, recursive: bool = True) -> List[Dict[str, Any]]:
    results: List[Dict[str, Any]] = []

    folder_item = get_drive_item_by_path(folder_rel)
    if not folder_item or folder_item.get("folder") is None:
        return results

    def walk_folder(folder_item_id: str, current_rel: str):
        endpoint = f"/drives/{_SP_DRIVE_ID}/items/{folder_item_id}/children?$top=999"

        while endpoint:
            data = graph_request("GET", endpoint, expected_statuses=(200,))
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

    walk_folder(folder_item["id"], folder_rel.strip("/"))
    return results


def download_file(drive_relative_path: str, local_path: str) -> None:
    init_sharepoint_graph()

    item = get_drive_item_by_path(drive_relative_path)
    if not item:
        raise RuntimeError(f"SharePoint file not found: {drive_relative_path}")

    response = requests.get(
        f"{GRAPH_BASE}/drives/{_SP_DRIVE_ID}/items/{item['id']}/content",
        headers={"Authorization": f"Bearer {get_graph_token()}"},
        timeout=120,
        allow_redirects=True,
    )

    if response.status_code != 200:
        raise RuntimeError(
            f"Failed downloading SharePoint file: {drive_relative_path}\n"
            f"Status: {response.status_code}\n"
            f"Response body:\n{response.text[:2000]}"
        )

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

def _safe_body(response: requests.Response) -> Any:
    """SMApply 500s return HTML, not JSON. Always surface something readable."""
    if not response.content:
        return {}
    try:
        return response.json()
    except ValueError:
        return {"text": response.text[:3000]}


def upload_applicant_copy_to_task(smapply_appid: str, pdf_filepath: str, filename: str) -> Dict[str, Any]:
    """Post one PDF to the configured SMApply task on the given application."""
    url = f"{BASE}/api/applications/{smapply_appid}/tasks/{SMAPPLY_TASK_ID}/files/"
    access_token = wmill.get_variable(SMAPPLY_TOKEN_VAR)
    with open(pdf_filepath, "rb") as f:
        file_data = f.read()
    response = requests.post(
        url,
        files={"file": (filename, file_data, "application/octet-stream")},
        headers={"Authorization": f"Bearer {access_token}"},
        timeout=120,
    )
    if response.status_code not in (200, 201, 202, 204):
        raise RuntimeError(
            f"SMApply file upload failed.\n"
            f"URL: {url}\n"
            f"Filename: {filename}\n"
            f"Status: {response.status_code}\n"
            f"Response body:\n{response.text[:3000]}"
        )
    return _safe_body(response)


def mark_task_complete(smapply_appid: str) -> Dict[str, Any]:
    """
    Mark the applicant-copies task complete.

    Status-only PATCH: sending anything alongside 'status' (notably the
    read-only 'completed_at') triggers a 500. The apply-automations header is
    required for any completion automation on the task to fire.
    """
    url = f"{BASE}/api/applications/{smapply_appid}/tasks/{SMAPPLY_TASK_ID}/"
    access_token = wmill.get_variable(SMAPPLY_TOKEN_VAR)

    response = requests.patch(
        url,
        json={"status": SMAPPLY_STATUS_COMPLETE},
        headers={
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
            "apply-automations": "full",
        },
        timeout=60,
    )

    body = _safe_body(response)

    if response.status_code not in (200, 201, 202, 204):
        raise RuntimeError(
            f"SMApply mark-complete PATCH failed.\n"
            f"URL: {url}\n"
            f"Status: {response.status_code}\n"
            f"Response body:\n{response.text[:3000]}"
        )

    return {"http_status": response.status_code, "body": body}


# =========================================================
# ENTRYPOINT
# =========================================================

def main(reference_id: Optional[str] = None,
         appid: Any = None,
         dry_run: bool = False,
         mark_complete: bool = True,
         force_complete: bool = False,
         **kwargs: Any):
    """
    Field mapping (same webhook shape as the package builder):
        reference_id (or title) -> NNAS-... ; names the SharePoint folder
                                    applications/tm/<reference_id> AND prefixes
                                    each uploaded filename.
        appid                   -> the NUMERIC SMApply application id used in the
                                    upload URL. Falls back to the webhook 'id' /
                                    'appid' key. If only the NNAS reference is
                                    available it is used as a last resort (some
                                    SMApply instances accept it) and flagged in
                                    debug.

    dry_run:        when True, applicant copies are still located and downloaded,
                    but nothing is posted to SMApply and the task is not completed.
    mark_complete:  when True (default), PATCH the task to status 3 after the
                    uploads.
    force_complete: when True, complete the task even if some uploads failed.
                    Default behaviour is to leave the task open on any error so
                    the run can be retried.
    """
    title             = kwargs.get("title")
    application_title = kwargs.get("application_title")

    # Windmill coerces webhook fields to strings; accept the usual truthy spellings.
    def _as_bool(value: Any, default: bool) -> bool:
        if value is None:
            return default
        if isinstance(value, bool):
            return value
        return str(value).strip().lower() in ("1", "true", "yes", "y", "on")

    dry_run        = _as_bool(dry_run, False)
    mark_complete  = _as_bool(mark_complete, True)
    force_complete = _as_bool(force_complete, False)

    # ---- reference_id: names the folder and the uploaded files ----
    reference = first_nonempty(reference_id, title, application_title)
    if not reference:
        return {"status": "error",
                "message": "No reference_id/title to locate the SharePoint folder."}

    # ---- numeric SMApply application id for the upload URL ----
    smapply_appid = first_nonempty(
        (str(appid) if appid is not None else ""),
        (str(kwargs.get("id")) if kwargs.get("id") is not None else ""),
        (str(kwargs.get("appid")) if kwargs.get("appid") is not None else ""),
    )
    appid_fallback = False
    if not smapply_appid:
        smapply_appid = reference   # last resort
        appid_fallback = True

    app_base_path = APP_BASE_TEMPLATE.format(app_id=reference)
    reports_path  = f"{app_base_path}/{CREDENTIAL_REPORTS_SUBFOLDER}"

    debug: Dict[str, Any] = {
        "reference_id": reference,
        "smapply_appid": smapply_appid,
        "smapply_appid_fallback_to_reference": appid_fallback,
        "task_id": SMAPPLY_TASK_ID,
        "base_path": app_base_path,
        "reports_path": reports_path,
    }

    if not folder_exists(app_base_path):
        return {"status": "error",
                "message": f"Application folder not found: {app_base_path}.",
                "debug": debug}
    if not folder_exists(reports_path):
        return {"status": "no_reports_folder",
                "message": f"No '{CREDENTIAL_REPORTS_SUBFOLDER}' folder under {app_base_path}.",
                "debug": debug}

    # ---- find every applicant-copy PDF under credential reports/<RB>/Applicant Copy/ ----
    all_files = list_folder_files(reports_path, recursive=True)
    found: List[Dict[str, str]] = []
    for entry in all_files:
        url = entry.get("url") or ""
        name = entry.get("name") or ""
        if not url or not name.lower().endswith(".pdf"):
            continue
        rel = url[len(reports_path):].lstrip("/")     # "<RB>/Applicant Copy/<file>.pdf"
        parts = rel.split("/")
        if len(parts) < 3:
            continue
        rb_folder, sub = parts[0], parts[1]
        if sub.strip().lower() != APPLICANT_COPY_SUBFOLDER.lower():
            continue
        found.append({"rb": rb_folder, "name": name, "url": url})

    debug["applicant_copies_found"] = [f"{i['rb']}/{i['name']}" for i in found]

    if not found:
        return {"status": "no_applicant_copies",
                "message": "No applicant-copy PDFs found under any regulator folder.",
                "debug": debug}

    # ---- name each file: <reference>_<RB>_Applicant_report.pdf ----
    # When a regulator has more than one copy (e.g. RN + LPN), insert a 1-based
    # sequence so the names stay distinct.
    found.sort(key=lambda i: (i["rb"], i["name"]))
    counts = Counter(i["rb"] for i in found)
    running: Dict[str, int] = {}

    workdir = os.path.join("tmp", f"appcopy_{reference}_{datetime.utcnow():%Y%m%d_%H%M%S}")
    os.makedirs(workdir, exist_ok=True)

    uploaded: List[Dict[str, Any]] = []
    errors: List[str] = []

    try:
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
                errors.append(f"download failed: {item['rb']}/{item['name']}: {e}")
                continue

            record = {
                "regulator": rb,
                "source": item["url"],
                "uploaded_as": out_name,
                "task_id": SMAPPLY_TASK_ID,
            }

            if dry_run:
                record["uploaded"] = False
                uploaded.append(record)
                continue

            try:
                resp = upload_applicant_copy_to_task(smapply_appid, local, out_name)
                record["uploaded"] = True
                record["smapply_response"] = resp
                uploaded.append(record)
            except Exception as e:
                errors.append(f"upload failed: {out_name}: {e}")
    finally:
        shutil.rmtree(workdir, ignore_errors=True)

    uploaded_count = sum(1 for u in uploaded if u.get("uploaded"))

    # ---- mark the applicant-copies task complete ----
    # Separate PATCH from the uploads: SMApply rejects a combined write + status
    # change with a 500, so completion always runs as its own request.
    completion: Dict[str, Any] = {
        "requested": mark_complete,
        "completed": False,
        "skipped_reason": None,
        "response": None,
    }

    if not mark_complete:
        completion["skipped_reason"] = "mark_complete=False"
    elif dry_run:
        completion["skipped_reason"] = "dry_run"
    elif uploaded_count == 0:
        completion["skipped_reason"] = "no files uploaded"
    elif errors and not force_complete:
        completion["skipped_reason"] = (
            f"{len(errors)} error(s) during upload; task left open "
            "(pass force_complete=true to override)"
        )
    else:
        try:
            completion["response"] = mark_task_complete(smapply_appid)
            completion["completed"] = True
        except Exception as e:
            errors.append(f"mark complete failed: {e}")
            completion["error"] = str(e)

    debug["errors"] = errors
    debug["dry_run"] = dry_run
    debug["mark_complete"] = mark_complete
    debug["force_complete"] = force_complete

    return {
        "status": "ok" if not errors else "partial",
        "reference_id": reference,
        "smapply_appid": smapply_appid,
        "task_id": SMAPPLY_TASK_ID,
        "dry_run": dry_run,
        "found_count": len(found),
        "uploaded_count": uploaded_count,
        "task_completed": completion["completed"],
        "completion": completion,
        "files": uploaded,
        "debug": debug,
    }