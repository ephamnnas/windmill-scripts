# py: 3.12
# requirements:
# wmill
# requests

"""
NNAS Applicant-Copy Indexer  (self-contained Windmill script)
=============================================================

For one application, find every regulator's redacted applicant copy and push
it to a single SMApply task. Specifically:

  1. Resolve the SharePoint application folder  applications/ces/<reference_id>.
  2. Walk every regulator subfolder under "credential reports", and inside each
     one read the "Applicant Copy" subfolder:
         credential reports/<RB>/Applicant Copy/*.pdf
  3. Download each applicant-copy PDF.
  4. Upload each to SMApply task 2077121 on the SAME application, naming the
     file:  <reference_id>_<RB>_Applicant_report.pdf
     (a sequence number is inserted when one regulator has more than one copy,
     e.g. an RN and an LPN designation -> _1_ / _2_).

This is read-only against SharePoint; the only write is the SMApply upload.
"""

import os
import re
import shutil
from collections import Counter
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional, Tuple
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

# Application folder template. CES applications live under Applications/CES/;
# in Graph this path is relative to the Applications document library root.
APP_BASE_TEMPLATE            = "CES/{app_id}"
CREDENTIAL_REPORTS_SUBFOLDER = "credential reports"
APPLICANT_COPY_SUBFOLDER     = "Applicant Copy"

# SMApply
BASE              = "https://nnas.smapply.io"
SMAPPLY_TASK_ID   = "2077784"
SMAPPLY_TOKEN_VAR = "u/epham2/SMAPPLY_ACCESS_TOKEN"

_ILLEGAL = re.compile(r'[\\/:*?"<>|]+')

_GRAPH_TOKEN = None
_SP_SITE_ID = None
_SP_DRIVE_ID = None


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


def ctx():
    return None


def folder_exists(sp, folder_rel: str) -> bool:
    """True if the SharePoint folder exists (does not create it)."""
    item = _get_drive_item_by_path(folder_rel)
    return bool(item and item.get("folder") is not None)


def list_folder_files(sp, folder_rel: str, recursive: bool = True) -> List[Dict[str, Any]]:
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


def download_file(sp, drive_relative_path: str, local_path: str) -> None:
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

    os.makedirs(os.path.dirname(local_path), exist_ok=True)
    with open(local_path, "wb") as fh:
        fh.write(response.content)


def first_nonempty(*vals: Any) -> str:
    for v in vals:
        if v is not None and str(v).strip():
            return str(v).strip()
    return ""


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
    )
    response.raise_for_status()
    try:
        return response.json() if response.content else {}
    except ValueError:
        return {"raw": response.text}


def complete_smapply_task(smapply_appid: str,
                          completed_at: Optional[str] = None,
                          send_completed_at: bool = False) -> Dict[str, Any]:
    """
    Mark the SMApply task as complete (status 3) on the given application.
    `apply-automations: full` is sent so SMApply fires the "Task has been marked
    as complete" automations.

    By default only {"status": 3} is sent and SMApply stamps the completion time
    itself. Some instances reject a client-supplied completed_at (read-only /
    format), which is a common source of 400s -- set send_completed_at=True only
    if your instance accepts it.

    On a non-2xx response the API's validation detail is surfaced in the raised
    error (instead of a bare status code) so the cause is visible in debug.
    """
    url = f"{BASE}/api/applications/{smapply_appid}/tasks/{SMAPPLY_TASK_ID}/"
    access_token = wmill.get_variable(SMAPPLY_TOKEN_VAR)
    payload: Dict[str, Any] = {"status": 3}
    if send_completed_at:
        payload["completed_at"] = completed_at or datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    response = requests.patch(
        url,
        json=payload,
        headers={
            "Authorization": f"Bearer {access_token}",
            "apply-automations": "full",
        },
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
# ENTRYPOINT
# =========================================================

def main(reference_id: Optional[str] = None,
         appid: Any = None,
         dry_run: bool = False,
         **kwargs: Any):
    """
    Field mapping (same webhook shape as the package builder):
        reference_id (or title) -> NNAS-... ; names the SharePoint folder
                                    applications/ces/<reference_id> AND prefixes
                                    each uploaded filename.
        appid                   -> the NUMERIC SMApply application id used in the
                                    upload URL. Falls back to the webhook 'id' /
                                    'appid' key. If only the NNAS reference is
                                    available it is used as a last resort (some
                                    SMApply instances accept it) and flagged in
                                    debug.

    dry_run: when True, applicant copies are still located and downloaded, but
    nothing is posted to SMApply. Use it to verify which files WOULD be sent.
    """
    title             = kwargs.get("title")
    application_title = kwargs.get("application_title")

    sp = ctx()

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

    app_base_rel = _library_relative_path(APP_BASE_TEMPLATE.format(app_id=reference))
    reports_rel  = f"{app_base_rel}/{CREDENTIAL_REPORTS_SUBFOLDER}"

    debug: Dict[str, Any] = {
        "reference_id": reference,
        "smapply_appid": smapply_appid,
        "smapply_appid_fallback_to_reference": appid_fallback,
        "task_id": SMAPPLY_TASK_ID,
        "base_path": app_base_rel,
        "reports_path": reports_rel,
    }

    if not folder_exists(sp, app_base_rel):
        return {"status": "error",
                "message": f"Application folder not found: {app_base_rel}.",
                "debug": debug}
    if not folder_exists(sp, reports_rel):
        return {"status": "no_reports_folder",
                "message": f"No '{CREDENTIAL_REPORTS_SUBFOLDER}' folder under {app_base_rel}.",
                "debug": debug}

    # ---- find every applicant-copy PDF under credential reports/<RB>/Applicant Copy/ ----
    all_files = list_folder_files(sp, reports_rel, recursive=True)
    found: List[Dict[str, str]] = []
    for entry in all_files:
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
                download_file(sp, item["url"], local)
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

    # ---- mark the task complete once every copy is attached ----
    # Only fires on a real run where at least one file uploaded and nothing
    # failed; a partial/empty result leaves the task incomplete so it stays
    # visible as needing attention.
    task_completion: Dict[str, Any] = {"attempted": False, "completed": False}
    if dry_run:
        task_completion["skipped_reason"] = "dry run"
    elif errors:
        task_completion["skipped_reason"] = "upload errors; task left incomplete"
    elif uploaded_count == 0:
        task_completion["skipped_reason"] = "no files uploaded"
    else:
        task_completion["attempted"] = True
        try:
            task_completion["smapply_response"] = complete_smapply_task(smapply_appid)
            task_completion["completed"] = True
        except Exception as e:
            task_completion["error"] = str(e)
            errors.append(f"task completion failed: {e}")

    debug["errors"] = errors
    debug["dry_run"] = dry_run
    debug["task_completion"] = task_completion
    return {
        "status": "ok" if not errors else "partial",
        "reference_id": reference,
        "smapply_appid": smapply_appid,
        "task_id": SMAPPLY_TASK_ID,
        "dry_run": dry_run,
        "found_count": len(found),
        "uploaded_count": uploaded_count,
        "task_completed": task_completion["completed"],
        "files": uploaded,
        "debug": debug,
    }
