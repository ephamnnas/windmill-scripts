# py: 3.12
# requirements:
# office365-rest-python-client
# wmill
# requests

import os
import requests
import wmill

from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext


# =========================================================
# SHAREPOINT CONFIG
# =========================================================

SITE_URL      = "https://nnasca.sharepoint.com/sites/NNAS-Portal-Sharepoint"
CLIENT_ID     = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_ID")
CLIENT_SECRET = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_SECRET")

APPLICATIONS_BASE      = "applications/tm"
CRED_REPORTS_SUBFOLDER = "Credential reports"


# =========================================================
# TASK ID → (GROUP, SHORTNAME) ROUTING MAP
# =========================================================

TASK_ROUTE_MAP: Dict[int, Tuple[str, str]] = {
    # RN
    2072705: ("RN", "BCCNM"),
    2072706: ("RN", "CNO"),
    2072707: ("RN", "CRNA"),
    2072708: ("RN", "CRNM"),
    2072709: ("RN", "PEICNM"),
    2072710: ("RN", "NLCN"),
    2072711: ("RN", "CRNS"),
    2072712: ("RN", "CNNB"),
    2072713: ("RN", "NSCN"),
    # LPN
    2072714: ("LPN", "BCCNM"),
    2072715: ("LPN", "CNO"),
    2072716: ("LPN", "CLHA"),
    2072717: ("LPN", "CLPNM"),
    2072718: ("LPN", "CLPNS"),
    2072719: ("LPN", "PEICNM"),
    2072720: ("LPN", "NLCN"),
    2072721: ("LPN", "ANBLPN"),
    2072722: ("LPN", "NSCN"),
}

ALL_TASK_IDS = list(TASK_ROUTE_MAP.keys())


# =========================================================
# SHAREPOINT HELPERS
# =========================================================

def ctx() -> ClientContext:
    creds = ClientCredential(CLIENT_ID, CLIENT_SECRET)
    return ClientContext(SITE_URL).with_credentials(creds)


def _sp_url(folder_path: str) -> str:
    return f"/sites/NNAS-Portal-Sharepoint/{folder_path}"


def _ensure_sp_folder_tree(sp_ctx: ClientContext, folder_path: str):
    current = "/sites/NNAS-Portal-Sharepoint"
    for part in [p for p in folder_path.split("/") if p]:
        current = f"{current}/{part}"
        try:
            sp_ctx.web.get_folder_by_server_relative_url(current).get().execute_query()
        except Exception:
            sp_ctx.web.folders.add(current).execute_query()


def _demote_latest_in_folder(sp_ctx: ClientContext, folder_path: str):
    """Demote any existing *_latest.<ext> file to its plain timestamped name."""
    folder_url = _sp_url(folder_path)
    try:
        folder = sp_ctx.web.get_folder_by_server_relative_url(folder_url)
        files  = folder.files.get().execute_query()
    except Exception:
        return

    os.makedirs("tmp", exist_ok=True)
    for f in list(files):
        try:
            name = f.name
        except Exception:
            continue
        if "." not in name:
            continue
        stem, _, ext = name.rpartition(".")
        if not stem.endswith("_latest"):
            continue
        new_name = f"{stem[:-len('_latest')]}.{ext}"

        tmp_local = os.path.join("tmp", f"_demote_{new_name}")
        try:
            with open(tmp_local, "wb") as fh:
                f.download(fh).execute_query()
        except Exception:
            continue

        try:
            with open(tmp_local, "rb") as fh:
                content = fh.read()
            demoted = folder.upload_file(new_name, content).execute_query()
            li = demoted.listItemAllFields
            li.set_property("FileType", "Report")
            li.update()
            sp_ctx.execute_query()
        except Exception:
            continue

        try:
            f.delete_object().execute_query()
        except Exception:
            pass


def upload_to_sharepoint_folder(local_path: str, filename: str, folder_path: str,
                                sp_ctx: ClientContext) -> str:
    folder = sp_ctx.web.get_folder_by_server_relative_url(_sp_url(folder_path))
    with open(local_path, "rb") as fh:
        content = fh.read()
    uploaded = folder.upload_file(filename, content).execute_query()
    li = uploaded.listItemAllFields
    li.set_property("FileType", "Report")
    li.update()
    sp_ctx.execute_query()
    return uploaded.serverRelativeUrl


# =========================================================
# SMAPPLY HELPERS
# =========================================================

def _smapply_headers() -> Dict[str, str]:
    access_token = wmill.get_variable("u/epham2/SMAPPLY_ACCESS_TOKEN")
    return {"Authorization": f"Bearer {access_token}"}


def get_task_files(appid: int, task_id: int) -> List[Dict[str, Any]]:
    """Return list of files uploaded to a task by fetching the task endpoint directly.
    Files are embedded in task_data['data'] as lists of dicts with filename + url keys.
    Returns empty list if task has no files or request fails.
    """
    url  = f"https://nnas.smapply.io/api/applications/{appid}/tasks/{task_id}/"
    resp = requests.get(url, headers=_smapply_headers())
    if resp.status_code != 200:
        return []
    task_data = resp.json()
    files = []
    for value in task_data.get("data", {}).values():
        if isinstance(value, list):
            for entry in value:
                if isinstance(entry, dict) and "filename" in entry and "url" in entry:
                    files.append({
                        "id":       entry["id"],
                        "filename": entry["filename"],
                        "url":      entry["url"],
                        "task_id":  task_id,
                    })
    return files


def download_task_file(file_url: str, local_path: str):
    """Download a file from SMApply using its direct URL."""
    resp = requests.get(file_url, headers=_smapply_headers(), stream=True)
    resp.raise_for_status()
    os.makedirs(os.path.dirname(local_path) or ".", exist_ok=True)
    with open(local_path, "wb") as fh:
        for chunk in resp.iter_content(chunk_size=8192):
            fh.write(chunk)


# =========================================================
# ENTRYPOINT
# =========================================================

def main(
    appid: Any,
    reference_id: str,
):
    """
    Webhook fired when reviewer moves application to next stage.
    Checks each RB optional task for an uploaded file.
    If found: demotes the current _latest in SharePoint, uploads new file as _latest.
    If not found: skips.
    """
    app_id = int(appid) if not isinstance(appid, int) else appid
    ts     = datetime.utcnow().strftime("%Y%m%d_%H%M%S")

    os.makedirs("tmp", exist_ok=True)

    results = []

    for task_id, (group_code, shortname) in TASK_ROUTE_MAP.items():

        # --- Check if reviewer uploaded a file to this task ---
        try:
            files = get_task_files(app_id, task_id)
        except Exception as e:
            results.append({
                "task_id":    task_id,
                "group":      group_code,
                "shortname":  shortname,
                "status":     "error_fetching_files",
                "detail":     str(e),
            })
            continue

        if not files:
            results.append({
                "task_id":   task_id,
                "group":     group_code,
                "shortname": shortname,
                "status":    "skipped",
            })
            continue

        # Only one file expected per task (task configured for single upload)
        uploaded_file = files[0]
        file_url      = uploaded_file.get("url", "")
        if not file_url:
            results.append({
                "task_id":   task_id,
                "group":     group_code,
                "shortname": shortname,
                "status":    "error_no_file_url",
                "detail":    uploaded_file,
            })
            continue

        # --- Download the uploaded file locally ---
        tmp_path = os.path.join("tmp", f"{reference_id}_{shortname}_{group_code}_{ts}_edited.pdf")
        try:
            download_task_file(file_url, tmp_path)
        except Exception as e:
            results.append({
                "task_id":   task_id,
                "group":     group_code,
                "shortname": shortname,
                "status":    "error_downloading",
                "detail":    str(e),
            })
            continue

        # --- SharePoint: demote existing _latest, upload new one ---
        # Keep existing folder structure: <base>/<ref_id>/Credential reports/<shortname>/
        folder_path  = f"{APPLICATIONS_BASE}/{reference_id}/{CRED_REPORTS_SUBFOLDER}/{shortname}"
        latest_name  = f"{reference_id}_{shortname}_{ts}_latest.pdf"

        try:
            sp_ctx = ctx()
            _ensure_sp_folder_tree(sp_ctx, folder_path)
            _demote_latest_in_folder(sp_ctx, folder_path)
            sp_url = upload_to_sharepoint_folder(tmp_path, latest_name, folder_path, sp_ctx)
            results.append({
                "task_id":        task_id,
                "group":          group_code,
                "shortname":      shortname,
                "status":         "uploaded",
                "sharepoint_url": sp_url,
                "filename":       latest_name,
            })
        except Exception as e:
            results.append({
                "task_id":   task_id,
                "group":     group_code,
                "shortname": shortname,
                "status":    "error_sharepoint",
                "detail":    str(e),
            })

    uploaded_count = sum(1 for r in results if r["status"] == "uploaded")
    skipped_count  = sum(1 for r in results if r["status"] == "skipped")
    error_count    = sum(1 for r in results if r["status"].startswith("error"))

    return {
        "status":         "ok" if error_count == 0 else "partial",
        "appid":          app_id,
        "reference_id":   reference_id,
        "uploaded_count": uploaded_count,
        "skipped_count":  skipped_count,
        "error_count":    error_count,
        "results":        results,
    }
    