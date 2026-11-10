"""
Upload a file from SharePoint to an SM Apply upload task.

Defaults upload the IEN Forms placeholder letter to the Applicant Credential
Report Copies task (2077121) on application 35367543 in program 295185.

Run as a Windmill script: main(file_path, application_id, task_id, mark_complete).
"""

import mimetypes
import time
from urllib.parse import quote

import requests
import wmill

# =========================================================
# SMAPPLY CONFIG
# =========================================================

SMAPPLY_BASE = "https://nnas.smapply.io/api"
SMAPPLY_TOKEN_VAR = "u/epham2/SMAPPLY_ACCESS_TOKEN"

PROGRAM_ID = 295185
DEFAULT_APPLICATION_ID = 35367543
DEFAULT_TASK_ID = 2077121  # Applicant Credential Report Copies (upload task)
UPLOAD_TASK_TYPE = 2
COMPLETED_STATUS = 3

# =========================================================
# SHAREPOINT / MICROSOFT GRAPH CONFIG
# =========================================================

SHAREPOINT_HOSTNAME = "nnasca.sharepoint.com"
SHAREPOINT_SITE_PATH = "/sites/NNAS-Portal-Sharepoint"

# Document library holding the file.
SHAREPOINT_LIBRARY = "Applications"

# Path to the file inside the Applications library. A leading
# "Applications/" is also accepted and stripped automatically.
DEFAULT_FILE_PATH = "IEN-Forms/Placeholder_Letter_Template.docx"

GRAPH_TENANT_ID_VAR = "u/epham2/AZURE_TENANT_ID"
GRAPH_CLIENT_ID_VAR = "u/epham2/SHAREPOINT_CLIENT_ID"
GRAPH_CLIENT_SECRET_VAR = "u/epham2/SHAREPOINT_CLIENT_SECRET"

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# =========================================================
# SHAREPOINT / MICROSOFT GRAPH HELPERS
# =========================================================

def graph_token():
    tenant = wmill.get_variable(GRAPH_TENANT_ID_VAR).strip()
    resp = requests.post(
        f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token",
        data={
            "grant_type": "client_credentials",
            "client_id": wmill.get_variable(GRAPH_CLIENT_ID_VAR).strip(),
            "client_secret": wmill.get_variable(GRAPH_CLIENT_SECRET_VAR).strip(),
            "scope": "https://graph.microsoft.com/.default",
        },
        timeout=60,
    )
    if not resp.ok:
        raise RuntimeError(f"Graph token request failed: {resp.status_code}")
    return resp.json()["access_token"]


def library_relative_path(path):
    """Strip any leading site or library prefix so the path is relative to the library."""
    p = (path or "").strip().strip("/")
    site_prefix = SHAREPOINT_SITE_PATH.strip("/") + "/"
    if p.lower().startswith(site_prefix.lower()):
        p = p[len(site_prefix):]
    library_prefix = SHAREPOINT_LIBRARY + "/"
    if p.lower().startswith(library_prefix.lower()):
        p = p[len(library_prefix):]
    return p.strip("/")


def download_from_sharepoint(file_path):
    """Return (filename, content) for a file in the configured SharePoint library."""
    headers = {"Authorization": f"Bearer {graph_token()}"}

    site = requests.get(f"{GRAPH_BASE}/sites/{SHAREPOINT_HOSTNAME}:{SHAREPOINT_SITE_PATH}",
                        headers=headers, timeout=60)
    if not site.ok:
        raise RuntimeError(f"SharePoint site lookup failed: {site.status_code}")
    site_id = site.json()["id"]

    drives = requests.get(f"{GRAPH_BASE}/sites/{site_id}/drives", headers=headers, timeout=60)
    if not drives.ok:
        raise RuntimeError(f"SharePoint library lookup failed: {drives.status_code}")
    drive = next((d for d in drives.json()["value"] if d["name"] == SHAREPOINT_LIBRARY), None)
    if drive is None:
        names = ", ".join(d["name"] for d in drives.json()["value"])
        raise RuntimeError(f"Library '{SHAREPOINT_LIBRARY}' not found. Available: {names}")

    path = library_relative_path(file_path)
    resp = requests.get(
        f"{GRAPH_BASE}/drives/{drive['id']}/root:/{quote(path, safe='/')}:/content",
        headers=headers, timeout=120, allow_redirects=True,
    )
    if not resp.ok:
        raise RuntimeError(f"SharePoint download failed for '{path}': {resp.status_code}")
    return path.split("/")[-1], resp.content

# =========================================================
# SMAPPLY API HELPERS
# =========================================================

def smapply_session():
    session = requests.Session()
    session.headers.update({"Authorization": f"Bearer {wmill.get_variable(SMAPPLY_TOKEN_VAR)}"})
    return session


def smapply_request(session, method, path, **kwargs):
    url = path if path.startswith("http") else SMAPPLY_BASE + path
    for attempt in range(5):
        resp = session.request(method, url, timeout=120, **kwargs)
        if resp.status_code == 429 or resp.status_code >= 500:
            time.sleep(int(resp.headers.get("Retry-After", 2 ** attempt)))
            continue
        if not resp.ok:
            raise RuntimeError(f"SM Apply {method} {path} failed: {resp.status_code}")
        return resp.json() if resp.content else None
    raise RuntimeError(f"SM Apply {method} {path} failed after retries")


def find_task(session, application_id, task_id):
    url = f"/applications/{application_id}/tasks/"
    while url:
        page = smapply_request(session, "GET", url)
        for task in page["results"]:
            if task["id"] == task_id:
                return task
        url = page.get("next")
    return None

# =========================================================
# MAIN
# =========================================================

def main(
    file_path: str = DEFAULT_FILE_PATH,
    application_id: int = DEFAULT_APPLICATION_ID,
    task_id: int = DEFAULT_TASK_ID,
    mark_complete: bool = False,
):
    session = smapply_session()

    # Confirm the application and task before uploading anything.
    app = smapply_request(session, "GET", f"/applications/{application_id}/")
    program_id = (app.get("program") or {}).get("id")
    if program_id != PROGRAM_ID:
        raise RuntimeError(f"Application {application_id} is in program {program_id}, "
                           f"not {PROGRAM_ID}. Nothing was uploaded.")

    task = find_task(session, application_id, task_id)
    if task is None:
        raise RuntimeError(f"Task {task_id} not found on application {application_id}.")
    if task["type"] != UPLOAD_TASK_TYPE:
        raise RuntimeError(f"Task {task_id} is type {task['type']}, not an upload task.")

    filename, content = download_from_sharepoint(file_path)
    print(f"Downloaded '{filename}' from SharePoint ({len(content)} bytes)")

    # Upload tasks take the file under the "file" field.
    content_type = mimetypes.guess_type(filename)[0] or "application/octet-stream"
    smapply_request(session, "POST", f"/applications/{application_id}/tasks/{task_id}/files/",
                    files={"file": (filename, content, content_type)})
    print(f"Uploaded '{filename}' to task {task_id} on application {application_id}")

    if mark_complete:
        # No "apply-automations" header, so task automations do not fire.
        smapply_request(session, "PATCH", f"/applications/{application_id}/tasks/{task_id}/",
                        json={"status": COMPLETED_STATUS})
        print(f"Marked task {task_id} as completed")

    return {
        "application_id": application_id,
        "task_id": task_id,
        "task_name": task["name"],
        "uploaded_file": filename,
        "size_bytes": len(content),
        "marked_complete": mark_complete,
    }