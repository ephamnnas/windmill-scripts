# requirements:
# office365-rest-python-client
# paramiko
# wmill

import os
import posixpath
import tempfile
import shutil
from typing import Optional

import paramiko
import wmill
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext


# =========================================================
# CONFIG — SHAREPOINT
# =========================================================

SITE_URL      = "https://nnasca.sharepoint.com/sites/NNAS-Portal-Sharepoint"
CLIENT_ID     = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_ID")
CLIENT_SECRET = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_SECRET")
SP_BASE_FOLDER = "Applications/TM"


# =========================================================
# CONFIG — SFTP
# =========================================================

SFTP_HOST     = "rb-ftp-connect.nnas.ca"
SFTP_PORT     = 2022
SFTP_USER_VAR = wmill.get_variable("u/epham2/NNAS_SFTP_USERNAME")
SFTP_PW_VAR   = wmill.get_variable("u/epham2/NNAS_SFTP_PASSWORD")
SFTP_BASE_DIR = "/TO_TRUMERIT"


# =========================================================
# CONFIG — FILE TYPE MAPPING
# =========================================================

FILETYPE_TO_SUBFOLDER = {
    "Education":    "images/education",
    "Identity":     "images/identity",
    "Registration": "images/registration",
    "Other":        "images/other",
    "Report":       "images/report",
}

# Files to skip when uploading (trigger files, temp files etc.)
SKIP_FILENAMES = {".DS_Store", "Thumbs.db"}


# =========================================================
# HELPERS — LOGGING
# =========================================================

def log(msg: str, steps: list) -> None:
    print(msg)
    steps.append(msg)


# =========================================================
# HELPERS — SHAREPOINT
# =========================================================

def sp_ctx() -> ClientContext:
    creds = ClientCredential(CLIENT_ID, CLIENT_SECRET)
    return ClientContext(SITE_URL).with_credentials(creds)


def get_folder_files(application_id: str, steps: list) -> list[dict]:
    """
    Fetch all files in Applications/TM/{application_id}/ from SharePoint,
    including their FileType metadata tag.
    Returns a list of dicts with name, filetype, server_relative_url.
    """
    log(f"[STEP] Fetching files from SharePoint for {application_id}", steps)
    ctx         = sp_ctx()
    folder_path = f"/sites/NNAS-Portal-Sharepoint/{SP_BASE_FOLDER}/{application_id}"

    folder = ctx.web.get_folder_by_server_relative_url(folder_path)
    files  = folder.files
    ctx.load(files)
    ctx.execute_query()

    results = []
    for f in files:
        # Load file properties including metadata
        ctx.load(f, ["Name", "ServerRelativeUrl", "ListItemAllFields"])
        ctx.execute_query()

        try:
            item     = f.list_item_all_fields
            filetype = item.properties.get("FileType", "") or ""
        except Exception:
            filetype = ""

        results.append({
            "name":                f.name,
            "server_relative_url": f.server_relative_url,
            "filetype":            filetype.strip(),
        })
        log(f"[OK] Found file: {f.name} (FileType: '{filetype.strip()}')", steps)

    log(f"[OK] Total files found: {len(results)}", steps)
    return results


def download_file_from_sharepoint(server_relative_url: str, local_path: str, steps: list) -> None:
    """Download a single file from SharePoint to local_path."""
    ctx     = sp_ctx()
    sp_file = ctx.web.get_file_by_server_relative_url(server_relative_url)
    with open(local_path, "wb") as f:
        sp_file.download(f).execute_query()
    log(f"[OK] Downloaded: {os.path.basename(local_path)}", steps)


# =========================================================
# HELPERS — SFTP
# =========================================================

def ensure_remote_dir(sftp: paramiko.SFTPClient, remote_dir: str, steps: list) -> None:
    """Create remote_dir and all parents if they do not exist."""
    remote_dir = "/" + remote_dir.strip("/")
    path = ""
    for part in remote_dir.strip("/").split("/"):
        path = f"{path}/{part}"
        try:
            sftp.chdir(path)
        except IOError:
            log(f"[STEP] Creating remote dir: {path}", steps)
            sftp.mkdir(path)
            sftp.chdir(path)


def upload_files_to_sftp(
    file_map: list[dict],
    application_id: str,
    sftp_user: str,
    sftp_pw: str,
    steps: list,
) -> list[str]:
    """
    Upload all files in file_map to SFTPGo under /NNAS/{application_id}/{subfolder}/.
    file_map entries: { name, local_path, remote_subfolder }
    Returns list of uploaded remote paths.
    """
    uploaded: list[str] = []

    log(f"[STEP] Connecting to SFTP {SFTP_HOST}:{SFTP_PORT}", steps)
    transport = paramiko.Transport((SFTP_HOST, SFTP_PORT))
    try:
        transport.connect(username=sftp_user, password=sftp_pw)
        sftp = paramiko.SFTPClient.from_transport(transport)
        log("[OK] SFTP connected", steps)

        for entry in file_map:
            remote_dir  = posixpath.join(SFTP_BASE_DIR, application_id, entry["remote_subfolder"])
            remote_path = posixpath.join(remote_dir, entry["name"])
            ensure_remote_dir(sftp, remote_dir, steps)
            log(f"[STEP] Uploading {entry['name']} → {remote_path}", steps)
            sftp.put(entry["local_path"], remote_path)
            uploaded.append(remote_path)
            log(f"[OK] Uploaded: {remote_path}", steps)

        return uploaded

    finally:
        try:
            transport.close()
            log("[OK] SFTP connection closed", steps)
        except Exception:
            pass


# =========================================================
# ENTRYPOINT
# =========================================================

def main(
    application_id: str,
    trigger: Optional[dict] = None,
):
    """
    Windmill entrypoint — triggered by Power Automate when a specific
    trigger file is created in SharePoint for a given application.

    Flow:
      1. Resolve application_id from input or trigger payload
      2. Fetch all files from SharePoint Applications/TM/{application_id}/
      3. Download each file to a temp directory
      4. Map each file to the correct SFTPGo subfolder using FileType tag
      5. Upload all files to /NNAS/{application_id}/{subfolder}/
      6. Clean up temp directory

    Returns a structured result dict with status, uploaded paths, and debug steps.
    """
    steps: list[str] = []
    dbg:   dict      = {}

    # 1. Resolve application_id
    if trigger and isinstance(trigger, dict):
        log(f"[DEBUG] Trigger payload: {trigger}", steps)
        application_id = application_id or trigger.get("application_id", "")

    if not application_id:
        msg = "application_id is required but missing."
        log(f"[ERROR] {msg}", steps)
        return {"status": "error", "message": msg, "steps": steps}

    log(f"[START] Processing application: {application_id}", steps)

    sftp_user = wmill.get_variable(SFTP_USER_VAR)
    sftp_pw   = wmill.get_variable(SFTP_PW_VAR)

    if not sftp_user or not sftp_pw:
        return {"status": "error", "message": "Missing SFTP credential variables.", "steps": steps}

    temp_dir = tempfile.mkdtemp(prefix="sp_to_sftp_")

    try:
        # 2. Fetch file list from SharePoint
        sp_files = get_folder_files(application_id, steps)

        if not sp_files:
            msg = f"No files found in SharePoint for {application_id}"
            log(f"[WARN] {msg}", steps)
            return {"status": "warning", "message": msg, "steps": steps}

        # 3 & 4. Download files and build file map
        file_map: list[dict] = []
        skipped:  list[str]  = []

        for sp_file in sp_files:
            name = sp_file["name"]

            # Skip system files
            if name in SKIP_FILENAMES:
                skipped.append(name)
                log(f"[OK] Skipping system file: {name}", steps)
                continue

            # Determine subfolder from FileType tag
            filetype   = sp_file["filetype"]
            subfolder  = FILETYPE_TO_SUBFOLDER.get(filetype)

            if not subfolder:
                # FileType missing or unrecognised — route to Other
                log(f"[WARN] FileType '{filetype}' not recognised for {name}; routing to images/other", steps)
                subfolder = "images/other"

            # Download to temp
            local_path = os.path.join(temp_dir, name)
            download_file_from_sharepoint(sp_file["server_relative_url"], local_path, steps)

            file_map.append({
                "name":             name,
                "local_path":       local_path,
                "remote_subfolder": subfolder,
                "filetype":         filetype,
            })

        dbg["total_files"]    = len(sp_files)
        dbg["files_to_upload"] = len(file_map)
        dbg["skipped_files"]  = skipped
        dbg["file_map"] = [
            {"name": e["name"], "filetype": e["filetype"], "subfolder": e["remote_subfolder"]}
            for e in file_map
        ]

        # 5. Upload all files to SFTPGo
        uploaded_files = upload_files_to_sftp(
            file_map, application_id, sftp_user, sftp_pw, steps
        )

        log("[DONE] Workflow completed successfully", steps)

        return {
            "status":          "ok",
            "application_id":  application_id,
            "files_uploaded":  len(uploaded_files),
            "uploaded_paths":  uploaded_files,
            "skipped_files":   skipped,
            "debug":           dbg,
            "steps":           steps,
        }

    except Exception as exc:
        msg = f"Unexpected error: {exc}"
        log(f"[ERROR] {msg}", steps)
        return {"status": "error", "message": msg, "steps": steps}

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
        log(f"[OK] Temp directory cleaned up", steps)
