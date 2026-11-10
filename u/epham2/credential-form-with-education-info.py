# py: 3.12
# requirements:
# pypdf
# pdfplumber
# reportlab
# requests
# wmill
# python-docx
# pillow==10.4.0 --no-binary=:all:

import os
import json
import re
import time
import requests
from io import BytesIO
from typing import Dict, List, Any, Optional, Tuple, Set

from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading

import pdfplumber
import wmill
from pypdf import PdfReader, PdfWriter

from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.utils import simpleSplit
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Table, TableStyle, Paragraph

from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches, Emu
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL

# =========================================================
# STYLE CONSTANTS - PDF
# =========================================================

FONT_BODY  = 10
FONT_LABEL = 10

COLOR_RED       = colors.HexColor("#D12129")
COLOR_TEXT      = colors.HexColor("#211D1E")
COLOR_HEADER_BG = colors.HexColor("#E6E6E6")
COLOR_ALT_ROW   = colors.HexColor("#F2F2F2")

OUTER_BORDER_WIDTH = 1.5
INNER_BORDER_WIDTH = 0.5
CELL_PADDING       = 6

DOCX_RED    = RGBColor(0xD1, 0x21, 0x29)
DOCX_TEXT   = RGBColor(0x21, 0x1D, 0x1E)
DOCX_GREY   = RGBColor(0xE6, 0xE6, 0xE6)
DOCX_BLACK  = RGBColor(0x00, 0x00, 0x00)
DOCX_WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
DOCX_MAROON = RGBColor(0x6B, 0x0E, 0x13)

# =========================================================
# POSITIONING CONSTANTS - PDF
# =========================================================

ONE_INCH  = 72
HALF_INCH = 36

LEFT_MARGIN = 30
TABLE_WIDTH = 530

PERSONAL_INFO_HEADER_Y = 640
REGISTRATION_HEADER_Y  = 700
PART_C_Y_TOP           = 620 - ONE_INCH

NOTES_SECTION_X        = 30
NOTES_SECTION_W        = 550
GAP_BETWEEN_NOTE_BOXES = 20

EVALUATOR_LABEL_X1 = 73.7
DATE_LABEL_X1      = 431.1
EVALUATOR_VALUE_X  = EVALUATOR_LABEL_X1 + 6
DATE_VALUE_X       = DATE_LABEL_X1 + 6
EVALUATOR_Y        = 792 - 442.7

# =========================================================
# SHAREPOINT CONFIG
# =========================================================

SITE_URL      = "https://nnasca.sharepoint.com/sites/NNAS-Portal-Sharepoint"
TENANT_ID     = wmill.get_variable("u/epham2/AZURE_TENANT_ID")
CLIENT_ID     = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_ID")
CLIENT_SECRET = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_SECRET")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SHAREPOINT_HOSTNAME = "nnasca.sharepoint.com"
SITE_PATH = "/sites/NNAS-Portal-Sharepoint"
DOCUMENT_LIBRARY_NAME = "Applications"

TEMPLATE_FOLDER        = "Applications/Credential form"
TEMPLATE_FILENAME      = "CredentialReportBaseV1.pdf"       # no education sections
TEMPLATE_FILENAME_FULL = "CredentialReportBaseV1Full.pdf"   # includes A2 + B1 pages
LOGO_FILENAME          = "Credential_header.png"

CRED_REPORTS_SUBFOLDER = "Credential reports"

# Sibling of CRED_REPORTS_SUBFOLDER at the application's root
# (applications/<provider>/<reference_id>/). Reviewers may drop
# supplementary files into subfolders here; their bare subfolder NAMES
# (not the files inside, not dates) get listed as extra bullets under
# Part C's "C4: Other Documents".
ADDITIONAL_DOCUMENTS_FOLDER_NAME = "Additional documents"

SP_MAX_WORKERS = 8

# When create_folder loses a race to another worker thread (Graph returns
# 409 nameAlreadyExists), it re-fetches the folder the other thread made.
# Graph can briefly return 404 on a GET-by-path right after the name has
# been registered, so the re-fetch retries a few times with a short,
# linearly increasing delay (0.5s, 1.0s, 1.5s, 2.0s).
_FOLDER_CONFLICT_GET_RETRIES       = 4
_FOLDER_CONFLICT_GET_DELAY_SECONDS = 0.5

PAGE_NOTES = 0   # Page 1  — cover / notes / metadata
PAGE_A1    = 1   # Page 2  — A1 Personal Information
PAGE_A2    = 2   # Page 3  — A2 Nursing Education Overview (full template only)
PAGE_A3    = 3   # Page 4  — A3 Registration/License (first entry)

TEMPLATE_FULL_BASE_PAGES = 7

# =========================================================
# PROVIDER CONFIG
# =========================================================
#
# Provider ("TM" or "CES") is read from additional_content.application_info
# .provider. An unknown/missing provider is a HARD ERROR (no default):
# guessing wrong would hit the wrong SMApply program's reviews endpoint,
# complete the wrong SMApply task, and file into the wrong SharePoint tree.
# A manual `provider` input to main() overrides the payload value for ad
# hoc runs. Mirrors the pattern used in the ID-documents intake script.

PROVIDER_CONFIG = {
    "tm": {
        "program_id":        295185,
        "applications_base": "applications/tm",
        "smapply_task_id":   2051542,
    },
    "ces": {
        "program_id":        295865,
        "applications_base": "applications/ces",
        "smapply_task_id":   2053932,
    },
}


def resolve_provider(provider_override: str, payload: Dict[str, Any]) -> str:
    """
    Normalized provider key ('tm' / 'ces'). Precedence: explicit override,
    then additional_content.application_info.provider. Unknown or missing
    provider raises: no default, since a wrong guess would hit the wrong
    program's reviews endpoint, complete the wrong SMApply task, and file
    into the wrong SharePoint tree.
    """
    candidate = (provider_override or "").strip().lower()
    if not candidate:
        candidate = str(
            (payload.get("application_info") or {}).get("provider") or ""
        ).strip().lower()

    if candidate in PROVIDER_CONFIG:
        return candidate

    raise RuntimeError(
        "Cannot resolve provider. "
        f"Got '{candidate or '(empty)'}'; expected one of: "
        f"{', '.join(sorted(PROVIDER_CONFIG))}. "
        "Provider comes from additional_content.application_info.provider, "
        "or pass a PROVIDER input explicitly for manual runs."
    )

# =========================================================
# SMAPPLY CONFIG
# =========================================================
#
# SMAPPLY_TASK_ID is no longer a single global: it now comes from
# PROVIDER_CONFIG[provider]["smapply_task_id"], resolved in main().

# =========================================================
# PRIOR REPORT CONFIG
# =========================================================

PRIOR_REPORT_SUFFIX = "_priorreport.pdf"

# =========================================================
# SHAREPOINT HELPERS - MICROSOFT GRAPH
# =========================================================

_GRAPH_TOKEN = None
_SP_SITE_ID = None
_SP_DRIVE_ID = None
_SP_LIST_ID = None
_SP_INIT_LOCK = threading.Lock()
_SP_TOKEN_LOCK = threading.Lock()

def ctx():
    """
    Compatibility shim for older call sites.

    The old version returned an office365 ClientContext. The Graph version does
    not need a context object, but main() still passes ctx() into helper
    functions. Returning None lets those call sites remain unchanged.
    """
    return None

def _sp_url(folder_path: str) -> str:
    """Return the legacy server-relative URL used in debug/return payloads."""
    return f"/sites/NNAS-Portal-Sharepoint/{folder_path.strip('/')}"

def _drive_relative_path(path: str) -> str:
    """
    Convert old SharePoint paths to Microsoft Graph drive-root-relative paths.

    Examples:
      Applications/Credential form/file.pdf -> Credential form/file.pdf
      applications/tm/REF123              -> tm/REF123
      /sites/NNAS-Portal-Sharepoint/Applications/tm/REF123 -> tm/REF123
    """
    if not path:
        return ""

    p = path.strip()

    site_prefix = "/sites/NNAS-Portal-Sharepoint/"
    if p.lower().startswith(site_prefix.lower()):
        p = p[len(site_prefix):]

    p = p.strip("/")

    parts = [part for part in p.split("/") if part]
    if parts and parts[0].lower() == DOCUMENT_LIBRARY_NAME.lower():
        parts = parts[1:]

    return "/".join(parts)

def _encode_graph_path(path: str) -> str:
    """Encode a Graph drive path while preserving folder slashes."""
    from urllib.parse import quote
    return quote(path.strip("/"), safe="/")

def get_graph_token() -> str:
    """Get a Microsoft Graph app-only access token."""
    global _GRAPH_TOKEN

    if _GRAPH_TOKEN:
        return _GRAPH_TOKEN

    with _SP_TOKEN_LOCK:
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
    """Call Microsoft Graph and raise a useful error if the request fails."""
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

def init_sharepoint_graph() -> None:
    """Find the SharePoint site, Applications document library drive, and backing list."""
    global _SP_SITE_ID, _SP_DRIVE_ID, _SP_LIST_ID

    if _SP_SITE_ID and _SP_DRIVE_ID and _SP_LIST_ID:
        return

    with _SP_INIT_LOCK:
        if _SP_SITE_ID and _SP_DRIVE_ID and _SP_LIST_ID:
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

        lists = graph_request(
            "GET",
            f"/sites/{_SP_SITE_ID}/lists?$select=id,displayName",
            expected_statuses=(200,),
        )["value"]

        sp_list = next(
            (l for l in lists if l.get("displayName") == DOCUMENT_LIBRARY_NAME),
            None,
        )

        if not sp_list:
            available = [l.get("displayName") for l in lists]
            raise RuntimeError(
                f"Could not find SharePoint list '{DOCUMENT_LIBRARY_NAME}'. "
                f"Available lists: {available}"
            )

        _SP_LIST_ID = sp_list["id"]

        print("Connected to SharePoint via Microsoft Graph")
        print("Site ID:", _SP_SITE_ID)
        print("Drive ID:", _SP_DRIVE_ID)
        print("List ID:", _SP_LIST_ID)

def get_drive_item_by_path(path: str):
    """Return a drive item by old SharePoint or drive-relative path, or None if missing."""
    init_sharepoint_graph()

    drive_path = _drive_relative_path(path)

    if not drive_path:
        return graph_request(
            "GET",
            f"/drives/{_SP_DRIVE_ID}/root",
            expected_statuses=(200,),
        )

    return graph_get_or_none(
        f"/drives/{_SP_DRIVE_ID}/root:/{_encode_graph_path(drive_path)}"
    )

def list_drive_children(path: str) -> List[Dict[str, Any]]:
    """List files/folders under a SharePoint folder. Returns [] if the folder is missing."""
    init_sharepoint_graph()

    drive_path = _drive_relative_path(path)

    if drive_path:
        endpoint = f"/drives/{_SP_DRIVE_ID}/root:/{_encode_graph_path(drive_path)}:/children"
    else:
        endpoint = f"/drives/{_SP_DRIVE_ID}/root/children"

    response = requests.get(
        f"{GRAPH_BASE}{endpoint}",
        headers={"Authorization": f"Bearer {get_graph_token()}"},
        timeout=120,
    )

    if response.status_code == 404:
        return []

    if response.status_code != 200:
        raise RuntimeError(
            f"Failed listing SharePoint folder.\n"
            f"Path: {path}\n"
            f"Status: {response.status_code}\n"
            f"Response body:\n{response.text[:5000]}"
        )

    data = response.json()
    return data.get("value", [])

def create_folder(parent_path: str, folder_name: str):
    """
    Create a folder under parent_path if it does not already exist, and
    return its drive item.

    Safe to call concurrently. _process_rb runs in a thread pool, and
    several regulatory bodies for the same application share parent
    folders (applications/<provider>/<reference_id>/Credential reports).
    Two threads can both see "folder missing" in the existence check
    below and both POST to create it; the slower one gets a
    409 nameAlreadyExists from Graph. Since "the folder exists" is the
    outcome we wanted, that specific conflict is treated as success: the
    existing folder is fetched and returned instead of raising.

    conflictBehavior stays "fail" on purpose. "replace" could swap out a
    folder that other threads have already uploaded into, and "rename"
    would create duplicate folders like "<name> 1".
    """
    init_sharepoint_graph()

    parent_drive_path = _drive_relative_path(parent_path)
    full_drive_path = f"{parent_drive_path}/{folder_name}".strip("/")

    existing = get_drive_item_by_path(full_drive_path)
    if existing:
        return existing

    if parent_drive_path:
        endpoint = f"/drives/{_SP_DRIVE_ID}/root:/{_encode_graph_path(parent_drive_path)}:/children"
    else:
        endpoint = f"/drives/{_SP_DRIVE_ID}/root/children"

    # Accept 409 here so graph_request returns the error body instead of
    # raising; the body is inspected below to decide what to do with it.
    result = graph_request(
        "POST",
        endpoint,
        headers={"Content-Type": "application/json"},
        json={
            "name": folder_name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "fail",
        },
        expected_statuses=(201, 409),
    )

    # 201 Created: the response body is the new folder's drive item.
    if not (isinstance(result, dict) and "error" in result):
        return result

    # 409: only "nameAlreadyExists" is the benign race. Any other conflict
    # code is a real problem and should still fail loudly.
    error_code = str((result.get("error") or {}).get("code") or "")
    if error_code != "nameAlreadyExists":
        raise RuntimeError(
            "Microsoft Graph folder creation returned an unexpected conflict.\n"
            f"Endpoint: {endpoint}\n"
            f"Folder: {full_drive_path}\n"
            f"Response body:\n{json.dumps(result)[:5000]}"
        )

    # Lost the race to another thread or run. Fetch the folder it created,
    # retrying briefly in case Graph has registered the name but a
    # GET-by-path still returns 404 for a moment.
    for attempt in range(_FOLDER_CONFLICT_GET_RETRIES):
        existing = get_drive_item_by_path(full_drive_path)
        if existing:
            if "folder" not in existing:
                raise RuntimeError(
                    f"Cannot create folder '{full_drive_path}': an item with "
                    "that name already exists but it is a file, not a folder."
                )
            return existing
        time.sleep(_FOLDER_CONFLICT_GET_DELAY_SECONDS * (attempt + 1))

    raise RuntimeError(
        f"Graph reported that folder '{full_drive_path}' already exists "
        "(409 nameAlreadyExists), but it could not be retrieved after "
        f"{_FOLDER_CONFLICT_GET_RETRIES} attempts."
    )

def _ensure_sp_folder_tree(sp_ctx: Any, folder_path: str):
    """Create each folder segment under the Applications document library if missing."""
    drive_path = _drive_relative_path(folder_path)
    parts = [p for p in drive_path.split("/") if p]

    current_path = ""
    for part in parts:
        create_folder(current_path, part)
        current_path = f"{current_path}/{part}".strip("/")

def download_drive_item_to_path(drive_item_id: str, local_path: str) -> None:
    """Download a drive item by ID to a local file."""
    token = get_graph_token()

    response = requests.get(
        f"{GRAPH_BASE}/drives/{_SP_DRIVE_ID}/items/{drive_item_id}/content",
        headers={"Authorization": f"Bearer {token}"},
        timeout=300,
        allow_redirects=True,
    )

    if response.status_code != 200:
        raise RuntimeError(
            f"Failed downloading SharePoint file.\n"
            f"Item ID: {drive_item_id}\n"
            f"Status: {response.status_code}\n"
            f"Response body:\n{response.text[:5000]}"
        )

    with open(local_path, "wb") as fh:
        fh.write(response.content)

def download_file_by_path(path: str, local_path: str) -> None:
    """Download a file by SharePoint/drive path to a local file."""
    init_sharepoint_graph()

    drive_path = _drive_relative_path(path)
    token = get_graph_token()

    response = requests.get(
        f"{GRAPH_BASE}/drives/{_SP_DRIVE_ID}/root:/{_encode_graph_path(drive_path)}:/content",
        headers={"Authorization": f"Bearer {token}"},
        timeout=300,
        allow_redirects=True,
    )

    if response.status_code != 200:
        raise RuntimeError(
            f"Failed downloading SharePoint file.\n"
            f"Path: {path}\n"
            f"Drive path: {drive_path}\n"
            f"Status: {response.status_code}\n"
            f"Response body:\n{response.text[:5000]}"
        )

    with open(local_path, "wb") as fh:
        fh.write(response.content)

def patch_driveitem_fields(drive_item_id: str, fields: Dict[str, Any]) -> None:
    """Update SharePoint metadata fields for a drive item through its backing list item."""
    init_sharepoint_graph()

    item = graph_request(
        "GET",
        f"/drives/{_SP_DRIVE_ID}/items/{drive_item_id}?$expand=listItem",
        expected_statuses=(200,),
    )

    list_item = item.get("listItem")
    if not list_item or not list_item.get("id"):
        raise RuntimeError(f"Drive item {drive_item_id} has no listItem information.")

    list_item_id = list_item["id"]

    graph_request(
        "PATCH",
        f"/sites/{_SP_SITE_ID}/lists/{_SP_LIST_ID}/items/{list_item_id}/fields",
        headers={"Content-Type": "application/json"},
        json=fields,
        expected_statuses=(200, 204),
    )

def upload_bytes_to_sharepoint(path: str, content: bytes, file_type: str = "Report") -> Dict[str, Any]:
    """Upload bytes to SharePoint and set FileType metadata."""
    init_sharepoint_graph()

    drive_path = _drive_relative_path(path)

    uploaded_file = graph_request(
        "PUT",
        f"/drives/{_SP_DRIVE_ID}/root:/{_encode_graph_path(drive_path)}:/content",
        headers={"Content-Type": "application/octet-stream"},
        data=content,
        expected_statuses=(200, 201),
    )

    patch_driveitem_fields(uploaded_file["id"], {"FileType": file_type})
    return uploaded_file

def delete_drive_item(drive_item_id: str) -> None:
    graph_request(
        "DELETE",
        f"/drives/{_SP_DRIVE_ID}/items/{drive_item_id}",
        expected_statuses=(204,),
    )

def download_template(filename: str = TEMPLATE_FILENAME) -> str:
    """Download a base report template from SharePoint to tmp/.
    Pass TEMPLATE_FILENAME_FULL when a prior report was found.
    """
    os.makedirs("tmp", exist_ok=True)
    path = os.path.join("tmp", filename)
    download_file_by_path(f"{TEMPLATE_FOLDER}/{filename}", path)

    if os.path.getsize(path) == 0:
        raise RuntimeError(f"Downloaded template is empty: {path}")

    return path

def download_logo() -> Optional[str]:
    try:
        os.makedirs("tmp", exist_ok=True)
        path = os.path.join("tmp", LOGO_FILENAME)
        download_file_by_path(f"{TEMPLATE_FOLDER}/{LOGO_FILENAME}", path)

        if os.path.getsize(path) == 0:
            return None

        return path
    except Exception:
        return None

def _demote_latest_in_folder(sp_ctx: Any, folder_path: str,
                             name_filter: Optional[str] = None):
    """
    Rename any *_latest files to non-latest names by copying to the demoted
    filename and deleting the original. Graph does not have a direct overwrite
    semantic that exactly matches the old helper, so this preserves behavior.
    """
    try:
        files = list_drive_children(folder_path)
    except Exception:
        return

    os.makedirs("tmp", exist_ok=True)

    for f in list(files):
        name = f.get("name", "")
        if not name or "." not in name:
            continue
        if "folder" in f:
            continue
        if name_filter is not None and name_filter not in name:
            continue

        stem, _, ext = name.rpartition(".")
        if not stem.endswith("_latest"):
            continue

        new_name = f"{stem[:-len('_latest')]}.{ext}"
        tmp_local = os.path.join("tmp", f"_demote_{new_name}")

        try:
            download_drive_item_to_path(f["id"], tmp_local)
        except Exception:
            continue

        try:
            with open(tmp_local, "rb") as fh:
                content = fh.read()

            upload_bytes_to_sharepoint(
                f"{folder_path}/{new_name}",
                content,
                file_type="Report",
            )
        except Exception:
            continue

        try:
            delete_drive_item(f["id"])
        except Exception:
            pass

def upload_to_sharepoint_folder(local_path: str, filename: str, folder_path: str,
                                file_type: str = "Report",
                                sp_ctx: Any = None,
                                ensure_folder: bool = True) -> str:
    if ensure_folder:
        _ensure_sp_folder_tree(sp_ctx, folder_path)

    with open(local_path, "rb") as fh:
        content = fh.read()

    uploaded_file = upload_bytes_to_sharepoint(
        f"{folder_path}/{filename}",
        content,
        file_type=file_type,
    )

    # Keep the old return shape: server-relative URL.
    return _sp_url(f"{folder_path}/{filename}")

# =========================================================
# PRIOR REPORT DISCOVERY
# =========================================================

def find_prior_report_sharepoint(sp_ctx: Any,
                                  folder_path: str,
                                  local_tmp_dir: str = "tmp") -> Optional[str]:
    """
    Search a SharePoint folder for a file ending with '_priorreport.pdf'
    and download it locally. Returns the local path, or None if not found.

    Silently returns None if the folder does not exist (first-time jobs).
    The suffix match is case-insensitive.
    """
    try:
        files = list_drive_children(folder_path)
    except Exception:
        return None

    for f in files:
        name = f.get("name", "")
        if not name or "folder" in f:
            continue

        if name.lower().endswith(PRIOR_REPORT_SUFFIX):
            os.makedirs(local_tmp_dir, exist_ok=True)
            local_path = os.path.join(local_tmp_dir, name)

            try:
                download_drive_item_to_path(f["id"], local_path)
                if os.path.getsize(local_path) > 0:
                    return local_path
            except Exception:
                continue

    return None

def find_prior_report_local(search_dir: str) -> Optional[str]:
    """
    Scan a local directory for a file ending with '_priorreport.pdf'.
    Useful for local development / testing without SharePoint.
    Returns the full path of the first match, or None.
    """
    if not os.path.isdir(search_dir):
        return None
    for fname in sorted(os.listdir(search_dir)):
        if fname.lower().endswith(PRIOR_REPORT_SUFFIX):
            return os.path.join(search_dir, fname)
    return None

# =========================================================
# "ADDITIONAL DOCUMENTS" FOLDER SCAN -> C4
# ---------------------------------------------------------
# C4 ("Other Documents") is built ENTIRELY from this folder scan -- not
# from any SMApply review task. ADDITIONAL_DOCUMENTS_FOLDER_NAME is a
# sibling of CRED_REPORTS_SUBFOLDER at the application's root
# (applications/<provider>/<reference_id>/). Reviewers organize ad hoc
# supplementary material into named subfolders there (e.g. "Marriage
# Certificate", "Court Order"), each of which may or may not actually
# contain files. Per spec: step through each direct subfolder, and for
# any subfolder that has files in it, list those files; subfolders with
# no files inside contribute nothing (not even their own name).
# =========================================================

def list_additional_documents_files_by_subfolder(application_folder_path: str) -> List[str]:
    """
    Scan <application_folder_path>/Additional documents, step through
    each direct subfolder, and for every subfolder that contains at
    least one file, list those file names (each prefixed with its
    subfolder name for context, e.g. "Marriage Certificate: scan.pdf").
    Subfolders with no files inside are skipped entirely -- only actual
    file content is surfaced, never a bare subfolder name.

    Returns [] if "Additional documents" doesn't exist for this
    application, has no subfolders, or none of its subfolders contain
    any files -- or on any listing failure, since this is supplementary
    content, not required for report generation.
    """
    root_path = f"{application_folder_path}/{ADDITIONAL_DOCUMENTS_FOLDER_NAME}"
    try:
        subfolders = list_drive_children(root_path)
    except Exception:
        return []

    bullets: List[str] = []
    for item in sorted(subfolders, key=lambda it: str(it.get("name", "")).lower()):
        if "folder" not in item:
            continue   # ignore loose files sitting directly under "Additional documents"
        subfolder_name = item.get("name", "").strip()
        if not subfolder_name:
            continue

        subfolder_path = f"{root_path}/{subfolder_name}"
        try:
            children = list_drive_children(subfolder_path)
        except Exception:
            continue   # unreadable subfolder -- skip rather than fail the whole report

        file_names = sorted(
            (
                child.get("name", "").strip()
                for child in children
                if "folder" not in child and child.get("name", "").strip()
            ),
            key=str.lower,
        )
        for fname in file_names:
            bullets.append(f"{subfolder_name}: {fname}")

    return bullets


def apply_c4_additional_documents_files(payload: Dict[str, Any],
                                        file_bullets: List[str]) -> Dict[str, Any]:
    """
    Set payload's C4 section to the given "Additional documents" file
    bullets (creating the section if none exists yet). Mutates and
    returns `payload`. No-op if `file_bullets` is empty -- C4 simply
    doesn't appear in that case, same as before this change.
    """
    if not file_bullets:
        return payload

    sections = payload.get("part_c_sections") or []
    c4_section = next(
        (s for s in sections if str(s.get("header", "")).startswith("C4")),
        None,
    )
    if c4_section is None:
        c4_section = {"header": "C4: Other Documents", "bullets": []}
        sections.append(c4_section)
    c4_section["bullets"] = list(file_bullets)

    payload["part_c_sections"] = sections
    return payload

# =========================================================
# SMAPPLY HELPERS
# =========================================================

def coerce_appid(appid: Any) -> int:
    if isinstance(appid, int):
        return appid
    if isinstance(appid, str) and appid.strip().isdigit():
        return int(appid.strip())
    if isinstance(appid, dict) and "id" in appid:
        return int(appid["id"])
    raise ValueError(f"Invalid appid. Got: {type(appid)}")

def _smapply_upload(*, appid, task_id, filepath, upload_filename, mime) -> Dict[str, Any]:
    access_token = wmill.get_variable("u/epham2/SMAPPLY_ACCESS_TOKEN")
    url = f"https://nnas.smapply.io/api/applications/{appid}/tasks/{task_id}/files/"
    with open(filepath, "rb") as f:
        file_data = f.read()
    resp = requests.post(
        url,
        files={"file": (upload_filename, file_data, mime)},
        headers={"Authorization": f"Bearer {access_token}"},
    )
    try:
        body = resp.json()
    except Exception:
        body = {"text": resp.text}
    if not resp.ok:
        raise RuntimeError(
            f"SMApply upload failed {resp.status_code} for task {task_id}, "
            f"file '{upload_filename}': {body}"
        )
    return {"status_code": resp.status_code, "response": body}

def upload_pdf_to_smapply_task(*, appid, task_id, pdf_filepath, upload_filename):
    return _smapply_upload(appid=appid, task_id=task_id, filepath=pdf_filepath,
                           upload_filename=upload_filename, mime="application/pdf")

def upload_docx_to_smapply_task(*, appid, task_id, docx_filepath, upload_filename):
    mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    return _smapply_upload(appid=appid, task_id=task_id, filepath=docx_filepath,
                           upload_filename=upload_filename, mime=mime)

SMAPPLY_STATUS_COMPLETED = 3

def mark_smapply_task_complete(*, appid, task_id) -> Dict[str, Any]:
    access_token = wmill.get_variable("u/epham2/SMAPPLY_ACCESS_TOKEN")
    url = f"https://nnas.smapply.io/api/applications/{appid}/tasks/{task_id}/"
    payload: Dict[str, Any] = {"status": SMAPPLY_STATUS_COMPLETED}
    resp = requests.patch(
        url,
        json=payload,
        headers={
            "Authorization":     f"Bearer {access_token}",
            "Content-Type":      "application/json",
            "apply-automations": "full",
        },
    )
    try:
        body = resp.json()
    except Exception:
        body = {"text": resp.text}
    return {
        "status_code":  resp.status_code,
        "ok":           resp.ok,
        "url":          url,
        "sent_payload": payload,
        "response":     body,
    }

SMAPPLY_BASE = "https://nnas.smapply.io/api"

def fetch_all_reviews(application_id: int, program_id: int) -> List[Dict[str, Any]]:
    """
    Fetch every review submission (Application Review, Education Document
    Review, Registration Document Review, Identity Document Review, etc.)
    for an application/program from SMApply's reviews endpoint, following
    pagination via the `next` link. Returns a flat list of review dicts.

    Called once per run in main() and reused both by
    build_payload_from_reviews (registration/identity/notes) and
    parse_manual_education_from_reviews (education fallback), so a single
    application never triggers more than one reviews fetch.
    """
    access_token = wmill.get_variable("u/epham2/SMAPPLY_ACCESS_TOKEN")
    headers = {"Authorization": f"Bearer {access_token}", "Accept": "application/json"}
    url = f"{SMAPPLY_BASE}/reviews/"
    params: Optional[dict] = {"application": application_id, "program": program_id}
    out: List[Dict[str, Any]] = []
    while url:
        resp = requests.get(url, headers=headers, params=params, timeout=120)
        if not resp.ok:
            raise RuntimeError(
                f"SMApply reviews fetch failed {resp.status_code} "
                f"application={application_id} program={program_id}: {resp.text[:2000]}"
            )
        data = resp.json()
        out.extend(data.get("results", []))
        url = data.get("next")   # absolute URL or None
        params = None            # `next` already carries the query
    return out

# ---------------------------------------------------------
# RESTORED FROM SCRIPT 1: clear existing files on the SMApply upload
# task before re-uploading. The .../tasks/{id}/files/ collection forbids
# GET, so files are read from the task detail. Called ONCE per delivery
# (before the per-RB upload loop in main()), because every RB uploads to
# the same task; clearing per-upload would delete files uploaded by
# earlier RBs in the same run.
# ---------------------------------------------------------

_SMAPPLY_FILE_SIGNATURE_KEYS = {
    "file", "filename", "original_filename", "download_url", "url",
    "content_type", "mimetype", "size", "filesize", "extension",
}


def _smapply_headers() -> Dict[str, str]:
    access_token = wmill.get_variable("u/epham2/SMAPPLY_ACCESS_TOKEN")
    return {"Authorization": f"Bearer {access_token}"}


def _extract_task_files(task_json: Any) -> List[Dict[str, Any]]:
    """Locate uploaded-file objects in a task detail payload. Tries the known
    'data.files' location first, then a top-level 'files', then a recursive
    search for file-like objects (an 'id' alongside a file/name/url/size-type
    key). Returns file dicts de-duplicated by id."""
    if not isinstance(task_json, dict):
        return []

    data = task_json.get("data")
    if isinstance(data, dict) and isinstance(data.get("files"), list):
        candidates = data["files"]
    elif isinstance(task_json.get("files"), list):
        candidates = task_json["files"]
    else:
        candidates = None

    if candidates:
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


def clear_smapply_task_files(appid, task_id) -> Dict[str, Any]:
    """Delete every file currently on the task before new uploads, so a re-run
    doesn't accumulate stale reports. Run ONCE before the upload loop."""
    headers = _smapply_headers()
    task_url  = f"https://nnas.smapply.io/api/applications/{appid}/tasks/{task_id}/"
    files_url = f"https://nnas.smapply.io/api/applications/{appid}/tasks/{task_id}/files/"

    task_resp = requests.get(task_url, headers=headers, timeout=120)
    if not task_resp.ok:
        raise RuntimeError(
            f"SMApply get task failed {task_resp.status_code} task {task_id}: "
            f"{task_resp.text[:2000]}"
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
                f"task {task_id}: {del_resp.text[:2000]}"
            )
        deleted.append(file_id)

    return {"task_id": task_id, "found": len(existing), "deleted": deleted}

# =========================================================
# SMAPPLY REVIEWS -> PAYLOAD ADAPTER
# ---------------------------------------------------------
# Builds registration, identity (C1), and assessment-notes data straight
# from live SMApply review responses, replacing what additional_content
# used to supply for these sections (personal_info and application_info
# still come from additional_content, since reviews don't carry them).
# Education (A2/B1) is handled separately: prior-report PDF first, this
# same fetched review list as a fallback (see parse_manual_education_
# from_reviews) -- education keeps its two-step lookup rather than
# switching to reviews-only, per current design.
# =========================================================

# Review tasks are matched by NAME (normalized), not by hardcoded task id,
# because the ids differ per application / round.
TASK_NAME_IDENTITY     = "identity document review"
TASK_NAME_EDUCATION    = "education document review"
TASK_NAME_REGISTRATION = "registration document review"
TASK_NAME_NOTES        = "assessment notes"

# -----------------------------------------------------------------
# !!! VERIFY BEFORE TRUSTING !!!
# SMApply returns yes/no answers as the integer choice-INDEX, and the
# index->label mapping is NOT included in the reviews payload. Testing
# against real data showed the "ever suspended/restricted" question
# returns 1 for licenses that are clearly clean ("Unencumbered"), which
# means a single {0:No,1:Yes} map is WRONG for at least that question.
# Confirm each map against a known-good report before relying on this in
# production; until verified, treat these two fields' rendered values
# with suspicion.
# -----------------------------------------------------------------
LIFETIME_FLAG_MAP   = {0: "Yes", 1: "No"}   # SMApply choice order [Yes, No] -> index 1 = No
DISCIPLINE_FLAG_MAP = {0: "Yes", 1: "No"}   # SMApply choice order [Yes, No] -> index 1 = No

# ---------------------------------------------------------
# RESTORED FROM SCRIPT 1: only documents with Document Status == 4
# (accepted) are kept (registration and education). Blocks with any
# other status value (e.g. Waived, Rejected) are dropped; blocks with no
# Document Status field are kept (don't drop on missing data). This is
# IN ADDITION TO the "Add another?" stop-gating below -- both gates must
# pass for a block to survive.
#
# !!! VERIFY BEFORE TRUSTING !!! same caveat as the yes/no maps above:
# confirm the reviews API actually returns Document Status as an integer
# choice-index (with 4 == "Meets Requirements"/accepted) rather than a
# string label, before relying on this in production. If it comes back
# as a non-numeric string, _status_accepted() currently fails open
# (keeps the block) rather than silently dropping good data.
# -----------------------------------------------------------------
FILTER_BY_STATUS         = True
ACCEPTED_DOCUMENT_STATUS = {4}


def _status_accepted(bm: Dict[str, Any]) -> bool:
    """True if this block should be kept on status grounds. Keep only when
    Document Status == 4 (accepted); missing status is kept, any other value
    is dropped."""
    if not FILTER_BY_STATUS:
        return True
    v = _get(bm, "Document Status")
    if v is None or v == "":
        return True
    try:
        return int(v) in ACCEPTED_DOCUMENT_STATUS
    except (TypeError, ValueError):
        return True


def _nl(s: Any) -> str:
    return re.sub(r"\s+", " ", (s or "").strip()).lower()


def _fmt_review_date(v: Any) -> str:
    """'2026-06-30T00:00:00' -> '2026/06/30' to match the report's date style."""
    s = safe_str(v)
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", s)
    return f"{m.group(1)}/{m.group(2)}/{m.group(3)}" if m else s


def _split_blocks(responses: List[dict], sentinel_prefix: str) -> List[Tuple[List[Tuple[str, Any]], Any]]:
    """Segment a task's flat responses list into per-document blocks on the
    'Add another ...' delimiter.

    RESTORED FROM SCRIPT 1: returns (block_fields, terminator_value) per
    block, where terminator_value is the answer to the 'Add another ...?'
    question that closed the block (None for the final, unterminated
    block). The terminator is used by _is_stop() below to gate ingestion
    of the block that immediately follows.
    """
    sp = _nl(sentinel_prefix)
    blocks: List[Tuple[List[Tuple[str, Any]], Any]] = []
    cur: List[Tuple[str, Any]] = []
    for r in responses or []:
        if _nl(r.get("label")).startswith(sp):
            blocks.append((cur, r.get("response")))
            cur = []
        else:
            cur.append((r.get("label"), r.get("response")))
    if cur:
        blocks.append((cur, None))
    return blocks


def _bmap(block: List[Tuple[str, Any]]) -> Dict[str, Any]:
    """Normalized-label -> value for one block. First occurrence wins
    (duplicate labels like 'If Other, please specify:' are irrelevant here)."""
    d: Dict[str, Any] = {}
    for label, val in block:
        d.setdefault(_nl(label), val)
    return d


def _get(bm: Dict[str, Any], *candidates: str) -> Any:
    """Fetch a value by trying candidate labels; falls back to prefix match
    to survive truncated API labels (e.g. '...suspended/restri...')."""
    for c in candidates:
        nc = _nl(c)
        if nc in bm:
            return bm[nc]
    for c in candidates:
        nc = _nl(c).rstrip(".")
        for k, v in bm.items():
            if k.startswith(nc[:30]):
                return v
    return None


def _map_yesno(v: Any, m: Dict[int, str]) -> str:
    if v is None or v == "":
        return ""
    if isinstance(v, int):
        return m.get(v, str(v))
    return safe_str(v)


def _is_stop(add_another: Any) -> bool:
    """RESTORED FROM SCRIPT 1: 'Add another ...?' answered No -> the block
    that immediately follows is dropped. No is choice index 1
    ([Yes, No]); also tolerate string forms. NOTE: this gates only the
    NEXT block, not the rest of the list. A later Yes re-enables ingestion
    (skip-and-resume)."""
    if add_another is None or add_another == "":
        return False
    try:
        return int(add_another) == 1
    except (TypeError, ValueError):
        return _nl(add_another) in ("no", "n", "false")


def _reviews_by_task(reviews: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    out: Dict[str, Dict[str, Any]] = {}
    for rv in reviews:
        out.setdefault(_nl((rv.get("task") or {}).get("name")), rv)
    return out


# Confirmed mapping (2026-07-16): 0 = Yes, 1 = No for this specific
# question. Unlike LIFETIME_FLAG_MAP / DISCIPLINE_FLAG_MAP above, this one
# has been verified against real data and is not a guess.
PRIOR_REPORT_FLAG_MAP = {0: "Yes", 1: "No"}
_PRIOR_REPORT_FLAG_LABEL_PREFIX = "does this applicant have a previously issued advisory"


def _get_prior_report_flag(reviews: List[Dict[str, Any]]) -> Optional[str]:
    """
    Look up the reviewer's "Does this applicant have a previously issued
    advisory report..." Yes/No answer from the Education Document Review
    task. Returns "Yes", "No", or None if unanswered or not found.
    """
    ed_rv = _reviews_by_task(reviews).get(TASK_NAME_EDUCATION)
    if not ed_rv:
        return None
    for r in ed_rv.get("responses") or []:
        if _nl(r.get("label")).startswith(_PRIOR_REPORT_FLAG_LABEL_PREFIX):
            mapped = _map_yesno(r.get("response"), PRIOR_REPORT_FLAG_MAP)
            return mapped or None
    return None


def build_payload_from_reviews(reviews: List[Dict[str, Any]],
                               base_payload: Dict[str, Any]) -> Tuple[Dict[str, Any], Dict[str, Any]]:
    """Build the render `payload` from already-fetched SMApply reviews.

    `base_payload` is the parsed additional_content, used ONLY for
    personal_info and application_info (which reviews do not carry).

    Returns (payload, review_debug).
    """
    by = _reviews_by_task(reviews)

    payload: Dict[str, Any] = {
        "application_info": base_payload.get("application_info") or {},
        "personal_info":    base_payload.get("personal_info") or {},
        "registration":     {},
        "assessment_notes": {},
        "part_c_sections":  [],
    }

    # ---- Registration Document Review -> registration.entry_N + C3 ----
    # RESTORED FROM SCRIPT 1: "Add another registration?" stop-gating
    # (prev_allows / _is_stop) AND Document Status filtering
    # (_status_accepted) are both applied here. A block is kept only if
    # it passes BOTH gates.
    reg_rv = by.get(TASK_NAME_REGISTRATION)
    c3_bullets: List[str] = []
    if reg_rv:
        i = 0
        prev_allows = True   # first block is always eligible
        for block, add_another in _split_blocks(reg_rv.get("responses"), "add another registration"):
            allowed = prev_allows
            # This block's "Add another?" answer gates ONLY the next block.
            # No -> next block dropped; a later Yes re-enables (skip-and-resume).
            prev_allows = not _is_stop(add_another)
            if not allowed:
                continue

            bm = _bmap(block)
            board = safe_str(_get(bm, "Name of Registration Board/Authority"))
            num   = safe_str(_get(bm, "Registration/license number"))
            if not (board or num):          # empty / abandoned block
                continue
            if not _status_accepted(bm):    # keep only Document Status == 4
                continue

            i += 1
            payload["registration"][f"entry_{i}"] = {
                f"sa_jurisdiction_name_{i}":            board,
                f"sa_jurisdiction_address_{i}":         safe_str(_get(bm, "Address of Registration Board/Authority")),
                f"sa_registrant_name_issued_{i}":       safe_str(_get(bm, "Applicant Name when registered/licensed")),
                f"sa_registration_title_{i}":           safe_str(_get(bm, "Title of registration/license")),
                f"sa_registration_number_{i}":          num,
                f"sa_registration_method_{i}":          safe_str(_get(bm, "Method by which this applicant was registered")),
                f"sa_exam_language_{i}":                safe_str(_get(bm, "Language of examination")),
                f"sa_registration_issued_date_{i}":     _fmt_review_date(_get(bm, "Date registration/license was issued")),
                f"sa_registration_lifetime_flag_{i}":   _map_yesno(_get(bm, "Lifetime practice registration/license"), LIFETIME_FLAG_MAP),
                f"sa_registration_expiry_date_{i}":     _fmt_review_date(_get(bm, "Date license/registration expired/expires")),
                f"sa_registration_status_{i}":          safe_str(_get(bm, "Status of license/registration")),
                f"sa_registration_discipline_flag_{i}": _map_yesno(_get(bm, "Has the license/registration ever been suspended"), DISCIPLINE_FLAG_MAP),
            }
            recv = _fmt_review_date(_get(bm, "Document Received Date"))
            line = board or ""
            if num:  line = f"{line}: {num}" if line else num
            if recv: line = f"{line} - {recv}" if line else recv
            if line: c3_bullets.append(line)

    # ---- Identity Document Review -> C1 ----
    # (Script 1 did not apply stop-gating or status filtering to identity
    # documents either, so that behavior is preserved unchanged here.)
    id_rv = by.get(TASK_NAME_IDENTITY)
    c1_bullets: List[str] = []
    if id_rv:
        for block, _add in _split_blocks(id_rv.get("responses"), "add another identity"):
            bm = _bmap(block)
            dtype = safe_str(_get(bm, "Document Type"))
            dnum  = safe_str(_get(bm, "Identity Document Number"))
            recv  = _fmt_review_date(_get(bm, "Document Received Date"))
            if not (dtype or dnum):
                continue
            parts = [p for p in [dtype, dnum, recv] if p]
            if parts:
                c1_bullets.append(" - ".join(parts))

    # ---- C4 no longer comes from the "Education Document Review" task
    # ---- at all: it's now built purely from the "Additional documents"
    # ---- SharePoint folder in main() via
    # ---- list_additional_documents_files_by_subfolder() /
    # ---- apply_c4_additional_documents_files(). Nothing to build here.

    if c1_bullets:
        payload["part_c_sections"].append({"header": "C1: Identity Documents", "bullets": c1_bullets})
    if c3_bullets:
        payload["part_c_sections"].append({"header": "C3: Nursing Registration/License Documents", "bullets": c3_bullets})

    # ---- Assessment Notes ----
    n_rv = by.get(TASK_NAME_NOTES)
    if n_rv:
        bm = _bmap([(r.get("label"), r.get("response")) for r in n_rv.get("responses", [])])
        payload["assessment_notes"] = {
            "public":                    safe_str(_get(bm, "Public Assessment Note (All RB's)", "Public Assessment Note")),
            "assessment":                safe_str(_get(bm, "RB Assessment Note (RB-specific)", "RB Assessment Note")),
            "identity_section_note":     safe_str(_get(bm, "ID Section Note")),
            "education_section_note":    safe_str(_get(bm, "Education Section Note")),
            "registration_section_note": safe_str(_get(bm, "Registration Section Note")),
        }

    review_dbg = {
        "reviews_fetched":      len(reviews),
        "review_task_names":    [ (rv.get("task") or {}).get("name") for rv in reviews ],
        "matched_registration": reg_rv is not None,
        "matched_identity":     id_rv is not None,
        "matched_education":    by.get(TASK_NAME_EDUCATION) is not None,
        "matched_notes":        n_rv is not None,
        "registration_entries": len(payload["registration"]),
        "part_c_headers":       [s.get("header") for s in payload["part_c_sections"]],
    }
    return payload, review_dbg

# =========================================================
# CORE PDF HELPERS
# =========================================================

def build_overlay(w, h, draw_fn):
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=(w, h))
    draw_fn(c)
    c.showPage()
    c.save()
    buf.seek(0)
    return PdfReader(buf)

def wrap_lines(text, font, size, max_width):
    lines = []
    for para in (text or "").split("\n"):
        if para.strip() == "":
            lines.append("")
        else:
            lines.extend(simpleSplit(para, font, size, max_width))
    return lines

def styled_table(data, col_widths, header_rows=0):
    table = Table(data, colWidths=col_widths)
    style = [
        ("BOX",          (0, 0), (-1, -1), OUTER_BORDER_WIDTH, colors.black),
        ("INNERGRID",    (0, 0), (-1, -1), INNER_BORDER_WIDTH, colors.black),
        ("LEFTPADDING",  (0, 0), (-1, -1), CELL_PADDING),
        ("RIGHTPADDING", (0, 0), (-1, -1), CELL_PADDING),
        ("TOPPADDING",   (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 4),
        ("VALIGN",       (0, 0), (-1, -1), "TOP"),
    ]
    if header_rows:
        style.append(("BACKGROUND", (0, 0), (-1, header_rows - 1), COLOR_HEADER_BG))
    for row in range(header_rows, len(data)):
        if (row - header_rows) % 2 == 1:
            style.append(("BACKGROUND", (0, row), (-1, row), COLOR_ALT_ROW))
    table.setStyle(TableStyle(style))
    return table

def normalize_yes_no(v):
    if v is None: return ""
    if isinstance(v, bool): return "Yes" if v else "No"
    s = str(v).strip()
    if s.lower() in ("yes", "y", "true", "1"): return "Yes"
    if s.lower() in ("no", "n", "false", "0"): return "No"
    return s

# =========================================================
# HTML STRIPPING / VALUE NORMALIZATION
# =========================================================

_TAG_RE      = re.compile(r"<[^>]+>")
_LI_CLOSE_RE = re.compile(r"</li\s*>", re.IGNORECASE)

def strip_html(s):
    if not s: return ""
    s = s.replace("<br/>", "\n").replace("<br />", "\n").replace("<br>", "\n")
    s = _LI_CLOSE_RE.sub("\n", s)
    s = _TAG_RE.sub("", s)
    lines = [ln.strip() for ln in s.splitlines()]
    return "\n".join(l for l in lines if l).strip()

def safe_str(v):
    if v is None: return ""
    if isinstance(v, dict):
        for k in ("label", "value", "name", "text"):
            if k in v and v[k] is not None:
                return strip_html(str(v[k]).strip())
        return ""
    return strip_html(str(v).strip())

# ---------------------------------------------------------
# additional_content sometimes arrives with a literal, unescaped
# double-quote embedded INSIDE an HTML text value (e.g. an assessment
# note that quotes a phrase like: selected "Practical Nursing " for the
# program category). json.loads treats that embedded quote as the end
# of the string, which throws it off the rails at that exact point
# (typically surfacing as "Expecting ',' delimiter" a few characters
# later, once it hits whatever follows the quoted phrase).
#
# _repair_unescaped_quotes is a heuristic repair pass for exactly this
# case: it locates each "key": "value" pair using a lookahead that only
# accepts the closing quote when it's immediately followed by the start
# of the NEXT key (", "nextkey":) or a closing brace/bracket -- i.e. a
# real structural boundary, not just any bare quote. Any quote found
# before that true boundary is inside the value and gets escaped.
#
# This assumes keys are simple quoted identifiers (true throughout this
# payload's schema) and that no value legitimately contains a substring
# shaped exactly like `, "key":` immediately followed by more quoted
# text -- a safe bet for this HTML-note content, but still a patch over
# malformed input rather than a real JSON parser. If there's ever
# control over whatever assembles additional_content upstream, the
# durable fix is to have that system emit the value through a proper
# JSON serializer (e.g. json.dumps) instead of hand-concatenating text
# into a JSON literal, which auto-escapes embedded quotes and makes
# this repair pass unnecessary.
# ---------------------------------------------------------

_REPAIR_KV_RE = re.compile(
    r'("(?:[^"\\]|\\.)*"\s*:\s*)"(.*?)"'
    r'(?=\s*(?:,\s*"(?:[^"\\]|\\.)*"\s*:|[\r\n\s]*[}\]]))',
    re.DOTALL,
)


def _repair_unescaped_quotes(s: str) -> str:
    """Escape stray double-quotes found inside JSON string values, using
    'the next key or a closing brace/bracket' as the true end-of-value
    boundary instead of trusting the first bare quote encountered."""
    def _fix(m):
        key_part, value = m.group(1), m.group(2)
        fixed = re.sub(r'(?<!\\)"', r'\\"', value)
        fixed = fixed.replace("\r\n", "\\n").replace("\n", "\\n")
        return f'{key_part}"{fixed}"'
    return _REPAIR_KV_RE.sub(_fix, s)


def parse_additional_content(additional_content):
    dbg = {
        "received_additional_content_type": str(type(additional_content)),
        "received_additional_content_preview": (
            (additional_content[:220] + "...")
            if isinstance(additional_content, str) and len(additional_content) > 220
            else additional_content
        ),
    }
    if not additional_content or not isinstance(additional_content, str):
        raise ValueError("additional_content must be a non-empty JSON string.")

    s = additional_content.strip()

    def _try_parse(text):
        parsed = json.loads(text)
        if isinstance(parsed, str):
            parsed = json.loads(parsed)
        return parsed

    try:
        raw = _try_parse(s)
    except json.JSONDecodeError as e1:
        raw = None

        # Fallback 1: whole payload double-encoded as a quoted string.
        if (s.startswith('"') and s.endswith('"')) or (s.startswith("'") and s.endswith("'")):
            try:
                raw = _try_parse(s[1:-1])
                dbg["parse_fallback"] = "outer_quote_stripped"
            except json.JSONDecodeError:
                raw = None

        # Fallback 2: repair unescaped quotes embedded in a value (the
        # "selected \"Practical Nursing \"" style failure).
        if raw is None:
            try:
                repaired = _repair_unescaped_quotes(s)
                raw = _try_parse(repaired)
                dbg["parse_fallback"] = "unescaped_quotes_repaired"
            except json.JSONDecodeError as e2:
                raise ValueError(
                    "additional_content is not valid JSON, even after repair "
                    f"attempts. Original error: {e1}. Repair attempt error: {e2}"
                ) from e2

    if not isinstance(raw, dict):
        raise ValueError(f"Parsed additional_content is not a dict. Got: {type(raw)}")

    dbg.setdefault("parse_fallback", None)
    dbg["parsed_type"] = str(type(raw))
    dbg["parsed_keys_sample"] = sorted(list(raw.keys()))[:12]
    return raw, dbg

def join_address_from_flat(prefix, pi):
    a1      = safe_str(pi.get(f"{prefix}address_line1"))
    a2      = safe_str(pi.get(f"{prefix}address_line2"))
    city    = safe_str(pi.get(f"{prefix}city"))
    prov    = safe_str(pi.get(f"{prefix}province"))
    pc      = safe_str(pi.get(f"{prefix}postal_code"))
    country = safe_str(pi.get(f"{prefix}country"))
    lines   = [l for l in [a1, a2] if l]
    line3   = " ".join(p for p in [city, prov, pc] if p).strip()
    if line3:   lines.append(line3)
    if country: lines.append(country)
    return "\n".join(lines)

# =========================================================
# APPLICATION INFO / REGULATORY BODY LOOKUP
# =========================================================

NURSING_GROUP_DISPLAY = {
    "RN":  "Registered Nurse (RN)",
    "LPN": "Licensed Practical Nurse (LPN)",
}

RB_PROVINCE_MAP = {
    "NL College of Nurses":                                                      "Newfoundland and Labrador",
    "PEI College of Nursing & Midwifery":                                        "Prince Edward Island",
    "College of Nursing of New Brunswick":                                       "New Brunswick",
    "Nova Scotia College of Nursing":                                            "Nova Scotia",
    "College of Nurses of Ontario":                                              "Ontario",
    "College of Registered Nurses of Manitoba":                                  "Manitoba",
    "College of Registered Nurses of Saskatchewan":                              "Saskatchewan",
    "College of Registered Nurses of Alberta":                                   "Alberta",
    "BC College of Nurses and Midwives":                                         "British Columbia",
    "Association of New Brunswick Licensed Practical Nurses":                    "New Brunswick",
    "College of Licensed Practical Nurses of Manitoba":                          "Manitoba",
    "College of Licensed Practical Nurses of Saskatchewan":                      "Saskatchewan",
    "College of Licensed Practical Nurses and Health Care Aides of Alberta":     "Alberta",
}

RB_SHORTNAME_MAP = {
    "NL College of Nurses":                                                      "NLCN",
    "PEI College of Nursing & Midwifery":                                        "PEICNM",
    "College of Nursing of New Brunswick":                                       "CNNB",
    "Nova Scotia College of Nursing":                                            "NSCN",
    "College of Nurses of Ontario":                                              "CNO",
    "College of Registered Nurses of Manitoba":                                  "CRNM",
    "College of Registered Nurses of Saskatchewan":                              "CRNS",
    "College of Registered Nurses of Alberta":                                   "CRNA",
    "BC College of Nurses and Midwives":                                         "BCCNM",
    "Association of New Brunswick Licensed Practical Nurses":                    "ANBLPN",
    "College of Licensed Practical Nurses of Manitoba":                          "CLPNM",
    "College of Licensed Practical Nurses of Saskatchewan":                      "CLPNS",
    "College of Licensed Practical Nurses and Health Care Aides of Alberta":     "CLHA",
}

def rb_short_name(rb_full: str) -> str:
    rb_full = (rb_full or "").strip()
    if rb_full in RB_SHORTNAME_MAP:
        return RB_SHORTNAME_MAP[rb_full]
    for k, v in RB_SHORTNAME_MAP.items():
        if k.lower() == rb_full.lower():
            return v
    acronym = "".join(w[0] for w in re.split(r"\s+", rb_full) if w and w[0].isalpha()).upper()
    acronym = re.sub(r"[^A-Za-z0-9_-]", "", acronym)
    return acronym or "UNKNOWN_RB"

def build_application_info(payload: Dict[str, Any],
                            application_number: str = "",
                            nnas_id: str = "") -> List[Dict[str, str]]:
    ai = payload.get("application_info") or {}
    results = []
    for group_code, raw in (
        ("RN",  safe_str(ai.get("regulatory_body_rn"))),
        ("LPN", safe_str(ai.get("regulatory_body_lpn"))),
    ):
        group_display = NURSING_GROUP_DISPLAY.get(group_code, group_code)
        for rb in [b.strip() for b in raw.split(",") if b.strip()]:
            results.append({
                "nursing_group":      group_display,
                "nursing_group_code": group_code,
                "regulatory_body":    rb,
                "province":           RB_PROVINCE_MAP.get(rb, ""),
                "nnas_id":            nnas_id,
                "application_number": application_number,
            })
    return results

def parse_application_number(title: Optional[str]) -> str:
    if not title: return ""
    return title.strip()

# =========================================================
# SHARED DATA BUILDERS
# =========================================================

def build_personal_info(payload: Dict[str, Any]) -> Dict[str, str]:
    pi        = payload.get("personal_info") or {}
    full_name = " ".join(
        p for p in [safe_str(pi.get("sa_first_name")),
                    safe_str(pi.get("sa_middle_name")),
                    safe_str(pi.get("sa_last_name"))] if p
    ).strip()
    same_flag    = normalize_yes_no(pi.get("sa_mailing_same_as_current_flag"))
    current_addr = join_address_from_flat("sa_current_", pi)
    mailing_addr = join_address_from_flat("sa_mailing_", pi)
    return {
        "name":                      full_name,
        "other_names":               safe_str(pi.get("sa_other_names")),
        "dob":                       safe_str(pi.get("sa_dob")),
        "initial_education_country": safe_str(pi.get("sa_initial_education_country")),
        "mailing_address":           current_addr if same_flag == "Yes" else mailing_addr,
        "phone":                     safe_str(pi.get("sa_phone")),
        "email":                     safe_str(pi.get("sa_email")),
    }

def personal_info_rows(payload: Dict[str, Any]) -> List[Tuple[str, str]]:
    d = build_personal_info(payload)
    return [
        ("Name:",                                                       d["name"]),
        ("Other name(s):",                                              d["other_names"]),
        ("Date of birth:",                                              d["dob"]),
        ("Country where applicant received initial nursing education:", d["initial_education_country"]),
        ("Mailing address:",                                            d["mailing_address"]),
        ("Phone number:",                                               d["phone"]),
        ("E-mail address:",                                             d["email"]),
    ]

def extract_registration_entries(payload: Dict[str, Any]) -> List[Dict[str, Any]]:
    reg     = payload.get("registration") or {}
    entries = []
    for i in range(1, 8):
        e     = reg.get(f"entry_{i}") or {}
        if not isinstance(e, dict): continue
        jname = safe_str(e.get(f"sa_jurisdiction_name_{i}"))
        lic   = safe_str(e.get(f"sa_registration_number_{i}"))
        if not jname and not lic: continue
        entries.append({
            "index": i,
            "jurisdiction_name":            jname,
            "jurisdiction_address":         safe_str(e.get(f"sa_jurisdiction_address_{i}")),
            "registrant_name_issued":       safe_str(e.get(f"sa_registrant_name_issued_{i}")),
            "registration_title":           safe_str(e.get(f"sa_registration_title_{i}")),
            "registration_number":          lic,
            "registration_method":          safe_str(e.get(f"sa_registration_method_{i}")),
            "exam_language":                safe_str(e.get(f"sa_exam_language_{i}")),
            "registration_issued_date":     safe_str(e.get(f"sa_registration_issued_date_{i}")),
            "registration_lifetime_flag":   safe_str(e.get(f"sa_registration_lifetime_flag_{i}")),
            "registration_expiry_date":     safe_str(e.get(f"sa_registration_expiry_date_{i}")),
            "registration_status":          safe_str(e.get(f"sa_registration_status_{i}")),
            "registration_discipline_flag": safe_str(e.get(f"sa_registration_discipline_flag_{i}")),
        })
    return entries

def registration_rows(reg: Dict[str, Any]) -> List[Tuple[str, str]]:
    return [
        ("Name of jurisdiction / registering board / authority:",
         reg.get("jurisdiction_name", "")),
        ("Current address of jurisdiction / registering board / authority:",
         reg.get("jurisdiction_address", "")),
        ("Applicant name when registration / license / diploma was issued:",
         reg.get("registrant_name_issued", "")),
        ("Title of registration / license:",
         reg.get("registration_title", "")),
        ("Registration / license number:",
         reg.get("registration_number", "")),
        ("Method by which this applicant was registered:",
         reg.get("registration_method", "")),
        ("Language of examination:",
         reg.get("exam_language", "")),
        ("Date registration / license issued:",
         reg.get("registration_issued_date", "")),
        ("Is this a lifetime practice registration / license:",
         reg.get("registration_lifetime_flag", "")),
        ("Date registration / license expired / expires:",
         reg.get("registration_expiry_date", "")),
        ("Status of registration / license:",
         reg.get("registration_status", "")),
        ("Has the applicant's registration / license ever been suspended, "
         "restricted, surrendered, revoked, or subject to individual terms "
         "and conditions:",
         reg.get("registration_discipline_flag", "")),
    ]

def build_part_c_sections(payload: Dict[str, Any]) -> List[Dict[str, Any]]:
    sections = payload.get("part_c_sections")
    if isinstance(sections, list) and sections:
        return sections
    sections = []
    csecs = payload.get("c_sections") or {}
    if not isinstance(csecs, dict): return sections

    c1 = csecs.get("c1_identity_documents") or {}
    if isinstance(c1, dict):
        bullets = []
        for key in ("entry_1", "entry_2"):
            d         = c1.get(key) or {}
            if not isinstance(d, dict): continue
            cat       = safe_str(d.get("category"))
            cat_other = safe_str(d.get("category_other"))
            doc_num   = safe_str(d.get("identity_document_number"))
            recv_date = safe_str(d.get("document_received_date"))
            cat_disp  = cat_other if cat.strip().lower() == "other" and cat_other else cat
            if not (cat_disp or doc_num or recv_date): continue
            parts = [p for p in [cat_disp, doc_num, recv_date] if p]
            if parts: bullets.append(" - ".join(parts))
        if bullets:
            sections.append({"header": "C1: Identity Documents", "bullets": bullets})

    c3 = csecs.get("c3_registration_documents") or {}
    if isinstance(c3, dict):
        bullets = []
        for i in range(1, 8):
            d         = c3.get(f"entry_{i}") or {}
            if not isinstance(d, dict): continue
            board     = safe_str(d.get("registration_board_authority"))
            regno     = safe_str(d.get("registration_license_number"))
            recv_date = safe_str(d.get("document_received_date"))
            if not (board or regno or recv_date): continue
            line = board or ""
            if regno:     line = f"{line}: {regno}" if line else regno
            if recv_date: line = f"{line} - {recv_date}" if line else recv_date
            if line: bullets.append(line)
        if bullets:
            sections.append({"header": "C3: Nursing Registration/License Documents", "bullets": bullets})

    c4 = csecs.get("c4_other_documents") or {}
    if isinstance(c4, dict):
        bullets = []
        for i in range(1, 6):
            d         = c4.get(f"entry_{i}") or {}
            if not isinstance(d, dict): continue
            doc_name  = safe_str(d.get("document_name"))
            recv_date = safe_str(d.get("document_received_date"))
            if not (doc_name or recv_date): continue
            line = doc_name or ""
            if recv_date: line = f"{line} - {recv_date}" if line else recv_date
            if line: bullets.append(line)
        if bullets:
            sections.append({"header": "C4: Other Documents", "bullets": bullets})

    return sections

# =========================================================
# PRIOR REPORT EXTRACTION  (pdfplumber — text-based, no OCR)
#
# A2 and B1 are located by scanning every page for their section header
# text rather than relying on a hardcoded page number.
# =========================================================

_A2_PATTERN     = re.compile(r"A2[:\s]*NURSING EDUCATION OVERVIEW", re.IGNORECASE)
# NOTE: matched loosely on "...EDUCATION DOMAIN BREAKDOWN" rather than
# hardcoding "RN / PN" in the middle. The static template header text
# varies by nursing group -- e.g. "B1: RN EDUCATION DOMAIN BREAKDOWN" for
# an RN-only applicant, "B1: RN / PN EDUCATION DOMAIN BREAKDOWN" when both
# groups apply, and possibly "B1: PN EDUCATION DOMAIN BREAKDOWN" /
# "B1: LPN EDUCATION DOMAIN BREAKDOWN" for PN/LPN-only applicants. The old
# pattern only matched the "RN / PN" form, so it silently failed to detect
# the B1 section (and therefore extracted ZERO domain-breakdown rows -- not
# just "only the first one") on any report where the header read
# differently, and as a side effect never told _extract_a2_raw_rows where
# to stop, letting it sweep unrelated later pages into the A2 data too.
_B1_PATTERN     = re.compile(r"B1[:\s].*?EDUCATION DOMAIN BREAKDOWN", re.IGNORECASE | re.DOTALL)
_PART_B_PATTERN = re.compile(r"Part\s*B[:\s]*Nursing Education Analysis", re.IGNORECASE)
_B1_HEADER_SUBJECTS = {"SUBJECT AREA", "THEORY\nHOURS", "SIMULATION/\nLAB\nHOURS"}

def _plumber_page_text(page) -> str:
    return page.extract_text() or ""

def _parse_a2_table(raw_rows) -> Dict[str, str]:
    result: Dict[str, str] = {}
    for row in raw_rows:
        if not row or len(row) < 2:
            continue
        label = (row[0] or "").replace("\n", " ").strip()
        value = (row[1] or "").strip()
        if label:
            result[label] = value
    return result

def _is_b1_header_row(row) -> bool:
    return any((cell or "").strip() in _B1_HEADER_SUBJECTS for cell in row)

def _parse_b1_row(row) -> Optional[Dict[str, str]]:
    if not row or len(row) < 4:
        return None
    subject = (row[0] or "").replace("\n", " ").strip()
    if not subject or _is_b1_header_row(row):
        return None
    def clean(v):
        s = (v or "").replace("\n", " ").strip()
        return "" if s == "---" else s
    return {
        "subject":        subject,
        "theory_hours":   clean(row[1]),
        "sim_lab_hours":  clean(row[2]),
        "clinical_hours": clean(row[3]),
    }

def _a2_key(raw: Dict[str, str], *fragments: str) -> str:
    for key, val in raw.items():
        for frag in fragments:
            if frag.lower() in key.lower():
                return val
    return ""

# =========================================================
# PRIOR REPORT EXTRACTION — MULTI-PROGRAM SEGMENTATION
# ---------------------------------------------------------
# Some prior reports append multiple A2/B1 program blocks directly one
# after another (no intervening "Part C" or other section marker), e.g.
# two schools stacked in the same table. The old implementation
# collapsed every row into a single dict (A2) / deduped by subject name
# (B1), which silently merges or drops all but one program. These
# helpers instead segment on natural block boundaries:
#   - A2: a new block starts every time the row labeled
#     "Name of school / educational institution:" reappears.
#   - B1: a block ends every time a "TOTALS" row is seen; the next row
#     starts a new block.
# =========================================================

_A2_NAME_ROW_LABEL = "name of school"


def _extract_a2_raw_rows(pdf) -> List[Tuple[str, str]]:
    """
    Collect every (label, value) row across all A2 pages, in document
    order, starting at the first page whose text matches _A2_PATTERN and
    continuing until a page's text matches _B1_PATTERN (that page's
    tables are NOT included — B1 extraction picks up from there).
    """
    collecting = False
    rows: List[Tuple[str, str]] = []
    for page in pdf.pages:
        text = _plumber_page_text(page)
        if not collecting and _A2_PATTERN.search(text):
            collecting = True
        if not collecting:
            continue
        if _B1_PATTERN.search(text) and rows:
            break
        for tbl in page.extract_tables():
            for row in tbl:
                if not row or len(row) < 2:
                    continue
                label = (row[0] or "").replace("\n", " ").strip()
                # NOTE: `value` is kept RAW here, embedded "\n" and all --
                # do NOT collapse it to a single line at extraction. Some
                # source rows deliberately encode two separate answers in
                # one cell via a line break (e.g. "Language of instruction"
                # stores theory on line 1, clinical on line 2 -- see the
                # lang_field.split("\n") logic in extract_nnas_education
                # below, which depends on that literal newline surviving
                # to this point). Collapsing here previously broke that
                # split silently: "English\nEnglish" became the single
                # string "English English", so both theory AND clinical
                # ended up rendering that same doubled text.
                #
                # A DIFFERENT wrapped-cell artifact (e.g. a long school
                # name wrapping across two lines inside its table cell,
                # coming back from pdfplumber as "Genesis Institute of
                # Nursing (FKA\nAnil Baghi College/School of Nursing)")
                # is real too, but is handled at its actual point of
                # failure instead: a2_identifying_line()'s caption is
                # drawn with plain drawString(), which is NOT newline-
                # aware (Helvetica/WinAnsi has no glyph for U+000A, so
                # reportlab silently substitutes a .notdef glyph -- a
                # solid black square, right where the line wrap was).
                # render_a2_on_page's reportlab Paragraph tolerates a raw
                # "\n" fine, so no fix is needed there. See
                # a2_identifying_line() for the targeted whitespace
                # collapse on just that one consumer.
                value = (row[1] or "").strip()
                if label:
                    rows.append((label, value))
    return rows


def _segment_a2_rows(rows: List[Tuple[str, str]]) -> List[Dict[str, str]]:
    """
    Split a flat list of A2 rows into one dict per program, breaking
    right before every recurrence of the "Name of school / educational
    institution:" row (except the very first one, which opens the first
    block).
    """
    blocks: List[List[Tuple[str, str]]] = []
    current: List[Tuple[str, str]] = []
    for label, value in rows:
        if _nl(label).startswith(_A2_NAME_ROW_LABEL) and current:
            blocks.append(current)
            current = []
        current.append((label, value))
    if current:
        blocks.append(current)

    programs: List[Dict[str, str]] = []
    for block in blocks:
        d: Dict[str, str] = {}
        for label, value in block:
            d[label] = value
        programs.append(d)
    return programs


def _extract_a2_from_pdf(pdf) -> List[Dict[str, str]]:
    """Return a list of raw A2 label->value dicts, one per program block."""
    rows = _extract_a2_raw_rows(pdf)
    return _segment_a2_rows(rows)


def _extract_b1_from_pdf(pdf) -> List[List[Dict[str, str]]]:
    """
    Collect B1 domain rows across potentially multiple pages, segmented
    into one block per program. A block closes as soon as a "TOTALS" row
    is seen; the next parsed row starts a new block. Within each block,
    rows are deduped by subject (first occurrence wins) to protect
    against repeated header/footer artifacts on multi-page tables.
    """
    collecting = False
    blocks: List[List[Dict[str, str]]] = []
    current: List[Dict[str, str]] = []
    for page in pdf.pages:
        text = _plumber_page_text(page)
        if _B1_PATTERN.search(text) or (_PART_B_PATTERN.search(text) and collecting):
            collecting = True
        if not collecting:
            continue
        if re.search(r"Part\s*C", text, re.IGNORECASE):
            break
        for tbl in page.extract_tables():
            for raw_row in tbl:
                parsed = _parse_b1_row(raw_row)
                if not parsed:
                    continue
                current.append(parsed)
                if parsed["subject"].strip().upper() == "TOTALS":
                    blocks.append(current)
                    current = []
    if current:
        blocks.append(current)

    deduped_blocks: List[List[Dict[str, str]]] = []
    for block in blocks:
        seen: set = set()
        uniq: List[Dict[str, str]] = []
        for r in block:
            if r["subject"] not in seen:
                seen.add(r["subject"])
                uniq.append(r)
        deduped_blocks.append(uniq)
    return deduped_blocks


def extract_nnas_education(pdf_path: str) -> Dict[str, Any]:
    """
    Open a *_priorreport.pdf and extract EVERY A2/B1 program pair into a
    normalized {"programs": [...]} structure (see
    normalize_education_programs). Programs are paired by position: the
    Nth A2 block with the Nth B1 block. If the counts don't match (e.g.
    a school listed in A2 with no matching B1 table, or vice versa), the
    missing side is filled in as empty rather than dropping the whole
    program, so partial data still surfaces instead of disappearing.
    """
    from itertools import zip_longest

    with pdfplumber.open(pdf_path) as pdf:
        a2_dicts  = _extract_a2_from_pdf(pdf)
        b1_blocks = _extract_b1_from_pdf(pdf)

    programs: List[Dict[str, Any]] = []
    for raw_a2, raw_b1 in zip_longest(a2_dicts, b1_blocks, fillvalue=None):
        raw_a2 = raw_a2 or {}
        raw_b1 = raw_b1 or []

        lang_field    = _a2_key(raw_a2, "language of instruction")
        lang_parts    = [l.strip() for l in lang_field.split("\n") if l.strip()]
        lang_theory   = lang_parts[0] if len(lang_parts) > 0 else ""
        lang_clinical = lang_parts[1] if len(lang_parts) > 1 else lang_theory

        a2: Dict[str, str] = {
            "school_name":                   _a2_key(raw_a2, "name of school"),
            "school_address":                _a2_key(raw_a2, "current address of school"),
            "school_type":                   _a2_key(raw_a2, "type of school"),
            "primary_language":              _a2_key(raw_a2, "primary language"),
            "applicant_name_at_school":      _a2_key(raw_a2, "applicant name used"),
            "program_category":              _a2_key(raw_a2, "category of program"),
            "entrance_requirements":         _a2_key(raw_a2, "entrance requirements"),
            "program_length":                _a2_key(raw_a2, "length of study"),
            "program_start":                 _a2_key(raw_a2, "program start"),
            "program_end":                   _a2_key(raw_a2, "program completion", "graduation"),
            "instruction_language_theory":   lang_theory,
            "instruction_language_clinical": lang_clinical,
            "credential":                    _a2_key(raw_a2, "credential", "degree obtained"),
            "comments":                      _a2_key(raw_a2, "comments"),
        }

        a2_found = bool(a2 and any(v for v in a2.values()))
        b1_found = bool(raw_b1)
        if a2_found or b1_found:
            programs.append({"a2": a2, "b1": raw_b1})

    return {"programs": programs}

# =========================================================
# PRIOR REPORT EXTRACTION - PART C: NURSING EDUCATION DOCUMENTS
# ---------------------------------------------------------
# The vendor still submits a full, multi-line "Nursing Education
# Documents" list (diplomas, nursing education forms, transcripts, etc,
# each with its own received date) inside the prior report's Part C,
# even though the live-review pipeline (build_payload_from_reviews) now
# follows a narrower "one accepted document" convention for
# registration and education document metadata. That narrower
# convention exists ONLY because the live SMApply reviews carry a
# Document Status field to filter on (see _status_accepted); the prior
# report's Part C is static, already-rendered text with no structured
# status data behind it, so there's no reliable way to pick "the one
# that matters" out of it. Per team decision, this section is therefore
# ported forward from the prior report VERBATIM (all lines, not
# filtered to one), rather than being trimmed to match the newer
# one-document convention used elsewhere in this report.
#
# In THIS report's Part C numbering it is filed as "C2: Nursing
# Education Documents". The review-derived sections already come out as
# C1 (Identity), C3 (Registration), C4 (Other) -- that adapter never
# produced a C2 -- so the ported section slots in without needing to
# renumber anything else. The detection pattern below matches
# "C<any digit>: Nursing Education Documents" rather than hardcoding a
# number, since an older prior report may have filed this same section
# under a different heading number; only the NEW report's inserted
# label is fixed at C2.
# =========================================================

_EDU_DOCS_HEADER_PATTERN = re.compile(
    r"C\d+[:\s]*NURSING EDUCATION DOCUMENTS", re.IGNORECASE
)
_PART_C_SUBHEADER_PATTERN = re.compile(r"^C\d+\s*:", re.IGNORECASE)

# ---------------------------------------------------------
# NOTE ON BULLET DETECTION: render_part_c_last_page draws each bullet as
# f"\u25a0 {b}" (BLACK SQUARE + space) using the base Helvetica font.
# Helvetica's WinAnsi encoding has no glyph for U+25A0, so reportlab
# silently substitutes some OTHER single character at that code
# position instead -- confirmed empirically to render (and therefore
# get extracted by pdfplumber) as things like "n " or "o " rather than
# "\u25a0 ". Which substitute character shows up is a font/version
# artifact, not something to hardcode. So instead of matching a specific
# character, a bullet's *first* physical line is detected generically as
# "exactly one non-whitespace character, then whitespace, then an
# uppercase letter or digit" -- every genuine bullet start has this
# shape (real content in this list -- board/institution names, license
# numbers, document types -- always begins with a capital letter or
# digit), while a wrapped CONTINUATION line of a long bullet does not
# (wrap_lines only adds the marker+space once, to the first line).
# Known edge case: a bullet whose real text legitimately starts with a
# genuine one-letter word (e.g. "A ...") would be mis-stripped; this
# hasn't occurred in this document-list context (board names, license
# numbers, document types, dates) so the tradeoff is accepted.
# ---------------------------------------------------------
_LEADING_BULLET_ARTIFACT_RE = re.compile(r"^\S\s+(?=[A-Z0-9])")


def _extract_education_documents_from_pdf(pdf) -> List[str]:
    """
    Scan a *_priorreport.pdf for its "Nursing Education Documents"
    heading under Part C (whatever number it was filed under in that
    report) and return its bullet lines verbatim (bullet-marker artifact
    stripped, wrapped continuation lines re-joined). Stops at the next
    'Cn:' sub-heading. Returns [] if the heading or its bullets aren't
    found -- callers should treat that as "nothing to port," not an
    error, since older prior reports may not have this section at all.
    """
    collecting = False
    bullets: List[str] = []
    for page in pdf.pages:
        text = _plumber_page_text(page)
        if not text:
            continue
        for raw_line in text.splitlines():
            line = raw_line.strip()
            if not line:
                continue
            if _EDU_DOCS_HEADER_PATTERN.search(line):
                collecting = True
                continue
            if not collecting:
                continue
            if _PART_C_SUBHEADER_PATTERN.match(line):
                # Reached the next Part C sub-section.
                return bullets
            m = _LEADING_BULLET_ARTIFACT_RE.match(line)
            if m:
                cleaned = line[m.end():].strip()
                if cleaned:
                    bullets.append(cleaned)
            elif bullets:
                # Wrapped continuation of the previous bullet (long lines
                # get word-wrapped across multiple text lines without
                # repeating the bullet marker).
                bullets[-1] = f"{bullets[-1]} {line}".strip()
            else:
                # Defensive fallback: the first content line under the
                # header didn't match the expected artifact+space
                # pattern (e.g. a future font substitution behaves
                # differently than observed). Keep the line rather than
                # silently dropping it.
                bullets.append(line)
    return bullets


def extract_education_documents(pdf_path: str) -> List[str]:
    """Open a *_priorreport.pdf and return its education-document bullet
    lines. Returns [] on any extraction failure or if the section isn't
    present, rather than raising -- this is supplementary content, not
    required for report generation."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            return _extract_education_documents_from_pdf(pdf)
    except Exception:
        return []


def apply_part_c_education_documents(payload: Dict[str, Any],
                                     prior_pdf_path: Optional[str]) -> Dict[str, Any]:
    """
    Insert a "C2: Nursing Education Documents" section (ported verbatim
    from the prior report, when one is available) right after C1 in
    payload["part_c_sections"]. Mutates and returns `payload`.

    No renumbering of the existing review-derived sections is needed:
    build_payload_from_reviews already labels them C1 (Identity), C3
    (Registration), C4 (Other) -- it never produced a C2 -- so C2 slots
    in cleanly on its own.

    If no prior report is available, or it has no education-documents
    section, no C2 entry is added -- the report simply omits that
    section, same as before this change.
    """
    sections = payload.get("part_c_sections") or []

    education_bullets = extract_education_documents(prior_pdf_path) if prior_pdf_path else []

    if education_bullets:
        insert_at = len(sections)
        for idx, sec in enumerate(sections):
            if str(sec.get("header", "")).startswith("C1"):
                insert_at = idx + 1
                break
        sections.insert(insert_at, {
            "header":  "C2: Nursing Education Documents",
            "bullets": education_bullets,
        })

    payload["part_c_sections"] = sections
    return payload


# =========================================================
# MANUAL EDUCATION DATA (SMAPPLY REVIEWS FALLBACK)
#
# Used when no prior report PDF is found on SharePoint. A reviewer fills
# in up to 5 (Nursing Education Overview + Education Domain Breakdown)
# pairs directly in the SMApply "Education Document Review" task. Each
# pair lives in its own repeated section, driven by an
# "Add another nursing education overview and education domain breakdown"
# flag (0 = Yes/continue, 1 or unanswered = No/stop).
# =========================================================

DOMAIN_BREAKDOWN_SUBJECTS = [
    "Nursing Care of the Adult - Medical",
    "Nursing Care of the Adult - Surgical",
    "Maternal/Infant Nursing",
    "Gynecology",
    "Pediatric Nursing",
    "Gerontology/Geriatric Nursing",
    "Mental Health Nursing",
    "Community Health/Public Health Nursing",
    "Anatomy & Physiology",
    "Pathophysiology",
    "Microbiology",
    "Pharmacology & Medications",
    "Infusion Therapy Theory & Skills",
    "Nutrition",
    "Fundamentals of Nursing",
    "Health Assessment Across the Lifespan",
    "Leadership",
    "Ethical/Legal Practice",
    "Applied Research",
    "Primary Health Care",
    "TOTALS",
]

_ADD_ANOTHER_LABEL_PREFIX = "add another nursing education overview and education do"
_DOMAIN_BREAKDOWN_LABEL_PREFIX = "Education Domain Breakdown |"
_EDU_OVERVIEW_LABEL_PREFIX = "Nursing Education Overview |"


def _response_str(resp: Any) -> str:
    return "" if resp is None else str(resp).strip()


def _group_by_field_prefix(responses: List[Dict[str, Any]],
                           label_prefix: str) -> Tuple[List[str], Dict[str, Dict[int, str]]]:
    """
    Group review responses whose label starts with label_prefix by the
    prefix of their field id (the text before the trailing _N). Returns
    the group prefixes in encounter order, and a prefix -> {index: value}
    map.
    """
    groups: Dict[str, Dict[int, str]] = {}
    order: List[str] = []
    for r in responses:
        if not isinstance(r.get("label"), str) or not r["label"].startswith(label_prefix):
            continue
        field_id = r.get("id", "")
        m = re.match(r"^(.*)_(\d+)$", field_id)
        if not m:
            continue
        prefix, idx = m.group(1), int(m.group(2))
        if prefix not in groups:
            groups[prefix] = {}
            order.append(prefix)
        groups[prefix][idx] = _response_str(r.get("response"))
    return order, groups


def parse_manual_education_from_reviews(reviews: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Parse reviewer-entered education data out of an already-fetched flat
    list of SMApply review submissions (as returned by fetch_all_reviews).
    Each repeated section pairs one Nursing Education Overview (A2, 14
    fields, matching _A2_DISPLAY_ROWS order) with one Education Domain
    Breakdown (B1, 63 fields = 21 fixed subjects x
    [theory, sim/lab, clinical]). Walks sections in order, stopping after
    the first section whose "Add another..." flag != 0.

    Returns {"programs": [{"a2": {...}, "b1": [...]}, ...]} — every
    subject row is always included (blank values rendered as "---" by
    _b1_display downstream), and every non-blank section is kept, per
    spec: reviewers may fill in up to 5 education programs.
    """
    all_responses: List[Dict[str, Any]] = []
    for result in reviews:
        all_responses.extend(result.get("responses") or [])

    a2_order, a2_groups = _group_by_field_prefix(all_responses, _EDU_OVERVIEW_LABEL_PREFIX)
    b1_order, b1_groups = _group_by_field_prefix(all_responses, _DOMAIN_BREAKDOWN_LABEL_PREFIX)
    add_another_responses = [
        r for r in all_responses
        if _nl(r.get("label")).startswith(_ADD_ANOTHER_LABEL_PREFIX)
    ]

    n_sections = min(len(a2_order), len(b1_order))
    programs: List[Dict[str, Any]] = []

    for i in range(n_sections):
        a2_values = a2_groups[a2_order[i]]
        a2: Dict[str, str] = {
            key: a2_values.get(idx, "")
            for idx, (_, key) in enumerate(_A2_DISPLAY_ROWS)
        }

        b1_values = b1_groups[b1_order[i]]
        b1: List[Dict[str, str]] = []
        for subj_i, subject in enumerate(DOMAIN_BREAKDOWN_SUBJECTS):
            base = subj_i * 3
            b1.append({
                "subject":        subject,
                "theory_hours":   b1_values.get(base, ""),
                "sim_lab_hours":  b1_values.get(base + 1, ""),
                "clinical_hours": b1_values.get(base + 2, ""),
            })

        programs.append({"a2": a2, "b1": b1})

        add_another = (
            add_another_responses[i].get("response")
            if i < len(add_another_responses) else None
        )
        if add_another != 0:
            break

    return {"programs": programs}


# =========================================================
# APPENDED DOMAIN ROWS (evaluator-hardcoded rows for reports missing
# some of the 20 canonical B1 subjects)
# ---------------------------------------------------------
# Gated behind a Yes/No question on the Education Document Review task
# ("Would you like to append to an existing education domain
# breakdown?"), then repeatable up to 5 times -- confirmed against a
# real submission on 2026-08-07 -- mirroring the exact same repeat-
# until-No pattern already used by the manual Nursing Education
# Overview / Domain Breakdown flow (parse_manual_education_from_reviews)
# elsewhere in this same task:
#   Block 1: gated by "Would you like to append to an existing
#     education domain breakdown?" (_APPEND_TO_EXISTING_LABEL_PREFIX).
#   Blocks 2-5: each gated by its OWN "Append to another education
#     domain breakdown?" question (_APPEND_ANOTHER_LABEL_PREFIX) --
#     note this is a DIFFERENT label than block 1's gate, not a repeat
#     of the same one.
# Each block, regardless of position, is immediately followed by its
# own "Which education document...?" target dropdown and its own
# 10-row "Rows to add to the existing...?" grid.
#
# "Which education document...associated with?" is a dropdown whose
# answer is a raw 0-based integer choice index straight from SMApply --
# "Education Document 1" -> 0, "Education Document 2" -> 1, etc. --
# NOT text to parse. Confirmed against a real submission on 2026-08-07.
# This maps directly onto education_data["programs"] index order (same
# ordering used everywhere else in this script for A2/B1 pairing).
#
# !!! LABEL TRUNCATION !!!
# SMApply truncates long question labels with "..." in the responses
# payload (confirmed 2026-08-07). The prefixes below are trimmed to
# stop safely BEFORE each label's truncation point -- do not lengthen
# them without checking against a real payload first, since a prefix
# that runs past the truncation point will silently match nothing (no
# error -- _get()/_group_by_field_prefix just return empty, and the
# feature quietly never fires). _APPEND_ANOTHER_LABEL_PREFIX happens to
# not be truncated in the confirmed sample, but keep this in mind if it
# ever needs adjusting.
#
# The "Which education document...?" label is reused VERBATIM up to TEN
# times within this one task's responses: once per append block (up to
# 5) AND once per manually-filled program block (up to 5, unrelated
# feature). A dict-keyed first-occurrence lookup (_bmap/_get) cannot
# disambiguate these at all once there's more than one of each. Instead,
# every lookup below is anchored relative to THIS function's own gate
# responses via _find_response_index_after (index-based, never a bare
# label search across the whole task) -- that's what keeps the two
# flows, and the up-to-5 blocks within this flow, from colliding.
# -----------------------------------------------------------------
_APPEND_TO_EXISTING_LABEL_PREFIX = "would you like to append to an existing education"
_APPEND_ANOTHER_LABEL_PREFIX     = "append to another education domain breakdown"
_APPEND_TARGET_DOC_LABEL_PREFIX  = "which education document is this nursing education over"
_APPENDED_ROW_LABEL_PREFIX       = "Rows to add to the existing"  # case-sensitive: passed directly to _group_by_field_prefix, which does NOT run labels through _nl()

_APPENDED_ROW_COUNT      = 10  # confirmed: Subject 1 .. Subject 10 (40 fields, 4 per row)
_APPENDED_ROW_MAX_BLOCKS = 5   # confirmed cap, same as the manual program-entry flow's cap


def _find_response_index_after(responses: List[Dict[str, Any]], start_index: int,
                               label_prefix: str) -> Optional[int]:
    """Index of the first response after `start_index` whose normalized
    label starts with `label_prefix`, or None. Index-based (not the
    response object itself) so callers can keep chaining searches
    forward from wherever the previous match landed."""
    for i in range(start_index + 1, len(responses)):
        if _nl(responses[i].get("label")).startswith(label_prefix):
            return i
    return None


def _parse_append_target_program_index(value: Any) -> Optional[int]:
    """SMApply sends this dropdown's answer as a raw 0-based integer
    choice index (0 = "Education Document 1", 1 = "Education Document
    2", ...), confirmed against a real submission. None if blank/absent
    or not actually int-like (defensive only -- not expected in
    practice)."""
    if value is None or value == "":
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def parse_appended_domain_rows(reviews: List[Dict[str, Any]]) -> List[Tuple[int, List[Dict[str, str]]]]:
    """
    Parse every "append to existing domain breakdown" block out of the
    Education Document Review task's responses -- up to
    _APPENDED_ROW_MAX_BLOCKS (5) repeatable blocks, each independently
    targeting any program.

    Returns a list of (program_index, rows) tuples, one per block the
    evaluator actually filled in (stops at the first "No"/missing gate,
    same semantics as the manual-entry flow's "Add another...?"
    pattern). `rows` is that block's list of {"subject", "theory_hours",
    "sim_lab_hours", "clinical_hours"} dicts for whichever of its 10
    slots were filled in (subject non-blank). Empty list if the
    evaluator answered No on the very first gate, or nothing else
    parses. Does NOT dedupe against any program's existing subjects --
    that happens in apply_appended_domain_rows once every block's target
    program is known.
    """
    ed_rv = _reviews_by_task(reviews).get(TASK_NAME_EDUCATION)
    if not ed_rv:
        return []

    responses = ed_rv.get("responses") or []

    # Row-grid groups, in encounter order -- one group per block, each
    # keyed by that block's own randomly-generated field-id prefix.
    # Block i's rows are matched to this list PURELY by position (the
    # i-th "Rows to add..." grid encountered pairs with the i-th append
    # block), since both appear exactly once per block, in the same
    # relative order.
    row_group_order, row_groups = _group_by_field_prefix(responses, _APPENDED_ROW_LABEL_PREFIX)

    first_gate_idx = next(
        (i for i, r in enumerate(responses)
         if _nl(r.get("label")).startswith(_APPEND_TO_EXISTING_LABEL_PREFIX)),
        None,
    )
    if first_gate_idx is None:
        return []

    results: List[Tuple[int, List[Dict[str, str]]]] = []
    gate_idx = first_gate_idx

    for block_num in range(_APPENDED_ROW_MAX_BLOCKS):
        gate_flag = _map_yesno(responses[gate_idx].get("response"), {0: "Yes", 1: "No"})
        if gate_flag != "Yes":
            break

        target_idx = _find_response_index_after(responses, gate_idx, _APPEND_TARGET_DOC_LABEL_PREFIX)
        if target_idx is None:
            break
        program_index = _parse_append_target_program_index(responses[target_idx].get("response"))

        if program_index is not None and block_num < len(row_group_order):
            values = row_groups[row_group_order[block_num]]
            rows: List[Dict[str, str]] = []
            for row_i in range(1, _APPENDED_ROW_COUNT + 1):
                base = (row_i - 1) * 4  # 4 fields per row: subject, theory, sim/lab, clinical
                subject = values.get(base, "")
                if not subject:
                    continue
                rows.append({
                    "subject":        subject,
                    "theory_hours":   values.get(base + 1, ""),
                    "sim_lab_hours":  values.get(base + 2, ""),
                    "clinical_hours": values.get(base + 3, ""),
                })
            if rows:
                results.append((program_index, rows))

        # Blocks 2-5 are gated by a DIFFERENT label than block 1
        # ("Append to another...?" vs "Would you like to append...?").
        # If this isn't found, there's no next block -- stop rather than
        # wandering into the unrelated manual-entry flow further down
        # this same task's responses (which uses its own, different
        # gate label and would never match this search anyway, but the
        # explicit break keeps the loop's intent clear).
        next_gate_idx = _find_response_index_after(responses, target_idx, _APPEND_ANOTHER_LABEL_PREFIX)
        if next_gate_idx is None:
            break
        gate_idx = next_gate_idx

    return results


def apply_appended_domain_rows(education_data: Dict[str, Any],
                               append_results: List[Tuple[int, List[Dict[str, str]]]]) -> Dict[str, Any]:
    """
    Apply every (program_index, rows) block from parse_appended_domain_rows
    to education_data["programs"], inserting each block's rows into its
    target program's "b1" list right before that program's TOTALS row.
    Skips any (program_index, rows) pair whose program_index is out of
    range.

    Dedupes each block's rows against subjects the target program's B1
    already has (case/whitespace-insensitive via _nl) -- an evaluator-
    entered row for a subject that's already present is dropped rather
    than appended as a duplicate. This applies cumulatively across
    blocks within the same call too: if two different blocks in one
    submission both target the same program with an overlapping
    subject, the later block's row for that subject is dropped, since a
    duplicate subject row would look like a rendering bug to whoever
    reads the final PDF, not like a deliberate addition.

    Does NOT recompute any TOTALS row itself -- callers should follow up
    with recompute_domain_totals() on every affected program's "b1" list
    once the merge is done, or the printed TOTALS will silently
    undercount by however many hours the newly appended rows carry.
    """
    programs = education_data.get("programs") or []

    for program_index, new_rows in append_results:
        if program_index is None or not new_rows or program_index >= len(programs):
            continue

        b1 = programs[program_index].setdefault("b1", [])
        existing_subjects = {_nl(r.get("subject")) for r in b1}
        rows_to_add = [r for r in new_rows if _nl(r["subject"]) not in existing_subjects]
        if not rows_to_add:
            continue

        totals_pos = next(
            (i for i, r in enumerate(b1) if r.get("subject", "").strip().upper() == "TOTALS"),
            len(b1),
        )
        b1[totals_pos:totals_pos] = rows_to_add

    return education_data


# =========================================================
# OMITTED EDUCATION DOCUMENTS (evaluator can exclude specific prior-
# report education programs from this report)
# ---------------------------------------------------------
# Gated the same way as the append-rows feature -- only offered when
# "Does this applicant have a previously issued advisory report?" =
# Yes -- and positioned right between that flag and the append flow's
# own gate in the real form (confirmed 2026-08-11):
#   "Does this applicant have a previously issued advisory report?"
#   -> "Would you like to omit any education documents from this
#      report?" (Yes/No)
#   -> "Which education document(s) should be omitted?" (multi-select)
#   -> [append-rows flow starts here]
#
# The multi-select answer comes back as a LIST of raw 0-based indices
# straight from SMApply -- [1] for "Education Document 2" -- same
# 0-based convention as the append feature's single-select target
# dropdown, just plural. Confirmed against a real submission on
# 2026-08-11.
#
# !!! ORDERING REQUIREMENT !!!
# apply_omitted_education_documents() MUST run AFTER
# apply_appended_domain_rows() (and its totals recompute), never
# before. Both features address programs by "Education Document N" =
# index N-1 into the ORIGINAL prior-report programs list -- append
# resolves its own target indices against that original numbering, so
# removing an entry first would shift every later index out from under
# it (e.g. omitting Document 2 first would make what was Document 4
# become index 2, silently redirecting an append meant for Document 4
# onto the wrong program). Removing entries here doesn't need to
# preserve indices for anything downstream, so it's safe to always run
# this step dead last, right before the education data is handed to the
# renderers.
# -----------------------------------------------------------------
_OMIT_QUESTION_LABEL_PREFIX = "would you like to omit any education documents from"
_OMIT_TARGET_LABEL_PREFIX   = "which education document(s) should be omitted"


def parse_omitted_education_document_indices(reviews: List[Dict[str, Any]]) -> List[int]:
    """
    Parse the evaluator's "omit education documents" answer out of the
    Education Document Review task's responses. Returns a sorted,
    de-duplicated list of 0-based indices into education_data["programs"]
    (original prior-report order) to exclude from the report. Empty list
    if the evaluator answered No, left it blank, the gate/target
    questions weren't found, or the target's response wasn't a list.
    """
    ed_rv = _reviews_by_task(reviews).get(TASK_NAME_EDUCATION)
    if not ed_rv:
        return []

    responses = ed_rv.get("responses") or []

    gate_idx = next(
        (i for i, r in enumerate(responses)
         if _nl(r.get("label")).startswith(_OMIT_QUESTION_LABEL_PREFIX)),
        None,
    )
    if gate_idx is None:
        return []

    gate_flag = _map_yesno(responses[gate_idx].get("response"), {0: "Yes", 1: "No"})
    if gate_flag != "Yes":
        return []

    target_idx = _find_response_index_after(responses, gate_idx, _OMIT_TARGET_LABEL_PREFIX)
    if target_idx is None:
        return []

    raw = responses[target_idx].get("response")
    if not isinstance(raw, list):
        return []

    indices: List[int] = []
    for v in raw:
        try:
            indices.append(int(v))
        except (TypeError, ValueError):
            continue

    return sorted(set(indices))


def apply_omitted_education_documents(education_data: Dict[str, Any],
                                      omit_indices: List[int]) -> Dict[str, Any]:
    """
    Remove the given 0-based indices from education_data["programs"].
    No-op if omit_indices is empty. See the ORDERING REQUIREMENT note
    above -- this must be the LAST education_data mutation before the
    data is handed off to rendering.
    """
    if not omit_indices:
        return education_data

    omit_set = set(omit_indices)
    programs = education_data.get("programs") or []
    education_data["programs"] = [p for i, p in enumerate(programs) if i not in omit_set]
    return education_data


def _collapse_ws(s: Any) -> str:
    """Collapse any run of whitespace (including an embedded literal
    '\\n' from a wrapped PDF table cell) into a single space. Same
    normalization a2_identifying_line() applies to a school_name before
    display -- used here so a school name pulled from education_data
    lines up with the same name's plain-text appearance in a Part C
    bullet line."""
    return re.sub(r"\s+", " ", str(s or "")).strip()


def filter_out_omitted_school_bullets(bullets: List[str],
                                      omitted_school_names: List[str]) -> List[str]:
    """
    Drop any bullet whose text starts with one of `omitted_school_names`
    (case-insensitive, whitespace-collapsed comparison on both sides).

    Built for "C2: Nursing Education Documents" specifically: that
    section is ported verbatim from the prior report's own Part C text
    (see apply_part_c_education_documents) rather than built from
    education_data["programs"], so removing a program via the "omit
    education documents" feature does NOT automatically remove its
    documents from this section -- confirmed against a real prior
    report, whose C2 bullets are consistently formatted as
    "<school name> - <country> - <document type> - <date>", i.e. the
    school name is always the bullet's own leading text. A prefix match
    (not a substring match) is required: one program's full name can
    legitimately contain another program's shorter name as a substring
    (e.g. "Genesis Institute of Nursing (FKA Anil Baghi College/School
    of Nursing)" contains "Anil Baghi College of Nursing"), and only a
    same-position prefix reliably distinguishes "this bullet belongs to
    the short-named program" from "this bullet belongs to the long-named
    program that happens to mention the short name later in its own
    title" -- a substring match would wrongly drop both.

    No-op (returns bullets unchanged) if omitted_school_names is empty.
    """
    if not omitted_school_names:
        return bullets

    normalized_omitted = [_collapse_ws(n).lower() for n in omitted_school_names if _collapse_ws(n)]
    if not normalized_omitted:
        return bullets

    kept: List[str] = []
    for b in bullets:
        b_norm = _collapse_ws(b).lower()
        if any(b_norm.startswith(name) for name in normalized_omitted):
            continue
        kept.append(b)
    return kept


def _parse_domain_hours(value: Any) -> int:
    """
    Parse one B1 hours cell into an int for TOTALS summation.

    This document uses several non-numeric placeholders in these cells
    -- "INTEG" (hours integrated/embedded into other coursework, not
    counted separately) and "---" (not applicable / blank) both appear
    throughout real reports. Confirmed against an actual prior report's
    own printed TOTALS row (Nursing Care of the Adult - Medical: theory
    105, sim/lab INTEG, clinical 600; ... TOTALS: theory 751, sim/lab 0,
    clinical 2894) that both placeholders -- and blank cells -- are
    meant to contribute exactly 0 to the sum, not to invalidate it.
    Thousands-separator commas are stripped before parsing. Any other
    unparseable text also falls back to 0 rather than raising, since one
    malformed cell shouldn't block a total for every other row.
    """
    s = safe_str(value).replace(",", "").strip()
    if not s:
        return 0
    try:
        return int(float(s))
    except (TypeError, ValueError):
        return 0


def recompute_domain_totals(b1_rows: List[Dict[str, str]]) -> List[Dict[str, str]]:
    """
    Recompute a program's TOTALS row (theory/sim-lab/clinical hours) as
    the sum of every other row in `b1_rows`, mutating and returning
    `b1_rows` in place. No-op if there's no TOTALS row.

    Meant to be called after any edit to a program's b1 list that could
    change its hour totals -- currently that's apply_appended_domain_rows,
    but this is intentionally generic rather than baked into that
    function, so future edits to b1 (e.g. a manual row-edit feature) can
    reuse it the same way instead of duplicating the summation logic.
    """
    totals_row = next(
        (r for r in b1_rows if str(r.get("subject", "")).strip().upper() == "TOTALS"),
        None,
    )
    if totals_row is None:
        return b1_rows

    other_rows = [r for r in b1_rows if r is not totals_row]
    totals_row["theory_hours"]   = str(sum(_parse_domain_hours(r.get("theory_hours"))   for r in other_rows))
    totals_row["sim_lab_hours"]  = str(sum(_parse_domain_hours(r.get("sim_lab_hours"))  for r in other_rows))
    totals_row["clinical_hours"] = str(sum(_parse_domain_hours(r.get("clinical_hours")) for r in other_rows))
    return b1_rows


def normalize_education_programs(education_data: Optional[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Normalize education_data into a flat list of {"a2": {...}, "b1": [...]}
    program dicts, regardless of source. Supports the current
    {"programs": [...]} shape (from extract_nnas_education or
    parse_manual_education_from_reviews) and, defensively, a legacy
    single-pair {"a2": {...}, "b1": [...]} shape.
    """
    if not education_data:
        return []
    if "programs" in education_data:
        return [p for p in (education_data.get("programs") or [])
                if p.get("a2") or p.get("b1")]
    a2 = education_data.get("a2") or {}
    b1 = education_data.get("b1") or []
    return [{"a2": a2, "b1": b1}] if (a2 or b1) else []

# =========================================================
# PDF - PAGE 1: NOTES + METADATA
# =========================================================

def draw_wrapped_text(c, text, *, x, y_top, w, font="Helvetica", size=12, leading=16):
    c.setFont(font, size)
    lines = wrap_lines(text, font, size, w)
    y = y_top
    for line in lines:
        c.drawString(x, y, line)
        y -= leading

def draw_notes_section(c, *, x, y_top, w, header_text, body_text,
                       header_height=34, body_height=120):
    total_height = header_height + body_height
    y_bottom     = y_top - total_height
    c.setStrokeColor(colors.black); c.setLineWidth(2)
    c.rect(x, y_bottom, w, total_height, fill=0, stroke=1)
    c.setFillColor(COLOR_HEADER_BG)
    c.rect(x, y_top - header_height, w, header_height, fill=1, stroke=0)
    c.setFillColor(COLOR_RED); c.setFont("Helvetica-Bold", 18)
    c.drawCentredString(x + w / 2, y_top - header_height + 10, header_text)
    c.setStrokeColor(colors.black); c.setLineWidth(1)
    c.line(x, y_top - header_height, x + w, y_top - header_height)
    c.setFillColor(colors.black)
    draw_wrapped_text(c, body_text, x=x + 10, y_top=y_top - header_height - 20,
                      w=w - 20, font="Helvetica", size=12, leading=16)

def render_assessment_notes_page1(c, payload, page_h, app_info: Dict[str, str]):
    notes             = payload.get("assessment_notes") or {}
    base_public       = safe_str(notes.get("public"))
    identity_note     = safe_str(notes.get("identity_section_note"))
    education_note    = safe_str(notes.get("education_section_note"))
    registration_note = safe_str(notes.get("registration_section_note"))

    section_lines = []
    if identity_note:
        section_lines.append(f"Identity Section Note: {identity_note}")
    if education_note:
        section_lines.append(f"Education Section Note: {education_note}")
    if registration_note:
        section_lines.append(f"Registration Section Note: {registration_note}")

    appended        = "\n".join(section_lines)
    public_text     = "\n\n".join(p for p in [base_public, appended] if p)
    assessment_text = safe_str(notes.get("assessment"))

    nursing_group      = app_info.get("nursing_group", "")
    province           = app_info.get("province", "")
    nnas_id            = app_info.get("nnas_id", "")
    application_number = app_info.get("application_number", "")

    public_y_top     = (page_h * 0.65) - 108 - 72 + HALF_INCH
    assessment_y_top = public_y_top - (34 + 60) - GAP_BETWEEN_NOTE_BOXES

    META_Y       = public_y_top + 160
    META_LEFT    = 30
    META_MID     = 310
    META_LEADING = 22
    LABEL_FONT   = "Helvetica-Bold"
    VALUE_FONT   = "Helvetica"
    LABEL_SIZE   = 12
    VALUE_SIZE   = 12

    def draw_meta_pair(label, value, x, y):
        c.setFont(LABEL_FONT, LABEL_SIZE)
        c.setFillColor(COLOR_RED)
        c.drawString(x, y, label)
        lw = c.stringWidth(label, LABEL_FONT, LABEL_SIZE)
        c.setFont(VALUE_FONT, VALUE_SIZE)
        c.setFillColor(COLOR_TEXT)
        c.drawString(x + lw + 4, y, value)

    draw_meta_pair("Nursing Group:",      nursing_group,      META_LEFT, META_Y)
    draw_meta_pair("Province:",           province,           META_MID,  META_Y)
    draw_meta_pair("NNAS ID Number:",     nnas_id,            META_LEFT, META_Y - META_LEADING)
    draw_meta_pair("Application Number:", application_number, META_MID,  META_Y - META_LEADING)

    if not public_text and not assessment_text:
        return

    if public_text:
        pub_lines  = wrap_lines(public_text, "Helvetica", 12, NOTES_SECTION_W - 20)
        pub_body_h = max(60, len(pub_lines) * 16 + 20)
        assessment_y_top = public_y_top - (34 + pub_body_h) - GAP_BETWEEN_NOTE_BOXES
        draw_notes_section(c, x=NOTES_SECTION_X, y_top=public_y_top, w=NOTES_SECTION_W,
                           header_text="Public Assessment Notes", body_text=public_text,
                           header_height=34, body_height=pub_body_h)
    if assessment_text:
        asmt_lines  = wrap_lines(assessment_text, "Helvetica", 12, NOTES_SECTION_W - 20)
        asmt_body_h = max(140, len(asmt_lines) * 16 + 20)
        draw_notes_section(c, x=NOTES_SECTION_X, y_top=assessment_y_top, w=NOTES_SECTION_W,
                           header_text="Assessment Notes", body_text=assessment_text,
                           header_height=34, body_height=asmt_body_h)

# =========================================================
# PDF - PAGE 2: PERSONAL INFO (A1)
# =========================================================

def render_personal_info(c, payload, *, y_top):
    styles      = getSampleStyleSheet()
    label_style = ParagraphStyle("label", parent=styles["Normal"],
                                 fontName="Helvetica-Bold", fontSize=11,
                                 textColor=COLOR_RED, leading=14)
    value_style = ParagraphStyle("value", parent=styles["Normal"],
                                 fontName="Helvetica", fontSize=10,
                                 textColor=COLOR_TEXT, leading=14)
    rows       = personal_info_rows(payload)
    table_data = [
        [Paragraph(label, label_style), Paragraph(value, value_style)]
        for label, value in rows
    ]
    table = Table(table_data, colWidths=[270, 260])
    table.setStyle(TableStyle([
        ("BOX",          (0, 0), (-1, -1), OUTER_BORDER_WIDTH, colors.black),
        ("INNERGRID",    (0, 0), (-1, -1), INNER_BORDER_WIDTH, colors.black),
        ("BACKGROUND",   (0, 0), (-1,  0), COLOR_HEADER_BG),
        ("BACKGROUND",   (0, 2), (-1,  2), COLOR_HEADER_BG),
        ("BACKGROUND",   (0, 4), (-1,  4), COLOR_HEADER_BG),
        ("BACKGROUND",   (0, 6), (-1,  6), COLOR_HEADER_BG),
        ("LEFTPADDING",  (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING",   (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 6),
        ("VALIGN",       (0, 0), (-1, -1), "TOP"),
    ]))
    _w, table_h = table.wrap(TABLE_WIDTH, 700)
    table.drawOn(c, LEFT_MARGIN, y_top - table_h)

# =========================================================
# PDF - PAGES 3+: REGISTRATION (A3)
# =========================================================

def render_registration_tables(writer, registrations, *, y_top, start_index: int = PAGE_A3) -> int:
    """
    Render each registration entry on its own page, starting at
    `start_index` (PAGE_A3 when the Full template's A2 education overview
    page precedes registration; PAGE_A2 when the Base template is used
    and there's no education program at all, so registration is the
    first content section after personal info).

    Dynamic pagination: the FIRST entry reuses the template's existing
    slot at start_index (inserting a blank page only if the template is
    too short to contain it at all). EVERY subsequent entry gets a
    brand-new page inserted at that exact position via insert_blank_page,
    which shifts every later page (education B1 slot, evaluator, Part C,
    or any already-inserted education pages) down automatically. This
    means the report never silently overwrites trailing pages, regardless
    of how many pre-built filler pages the template happens to ship with,
    or how many registration entries a given application has.

    Returns the page index immediately after the last registration page
    -- i.e., where subsequent inserted sections (extra education program
    pages beyond the first) should be inserted next.
    """
    styles      = getSampleStyleSheet()
    label_style = ParagraphStyle("label", parent=styles["Normal"],
                                 fontName="Helvetica-Bold", fontSize=FONT_LABEL,
                                 textColor=COLOR_RED, leading=12)
    value_style = ParagraphStyle("value", parent=styles["Normal"],
                                 fontName="Helvetica", fontSize=FONT_BODY,
                                 textColor=COLOR_TEXT, leading=12)
    page_index = start_index
    for i, reg in enumerate(registrations):
        if i == 0:
            # First entry reuses the template's existing A3 slot; only
            # insert if the template is shorter than expected.
            if page_index >= len(writer.pages):
                writer.insert_blank_page(width=612, height=792, index=page_index)
        else:
            # Every subsequent entry ALWAYS gets a brand-new page, pushing
            # everything after it (including education/evaluator/Part C)
            # down instead of overwriting it.
            writer.insert_blank_page(width=612, height=792, index=page_index)

        page = writer.pages[page_index]
        w, h = float(page.mediabox.width), float(page.mediabox.height)
        rows       = registration_rows(reg)
        table_data = [
            [Paragraph(k, label_style), Paragraph(v, value_style)]
            for k, v in rows
        ]
        table = styled_table(table_data, [260, 270], header_rows=0)
        def draw(c, _t=table):
            _w, table_h = _t.wrap(TABLE_WIDTH, 700)
            _t.drawOn(c, LEFT_MARGIN, y_top - table_h)
        overlay = build_overlay(w, h, draw)
        page.merge_page(overlay.pages[0])
        page_index += 1

    return page_index

# =========================================================
# PDF - PART B: A2 EDUCATION OVERVIEW + B1 DOMAIN BREAKDOWN
# =========================================================

_A2_DISPLAY_ROWS = [
    ("Name of school / educational institution:",                    "school_name"),
    ("Current address of school / educational institution:",         "school_address"),
    ("Type of school / educational institution:",                    "school_type"),
    ("Primary language of educational institution:",                 "primary_language"),
    ("Applicant name used while attending this school: (if different from current name):", "applicant_name_at_school"),
    ("Category of program:",                                         "program_category"),
    ("Minimum entrance requirements for admission to this program:", "entrance_requirements"),
    ("Length of study for this program:",                            "program_length"),
    ("Program start date:",                                          "program_start"),
    ("Program completion / graduation date:",                        "program_end"),
    ("Language of instruction - theory:",                            "instruction_language_theory"),
    ("Language of instruction - clinical:",                          "instruction_language_clinical"),
    ("Name of credential / degree obtained:",                        "credential"),
    ("Comments:",                                                    "comments"),
]

_B1_COL_HEADERS = [
    "SUBJECT AREA",
    "THEORY\nHOURS",
    "SIMULATION/\nLAB\nHOURS",
    "CLINICAL\nHOURS\n(Direct\nPatient Care)",
]
_B1_COL_WIDTHS = [230, 80, 90, 130]

def _b1_display(val: str) -> str:
    return val if val else "---"

def a2_identifying_line(a2: Dict[str, str]) -> Tuple[str, str]:
    """
    Pull the (school name, date range) identifying pair out of an A2
    dict, for display directly above the matching B1 domain-breakdown
    table -- B1 alone doesn't say which school/program it belongs to, so
    this gives that context without repeating the full A2 table.

    school_name is whitespace-collapsed HERE (not at extraction) because
    this is the one consumer that's fragile to an embedded "\n": its
    caller (render_b1_on_page) draws it with plain drawString(), which
    has no glyph for U+000A and renders a solid black square in its
    place for any wrapped-cell school name that came through with a
    literal newline. Collapsing at the source in
    _extract_a2_raw_rows() instead of here previously broke the
    unrelated "Language of instruction" theory/clinical split, which
    depends on a literal "\n" surviving in ITS raw cell value -- see the
    comment there for details.
    """
    school_name = re.sub(r"\s+", " ", a2.get("school_name", "") or "").strip()
    start       = a2.get("program_start", "")
    end         = a2.get("program_end", "")
    date_range  = f"{start}- {end}" if (start or end) else ""
    return school_name, date_range

def render_a2_on_page(writer: PdfWriter, a2: Dict[str, str],
                      page_index: int, *, y_top: float = 680):
    while len(writer.pages) <= page_index:
        writer.add_blank_page(width=612, height=792)
    page   = writer.pages[page_index]
    w, h   = float(page.mediabox.width), float(page.mediabox.height)
    styles = getSampleStyleSheet()
    label_style = ParagraphStyle("a2l", parent=styles["Normal"],
                                 fontName="Helvetica-Bold", fontSize=10,
                                 textColor=COLOR_RED, leading=13)
    value_style = ParagraphStyle("a2v", parent=styles["Normal"],
                                 fontName="Helvetica", fontSize=10,
                                 textColor=COLOR_TEXT, leading=13)
    table_data = [
        [Paragraph(label, label_style), Paragraph(a2.get(key, ""), value_style)]
        for label, key in _A2_DISPLAY_ROWS
    ]
    grey_rows  = {0, 2, 4, 6, 8, 10, 12}
    style_cmds = [
        ("BOX",          (0, 0), (-1, -1), OUTER_BORDER_WIDTH, colors.black),
        ("INNERGRID",    (0, 0), (-1, -1), INNER_BORDER_WIDTH, colors.black),
        ("LEFTPADDING",  (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING",   (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 6),
        ("VALIGN",       (0, 0), (-1, -1), "TOP"),
    ]
    for r in grey_rows:
        style_cmds.append(("BACKGROUND", (0, r), (-1, r), COLOR_HEADER_BG))
    tbl = Table(table_data, colWidths=[270, 260])
    tbl.setStyle(TableStyle(style_cmds))

    def draw(c):
        _w, tbl_h = tbl.wrap(TABLE_WIDTH, h)
        tbl.drawOn(c, LEFT_MARGIN, y_top - tbl_h)

    overlay = build_overlay(w, h, draw)
    page.merge_page(overlay.pages[0])

def render_b1_on_page(writer: PdfWriter, b1_rows: List[Dict[str, str]],
                      page_index: int, *, y_top: float = 620,
                      header_a2: Optional[Dict[str, str]] = None):
    while len(writer.pages) <= page_index:
        writer.add_blank_page(width=612, height=792)
    page   = writer.pages[page_index]
    w, h   = float(page.mediabox.width), float(page.mediabox.height)
    styles = getSampleStyleSheet()
    hdr_style  = ParagraphStyle("b1h", parent=styles["Normal"],
                                fontName="Helvetica-Bold", fontSize=9,
                                textColor=COLOR_RED, leading=11, alignment=1)
    cell_style = ParagraphStyle("b1c", parent=styles["Normal"],
                                fontName="Helvetica", fontSize=9,
                                textColor=COLOR_TEXT, leading=11)
    num_style  = ParagraphStyle("b1n", parent=styles["Normal"],
                                fontName="Helvetica", fontSize=9,
                                textColor=COLOR_TEXT, leading=11, alignment=1)

    header_row = [Paragraph(h_txt, hdr_style) for h_txt in _B1_COL_HEADERS]

    def make_row(rd):
        return [
            Paragraph(_b1_display(rd["subject"]),        cell_style),
            Paragraph(_b1_display(rd["theory_hours"]),   num_style),
            Paragraph(_b1_display(rd["sim_lab_hours"]),  num_style),
            Paragraph(_b1_display(rd["clinical_hours"]), num_style),
        ]

    all_data   = [header_row] + [make_row(r) for r in b1_rows]
    style_cmds = [
        ("BOX",          (0, 0), (-1, -1), OUTER_BORDER_WIDTH, colors.black),
        ("INNERGRID",    (0, 0), (-1, -1), INNER_BORDER_WIDTH, colors.black),
        ("LEFTPADDING",  (0, 0), (-1, -1), CELL_PADDING),
        ("RIGHTPADDING", (0, 0), (-1, -1), CELL_PADDING),
        ("TOPPADDING",   (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 4),
        ("VALIGN",       (0, 0), (-1, -1), "TOP"),
        ("BACKGROUND",   (0, 0), (-1,  0), COLOR_HEADER_BG),
        ("FONTNAME",     (0, 0), (-1,  0), "Helvetica-Bold"),
        ("TEXTCOLOR",    (0, 0), (-1,  0), COLOR_RED),
    ]
    for row_i in range(1, len(all_data)):
        if row_i % 2 == 0:
            style_cmds.append(("BACKGROUND", (0, row_i), (-1, row_i), COLOR_ALT_ROW))
    tbl = Table(all_data, colWidths=_B1_COL_WIDTHS)
    tbl.setStyle(TableStyle(style_cmds))

    def draw(c):
        if header_a2:
            school_name, date_range = a2_identifying_line(header_a2)
            if school_name:
                c.setFont("Helvetica-Bold", 11)
                c.setFillColor(COLOR_TEXT)
                c.drawString(LEFT_MARGIN, y_top + 28, school_name)
            if date_range:
                c.setFont("Helvetica", 10)
                c.setFillColor(COLOR_TEXT)
                c.drawString(LEFT_MARGIN, y_top + 14, date_range)
        _w, tbl_h = tbl.wrap(TABLE_WIDTH, h)
        tbl.drawOn(c, LEFT_MARGIN, y_top - tbl_h)

    overlay = build_overlay(w, h, draw)
    page.merge_page(overlay.pages[0])

# =========================================================
# PDF - SECOND-LAST PAGE: EVALUATOR FIELDS
# =========================================================

def render_evaluator_fields(c, evaluator_name: str, date_str: str):
    c.setFont("Helvetica", 10)
    c.setFillColor(COLOR_TEXT)
    c.drawString(EVALUATOR_VALUE_X, EVALUATOR_Y, evaluator_name)
    c.drawString(DATE_VALUE_X,      EVALUATOR_Y, date_str)

# =========================================================
# PDF - LAST PAGE: PART C
# =========================================================

def render_part_c_last_page(writer, payload):
    sections = build_part_c_sections(payload)
    if not sections: return
    last_index = len(writer.pages) - 1
    page       = writer.pages[last_index]
    w, h       = float(page.mediabox.width), float(page.mediabox.height)

    def draw(c):
        y = PART_C_Y_TOP
        for sec in sections:
            header  = safe_str(sec.get("header"))
            bullets = sec.get("bullets") or []
            if not isinstance(bullets, list): bullets = [safe_str(bullets)]
            if header:
                c.setFillColor(COLOR_RED); c.setFont("Helvetica-Bold", 16)
                c.drawString(40, y, header); y -= 22
            c.setFillColor(COLOR_TEXT); c.setFont("Helvetica", 11)
            for b in bullets:
                b = safe_str(b)
                if not b: continue
                lines = wrap_lines(f"\u25a0 {b}", "Helvetica", 11, 520)
                for line in lines:
                    c.drawString(55, y, line); y -= 14
                y -= 2
            y -= 10

    overlay = build_overlay(w, h, draw)
    page.merge_page(overlay.pages[0])

# =========================================================
# PDF MAIN BUILD
# =========================================================

def populate_pdf(template_path, payload, app_info: Dict[str, str],
                 evaluator_name: str = "", date_str: str = "",
                 education_data: Optional[Dict[str, Any]] = None) -> PdfWriter:
    """
    Build the final PDF by overlaying content onto the template.

    Page layout (full template, 7 base pages):
        Index 0             — Notes + metadata
        Index 1             — A1 Personal Information
        Index 2..2+N-1      — A2 Nursing Education Overview, one page per
                               education program, all grouped consecutively
                               (N = number of programs; template's fixed
                               slot at index 2 holds the first one)
        Index 2+N           — A3 Registration (first entry)
        ...                 — A3 additional entries (dynamic pagination)
        (3rd-from-last)..   — B1 Domain Breakdown, one page per education
                               program, all grouped consecutively (each
                               page labeled with its matching school name
                               and program dates); template's fixed slot
                               holds the first one
        2nd-from-last       — Evaluator
        Last                — Part C

    All A2 pages are rendered together before registration, and all B1
    pages are rendered together near the end, rather than interleaving
    one A2+B1 pair per program -- so Part A (overviews) and Part B
    (domain breakdowns) each read as one contiguous section.
    """
    reader = PdfReader(template_path)
    writer = PdfWriter()
    for p in reader.pages:
        writer.add_page(p)

    if PAGE_NOTES < len(writer.pages):
        page0  = writer.pages[PAGE_NOTES]
        w0, h0 = float(page0.mediabox.width), float(page0.mediabox.height)
        overlay0 = build_overlay(
            w0, h0,
            lambda c: render_assessment_notes_page1(c, payload, h0, app_info)
        )
        page0.merge_page(overlay0.pages[0])

    if PAGE_A1 >= len(writer.pages):
        raise RuntimeError(f"Template missing page index PAGE_A1={PAGE_A1}")
    page1    = writer.pages[PAGE_A1]
    w1, h1   = float(page1.mediabox.width), float(page1.mediabox.height)
    overlay1 = build_overlay(
        w1, h1,
        lambda c: render_personal_info(c, payload,
                                       y_top=PERSONAL_INFO_HEADER_Y - HALF_INCH)
    )
    page1.merge_page(overlay1.pages[0])

    programs = normalize_education_programs(education_data)

    # ---- All A2 overview pages, grouped consecutively ----
    if programs:
        render_a2_on_page(writer, programs[0].get("a2") or {}, PAGE_A2)
        a2_insert_index = PAGE_A2 + 1
        for prog in programs[1:]:
            writer.insert_blank_page(width=612, height=792, index=a2_insert_index)
            render_a2_on_page(writer, prog.get("a2") or {}, a2_insert_index, y_top=680)
            a2_insert_index += 1

    # Registration starts right after the LAST A2 page (PAGE_A2 + number
    # of programs) when programs exist, or right after personal info
    # (PAGE_A2) when there's no education program at all and the Base
    # template (with no A2 page) was downloaded instead.
    registration_start_index = PAGE_A2 + len(programs) if programs else PAGE_A2

    regs = extract_registration_entries(payload)
    registration_end_index = registration_start_index
    if regs:
        registration_end_index = render_registration_tables(
            writer, regs, y_top=REGISTRATION_HEADER_Y - HALF_INCH,
            start_index=registration_start_index,
        )

    # ---- All B1 domain-breakdown pages, grouped consecutively ----
    # The template's fixed B1 slot sits 3 pages from the end at this
    # point (nothing has touched the trailing [B1, evaluator, Part C]
    # pages yet), so that offset is still valid regardless of how many
    # A2/registration pages were inserted above.
    total_pages_before_b1 = len(writer.pages)
    b1_first_index         = total_pages_before_b1 - 3

    if programs and b1_first_index >= 0:
        render_b1_on_page(writer, programs[0].get("b1") or [], b1_first_index,
                          header_a2=programs[0].get("a2") or {})
        b1_insert_index = b1_first_index + 1
        for prog in programs[1:]:
            writer.insert_blank_page(width=612, height=792, index=b1_insert_index)
            render_b1_on_page(writer, prog.get("b1") or [], b1_insert_index, y_top=700,
                              header_a2=prog.get("a2") or {})
            b1_insert_index += 1

    total_pages          = len(writer.pages)
    evaluator_page_index = total_pages - 2
    part_c_page_index    = total_pages - 1

    if evaluator_page_index >= 0 and (evaluator_name or date_str):
        page_ev    = writer.pages[evaluator_page_index]
        w_ev, h_ev = float(page_ev.mediabox.width), float(page_ev.mediabox.height)
        overlay_ev = build_overlay(
            w_ev, h_ev,
            lambda c: render_evaluator_fields(c, evaluator_name, date_str)
        )
        page_ev.merge_page(overlay_ev.pages[0])

    render_part_c_last_page(writer, payload)

    return writer

# =========================================================
# DOCX LOW-LEVEL HELPERS
# =========================================================

def _hex(rgb: RGBColor) -> str:
    return f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"

def _set_cell_bg(cell, rgb: RGBColor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  _hex(rgb))
    tcPr.append(shd)

def _set_cell_borders(cell, size: int = 4, color: str = "000000"):
    tcPr = cell._tc.get_or_add_tcPr()
    for existing in tcPr.findall(qn("w:tcBorders")):
        tcPr.remove(existing)
    tcBords = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement(f"w:{side}")
        if size > 0:
            b.set(qn("w:val"),   "single")
            b.set(qn("w:sz"),    str(size))
            b.set(qn("w:space"), "0")
            b.set(qn("w:color"), color)
        else:
            b.set(qn("w:val"),   "none")
            b.set(qn("w:sz"),    "0")
            b.set(qn("w:space"), "0")
            b.set(qn("w:color"), "auto")
        tcBords.append(b)
    tcPr.append(tcBords)

def _cell_margins(cell, top=60, bottom=60, left=100, right=100):
    tcPr  = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for side, val in (("top", top), ("bottom", bottom),
                      ("left", left), ("right", right)):
        m = OxmlElement(f"w:{side}")
        m.set(qn("w:w"),    str(val))
        m.set(qn("w:type"), "dxa")
        tcMar.append(m)
    tcPr.append(tcMar)

def _write_cell(cell, text: str, bold: bool = False,
                color: RGBColor = DOCX_TEXT, size: int = 10,
                align=WD_ALIGN_PARAGRAPH.LEFT):
    cell.text = ""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    _cell_margins(cell)
    lines = str(text).split("\n") if text else [""]
    for idx, line in enumerate(lines):
        p = cell.paragraphs[0] if idx == 0 else cell.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)
        p.alignment = align
        run = p.add_run(line)
        run.bold           = bold
        run.font.size      = Pt(size)
        run.font.color.rgb = color
        run.font.name      = "Calibri"

def _set_table_outer_border(table, size: int = 8, color: str = "000000"):
    tbl   = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    for existing in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(existing)
    tblBdr = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"),   "single")
        b.set(qn("w:sz"),    str(size))
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), color)
        tblBdr.append(b)
    tblPr.append(tblBdr)

def _force_cell_borders(table, size: int = 8, color: str = "000000"):
    for row in table.rows:
        for cell in row.cells:
            tc   = cell._tc
            tcPr = tc.get_or_add_tcPr()
            for existing in tcPr.findall(qn("w:tcBorders")):
                tcPr.remove(existing)
            tcBords = OxmlElement("w:tcBorders")
            for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
                b = OxmlElement(f"w:{side}")
                b.set(qn("w:val"),   "single")
                b.set(qn("w:sz"),    str(size))
                b.set(qn("w:space"), "0")
                b.set(qn("w:color"), color)
                tcBords.append(b)
            tcPr.append(tcBords)

def _set_col_widths(table, widths_cm: List[float]):
    for row in table.rows:
        for i, w in enumerate(widths_cm):
            if i < len(row.cells):
                row.cells[i].width = Cm(w)

def _page_break(doc: Document):
    p   = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    run = p.add_run()
    run.add_break(WD_BREAK.PAGE)

def _para(doc: Document, text: str, bold: bool = False, size: int = 10,
          color: RGBColor = DOCX_TEXT, align=WD_ALIGN_PARAGRAPH.LEFT,
          space_before: float = 0, space_after: float = 6) -> None:
    p   = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after  = Pt(space_after)
    run = p.add_run(text)
    run.bold           = bold
    run.font.size      = Pt(size)
    run.font.color.rgb = color
    run.font.name      = "Calibri"

def _banner(doc: Document, text: str):
    tbl  = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.rows[0].cells[0]
    _set_cell_bg(cell, DOCX_MAROON)
    _set_cell_borders(cell, size=0, color=_hex(DOCX_MAROON))
    _cell_margins(cell, top=80, bottom=80, left=160, right=160)
    cell.text = ""
    p   = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.bold           = True
    run.font.size      = Pt(13)
    run.font.color.rgb = DOCX_WHITE
    run.font.name      = "Calibri"
    for row in tbl.rows:
        for c in row.cells:
            c.width = Cm(14.0)

def _build_two_col_table(doc: Document, rows: List[Tuple[str, str]],
                         grey_indices: Optional[Set[int]] = None,
                         all_grey: bool = False,
                         col_widths_cm: Tuple[float, float] = (9.0, 8.5)):
    if grey_indices is None:
        grey_indices = set()
    tbl = doc.add_table(rows=len(rows), cols=2)
    tbl.style = "Table Grid"
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    for r_idx, (label, value) in enumerate(rows):
        lc = tbl.rows[r_idx].cells[0]
        vc = tbl.rows[r_idx].cells[1]
        if all_grey or r_idx in grey_indices:
            _set_cell_bg(lc, DOCX_GREY)
            _set_cell_bg(vc, DOCX_GREY)
        _write_cell(lc, label, bold=True,  color=DOCX_RED,  size=10)
        _write_cell(vc, value, bold=False, color=DOCX_TEXT, size=10)
    _set_col_widths(tbl, list(col_widths_cm))
    _set_table_outer_border(tbl, size=8)
    _force_cell_borders(tbl, size=8)

# =========================================================
# DOCX PAGE BUILDERS
# =========================================================

def docx_page1_cover(doc: Document, payload: Dict[str, Any],
                     app_info: Dict[str, str], logo_path: Optional[str]):
    if logo_path:
        logo_para = doc.add_paragraph()
        logo_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        logo_para.paragraph_format.space_before = Pt(0)
        logo_para.paragraph_format.space_after  = Pt(8)
        run = logo_para.add_run()
        run.add_picture(logo_path, width=Cm(8.0))

    title_tbl = doc.add_table(rows=1, cols=1)
    title_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_outer_border(title_tbl, size=16)
    tc = title_tbl.rows[0].cells[0]
    _set_cell_borders(tc, size=16)
    _cell_margins(tc, top=120, bottom=120, left=200, right=200)
    tc.text = ""
    tp = tc.paragraphs[0]
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = tp.add_run("National Nursing Assessment Service Advisory Report")
    tr.bold           = True
    tr.font.size      = Pt(18)
    tr.font.color.rgb = DOCX_BLACK
    tr.font.name      = "Calibri"
    for row in title_tbl.rows:
        for cell in row.cells:
            cell.width = Cm(17.0)
    doc.add_paragraph()

    _para(doc, "Expedited Service", bold=False, size=12, color=DOCX_RED,
          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)

    meta_pairs = [
        ("Nursing Group:",      app_info.get("nursing_group", "")),
        ("Province:",           app_info.get("province", "")),
        ("NNAS ID Number:",     app_info.get("nnas_id", "")),
        ("Application Number:", app_info.get("application_number", "")),
    ]
    meta_tbl = doc.add_table(rows=2, cols=2)
    meta_tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    for i, (label, val) in enumerate(meta_pairs):
        row_i = i // 2
        col_i = i %  2
        cell  = meta_tbl.rows[row_i].cells[col_i]
        cell.text = ""
        _cell_margins(cell, top=40, bottom=40, left=0, right=0)
        _set_cell_borders(cell, size=0, color="FFFFFF")
        p    = cell.paragraphs[0]
        run1 = p.add_run(label + "  ")
        run1.bold           = True
        run1.font.size      = Pt(11)
        run1.font.color.rgb = DOCX_RED
        run1.font.name      = "Calibri"
        run2 = p.add_run(val)
        run2.font.size      = Pt(11)
        run2.font.color.rgb = DOCX_TEXT
        run2.font.name      = "Calibri"
    for row in meta_tbl.rows:
        for cell in row.cells:
            cell.width = Cm(8.75)
    doc.add_paragraph()

    _para(doc,
          "This report does not determine an applicant\'s eligibility to become registered or "
          "licensed in the jurisdiction of application. A regulatory body reserves the right to "
          "request additional information pertaining to the findings provided in this report and "
          "has the authority to require further testing and assessment, in accordance with the "
          "regulatory body\'s legislation.",
          size=9, color=DOCX_TEXT, space_after=6)

    _para(doc,
          "Disclaimer: The purpose of this Advisory Report is to support the Canadian nursing "
          "registration decisions of NNAS member nursing regulatory bodies. NNAS does not intend "
          "it for any other use nor does it endorse it for any other use. NNAS is not responsible "
          "or liable for anything that results from using this Report for any other purpose or for "
          "the use of the Report by a person who is not an NNAS member nursing regulatory body.",
          size=9, color=DOCX_RED, space_after=16)

    notes             = payload.get("assessment_notes") or {}
    base_public       = safe_str(notes.get("public"))
    identity_note     = safe_str(notes.get("identity_section_note"))
    education_note    = safe_str(notes.get("education_section_note"))
    registration_note = safe_str(notes.get("registration_section_note"))

    section_lines = []
    if identity_note:
        section_lines.append(f"Identity Section Note: {identity_note}")
    if education_note:
        section_lines.append(f"Education Section Note: {education_note}")
    if registration_note:
        section_lines.append(f"Registration Section Note: {registration_note}")

    appended        = "\n".join(section_lines)
    public_text     = "\n\n".join(p for p in [base_public, appended] if p)
    assessment_text = safe_str(notes.get("assessment"))

    def _notes_box(header: str, body: str):
        tbl = doc.add_table(rows=2, cols=1)
        tbl.style = "Table Grid"
        tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
        hc = tbl.rows[0].cells[0]
        _set_cell_bg(hc, DOCX_GREY)
        _cell_margins(hc, top=80, bottom=80, left=100, right=100)
        hc.text = ""
        hp = hc.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        hr = hp.add_run(header)
        hr.bold           = True
        hr.font.size      = Pt(14)
        hr.font.color.rgb = DOCX_RED
        hr.font.name      = "Calibri"
        bc = tbl.rows[1].cells[0]
        _cell_margins(bc, top=80, bottom=80, left=120, right=120)
        _write_cell(bc, body, bold=False, color=DOCX_TEXT, size=10)
        _set_table_outer_border(tbl, size=8)
        _force_cell_borders(tbl, size=8)
        doc.add_paragraph()

    if public_text:
        _notes_box("Public Assessment Notes", public_text)
    if assessment_text:
        _notes_box("Assessment Notes", assessment_text)

def docx_page2_personal_info(doc: Document, payload: Dict[str, Any]):
    _banner(doc, "Part A: Applicant Data Review")
    doc.add_paragraph()
    _para(doc, "A1: PERSONAL INFORMATION",
          bold=True, size=14, color=DOCX_RED,
          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
    rows = personal_info_rows(payload)
    _build_two_col_table(doc, rows, grey_indices={0, 2, 4, 6})

def docx_page3_registration(doc: Document, payload: Dict[str, Any]):
    regs = extract_registration_entries(payload)
    if not regs:
        return
    for idx, reg in enumerate(regs):
        if idx > 0:
            _page_break(doc)
        if idx == 0:
            # Title shows once, before the first entry only -- matches
            # the PDF, where "A3: REGISTRATION/LICENSE OVERVIEW" comes
            # from static text pre-printed on the template's first A3
            # page (see render_registration_tables' docstring); every
            # subsequently inserted PDF page has no title at all, since
            # render_registration_tables never draws one in code.
            _para(doc, "A3: REGISTRATION/LICENSE OVERVIEW",
                  bold=True, size=14, color=DOCX_RED,
                  align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
        rows = registration_rows(reg)
        _build_two_col_table(doc, rows, all_grey=True)

def docx_page4a_a2_overview(doc: Document, programs: List[Dict[str, Any]]):
    """
    Render A2 (Nursing Education Overview) for every education program,
    grouped together (one page per program). No banner is drawn here --
    docx_page2_personal_info already draws the single "Part A: Applicant
    Data Review" banner immediately before A1, and per the PDF layout
    that banner covers all of Part A (A1, A2, A3), not just A1.

    Split out of what used to be a single docx_page4_part_b_education
    (which rendered A2 AND B1 back to back) so build_docx can place A3
    between this and docx_page4b_b1_breakdown -- the combined version
    put A2 and B1 immediately adjacent, which meant A3 could only be
    called either fully before or fully after both, and it was placed
    before both by mistake. That produced A1 -> A3 -> A2 -> B1, when the
    PDF (and the A1/A2/A3 numbering itself) puts A2 before A3.

    The "A2: NURSING EDUCATION OVERVIEW" title itself shows once, before
    the first program only -- matches the PDF, where render_a2_on_page
    never draws that title in code at all; whatever appears on the
    PDF's first A2 page is static text pre-printed on the template, and
    every subsequently inserted PDF page has no title.
    """
    if not programs:
        _para(doc, "A2: NURSING EDUCATION OVERVIEW",
              bold=True, size=14, color=DOCX_RED,
              align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
        _para(doc, "Nursing education overview not available.", size=10, color=DOCX_TEXT)
        return

    for idx, prog in enumerate(programs):
        if idx > 0:
            _page_break(doc)

        a2 = prog.get("a2") or {}

        if idx == 0:
            _para(doc, "A2: NURSING EDUCATION OVERVIEW",
                  bold=True, size=14, color=DOCX_RED,
                  align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)

        tbl = doc.add_table(rows=len(_A2_DISPLAY_ROWS), cols=2)
        tbl.style = "Table Grid"
        tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
        grey_rows = {0, 2, 4, 6, 8, 10, 12}
        for r_idx, (label, key) in enumerate(_A2_DISPLAY_ROWS):
            lc = tbl.rows[r_idx].cells[0]
            vc = tbl.rows[r_idx].cells[1]
            if r_idx in grey_rows:
                _set_cell_bg(lc, DOCX_GREY)
                _set_cell_bg(vc, DOCX_GREY)
            _write_cell(lc, label,           bold=True,  color=DOCX_RED,  size=10)
            _write_cell(vc, a2.get(key, ""), bold=False, color=DOCX_TEXT, size=10)
        for row in tbl.rows:
            row.cells[0].width = Cm(9.0)
            row.cells[1].width = Cm(8.5)
        _set_table_outer_border(tbl, size=8)
        _force_cell_borders(tbl, size=8)


def docx_page4b_b1_breakdown(doc: Document, programs: List[Dict[str, Any]]):
    """
    Render the "Part B: Nursing Education Analysis" banner and B1
    (Domain Breakdown) for every education program, grouped together
    (one page per program, each labeled with its matching school name
    and program dates). See docx_page4a_a2_overview's docstring for why
    this is a separate function from A2 rather than the combined one it
    used to be.

    The "B1: RN / PN EDUCATION DOMAIN BREAKDOWN" title shows once,
    before the first program only -- matches the PDF, where
    render_b1_on_page never draws that literal title in code; whatever
    appears on the PDF's first B1 page is static template text. The
    school name + date range caption, by contrast, DOES repeat on every
    page in both formats -- render_b1_on_page draws it fresh via
    a2_identifying_line() on every PDF page too, since without it a
    reader would have no way to tell which program a given B1 table
    belongs to.
    """
    _banner(doc, "Part B: Nursing Education Analysis")
    doc.add_paragraph()

    if not programs:
        _para(doc, "B1: RN / PN EDUCATION DOMAIN BREAKDOWN",
              bold=True, size=14, color=DOCX_RED,
              align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
        _para(doc, "No education domain breakdown data available.", size=10, color=DOCX_TEXT)
        return

    b1_headers    = [
        "SUBJECT AREA", "THEORY\nHOURS",
        "SIMULATION/\nLAB\nHOURS", "CLINICAL\nHOURS\n(Direct Patient Care)",
    ]
    b1_col_widths = [8.5, 2.2, 2.5, 3.3]

    for idx, prog in enumerate(programs):
        if idx > 0:
            _page_break(doc)

        a2 = prog.get("a2") or {}
        b1 = prog.get("b1") or []

        if idx == 0:
            _para(doc, "B1: RN / PN EDUCATION DOMAIN BREAKDOWN",
                  bold=True, size=14, color=DOCX_RED,
                  align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)

        school_name, date_range = a2_identifying_line(a2)
        if school_name:
            school_p = doc.add_paragraph()
            school_p.paragraph_format.space_after = Pt(2)
            run_school = school_p.add_run(school_name)
            run_school.bold           = True
            run_school.font.size      = Pt(11)
            run_school.font.color.rgb = DOCX_TEXT
            run_school.font.name      = "Calibri"
        if date_range:
            dates_p = doc.add_paragraph()
            dates_p.paragraph_format.space_after = Pt(6)
            run_dates = dates_p.add_run(date_range)
            run_dates.font.size      = Pt(10)
            run_dates.font.color.rgb = DOCX_TEXT
            run_dates.font.name      = "Calibri"

        tbl2 = doc.add_table(rows=1 + len(b1), cols=4)
        tbl2.style = "Table Grid"
        tbl2.alignment = WD_TABLE_ALIGNMENT.LEFT

        for c_idx, hdr in enumerate(b1_headers):
            cell = tbl2.rows[0].cells[c_idx]
            _set_cell_bg(cell, DOCX_GREY)
            _write_cell(cell, hdr, bold=True, color=DOCX_RED, size=9,
                        align=WD_ALIGN_PARAGRAPH.CENTER)

        for r_idx, row_dict in enumerate(b1, start=1):
            is_totals = row_dict["subject"].strip().upper() == "TOTALS"
            bg        = DOCX_GREY if (r_idx % 2 == 0 or is_totals) else None
            values    = [
                row_dict["subject"],
                _b1_display(row_dict["theory_hours"]),
                _b1_display(row_dict["sim_lab_hours"]),
                _b1_display(row_dict["clinical_hours"]),
            ]
            for c_idx, val in enumerate(values):
                cell = tbl2.rows[r_idx].cells[c_idx]
                if bg:
                    _set_cell_bg(cell, bg)
                align = (WD_ALIGN_PARAGRAPH.CENTER if c_idx > 0
                         else WD_ALIGN_PARAGRAPH.LEFT)
                _write_cell(cell, val, bold=is_totals, color=DOCX_TEXT,
                            size=9, align=align)

        for row in tbl2.rows:
            for c_idx, cw in enumerate(b1_col_widths):
                row.cells[c_idx].width = Cm(cw)
        _set_table_outer_border(tbl2, size=8)
        _force_cell_borders(tbl2, size=8)

def docx_page5_evaluator(doc: Document, evaluator_name: str = "", date_str: str = ""):
    body_text = (
        "This report has been generated based on the compilation of applicant\'s data and the "
        "evaluation of nursing education coursework completed by the applicant. The evaluations "
        "contained herein are based on the criteria approved by the member regulatory bodies of "
        "the National Nursing Assessment Service and other information sources as noted. The "
        "evaluations reflect the general comparability of foreign nursing education attainment "
        "when measured against current Canadian nursing entry-to-practice benchmarks and are not "
        "intended to denote exact equivalence.\n\n"
        "All nursing licenses/registrations received, nursing employment history, and language "
        "proficiency test results, as applicable are also included in this report. Unless "
        "otherwise stated, all transcripts, verification of registrations, and other documents "
        "used in the evaluations were received directly from the source (e.g. educational "
        "institution, regulatory authority, employer, language testing agency).\n\n"
        "This report is advisory in nature, and is intended to provide information to the nursing "
        "regulatory body in the jurisdiction of application. This report only represents one "
        "component considered by a regulatory body in assessing an application for registration; "
        "it is not intended therefore as a substitute for the autonomous assessment and "
        "decision-making authority of the nursing regulatory body receiving this report.\n\n"
        "This report does not determine an applicant\'s eligibility to become registered or "
        "licensed in the jurisdiction of application. Regardless of the report content, a "
        "regulatory body reserves the right to request additional information pertaining to the "
        "findings provided in this report and has the authority to require further testing and "
        "assessment, in accordance with the regulatory body\'s legislation."
    )
    for para_text in body_text.split("\n\n"):
        _para(doc, para_text.strip(), size=10, color=DOCX_TEXT, space_after=8)
    doc.add_paragraph()
    ev_tbl = doc.add_table(rows=1, cols=2)
    ev_tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    for i, (label, value) in enumerate((("Evaluator:", evaluator_name), ("Date:", date_str))):
        cell = ev_tbl.rows[0].cells[i]
        _set_cell_borders(cell, size=0, color="FFFFFF")
        _cell_margins(cell, top=40, bottom=40, left=0, right=0)
        cell.text = ""
        p    = cell.paragraphs[0]
        run1 = p.add_run(label + "  ")
        run1.font.size      = Pt(10)
        run1.font.color.rgb = DOCX_RED
        run1.font.name      = "Calibri"
        run2 = p.add_run(value)
        run2.font.size      = Pt(10)
        run2.font.color.rgb = DOCX_TEXT
        run2.font.name      = "Calibri"
        cell.width = Cm(8.75)

def docx_page6_part_c(doc: Document, payload: Dict[str, Any]):
    sections = build_part_c_sections(payload)
    _banner(doc, "Part C: Supporting Evidence")
    doc.add_paragraph()
    _para(doc,
          "The documents transmitted below represent an aggregate of information that was "
          "received under each heading. This aggregate includes all required "
          "verified/authenticated documents that were used in the creation of the Advisory "
          "Report as well as documents submitted that were not used for the purpose of the "
          "Advisory Report creation, these documents may or may not have undergone specific "
          "verification/authentication procedure due to their superfluous nature.",
          size=10, color=DOCX_TEXT, space_after=10)
    if not sections:
        return
    for sec in sections:
        header  = safe_str(sec.get("header"))
        bullets = sec.get("bullets") or []
        if not isinstance(bullets, list): bullets = [safe_str(bullets)]
        if header:
            _para(doc, header, bold=True, size=14, color=DOCX_RED,
                  space_before=8, space_after=4)
        for b in bullets:
            b = safe_str(b)
            if not b: continue
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after  = Pt(1)
            p.paragraph_format.left_indent  = Cm(1.0)
            r = p.add_run(f"\u25a0  {b}")
            r.font.size      = Pt(10)
            r.font.name      = "Calibri"
            r.font.color.rgb = DOCX_TEXT

# =========================================================
# DOCX MAIN BUILD
# =========================================================

def build_docx(payload: Dict[str, Any], app_info: Dict[str, str],
               logo_path: Optional[str], out_path: str,
               evaluator_name: str = "", date_str: str = "",
               education_data: Optional[Dict[str, Any]] = None) -> str:
    doc = Document()
    for sec in doc.sections:
        sec.top_margin    = Cm(1.8)
        sec.bottom_margin = Cm(1.8)
        sec.left_margin   = Cm(1.8)
        sec.right_margin  = Cm(1.8)

    docx_page1_cover(doc, payload, app_info, logo_path)

    _page_break(doc)
    docx_page2_personal_info(doc, payload)

    programs = normalize_education_programs(education_data)

    _page_break(doc)
    docx_page4a_a2_overview(doc, programs)

    regs = extract_registration_entries(payload)
    if regs:
        _page_break(doc)
        docx_page3_registration(doc, payload)

    _page_break(doc)
    docx_page4b_b1_breakdown(doc, programs)

    _page_break(doc)
    docx_page5_evaluator(doc, evaluator_name=evaluator_name, date_str=date_str)

    _page_break(doc)
    docx_page6_part_c(doc, payload)

    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    doc.save(out_path)
    return out_path

# =========================================================
# ENTRYPOINT
# =========================================================

def main(appid: Any,
         additional_content: Optional[str] = None,
         reference_id: Optional[str] = None,
         application_title: Optional[str] = None,
         nnas_id: Optional[int] = None,
         provider: str = ""):

    payload_base, dbg = parse_additional_content(additional_content)
    app_id             = coerce_appid(appid)

    resolved_provider  = resolve_provider(provider, payload_base)
    cfg                = PROVIDER_CONFIG[resolved_provider]
    program_id         = cfg["program_id"]

    application_number = parse_application_number(application_title)
    nnas_id_str        = str(nnas_id) if nnas_id is not None else ""
    evaluator_name     = "NNAS Expedited Service"
    date_str           = datetime.utcnow().strftime("%Y-%m-%d")

    if not reference_id:
        reference_id = safe_str(payload_base.get("reference_id")) or str(app_id)

    # ---- Fetch every SMApply review submission ONCE. Registration,
    # ---- identity documents, and assessment notes are built entirely
    # ---- from these live review responses (replacing what
    # ---- additional_content used to supply for those sections);
    # ---- personal_info and application_info still come from
    # ---- additional_content, since reviews don't carry them. The same
    # ---- fetched list is reused below for the education fallback, so
    # ---- this API call happens at most once per run.
    reviews          = fetch_all_reviews(app_id, program_id)
    payload, rev_dbg = build_payload_from_reviews(reviews, payload_base)
    dbg["reviews"]   = rev_dbg

    app_info_list = build_application_info(
        payload,
        application_number=application_number,
        nnas_id=nnas_id_str,
    )
    if not app_info_list:
        raise ValueError("No regulatory bodies found in application_info.")

    # ---- Education data: prior report PDF first, manual SMApply review
    # ---- responses (the reviews list already fetched above) as a
    # ---- fallback when no prior report is found. If the reviewer says a
    # ---- prior report exists but the file itself is missing from
    # ---- SharePoint, that's a data-integrity problem (likely a missed
    # ---- upload) rather than a normal "no prior report" case, so it
    # ---- raises instead of silently falling back to manual review data.
    prior_folder   = f"{cfg['applications_base']}/{reference_id}"
    prior_pdf_path = find_prior_report_sharepoint(ctx(), prior_folder)
    if prior_pdf_path:
        try:
            education_data = extract_nnas_education(prior_pdf_path)
            dbg["prior_report_found"] = prior_pdf_path
            dbg["education_source"]  = "prior_report_pdf"
        except Exception as exc:
            education_data = {}
            dbg["prior_report_error"] = str(exc)
    else:
        education_data = {}
        dbg["prior_report_found"] = None

        prior_report_flag = _get_prior_report_flag(reviews)
        dbg["prior_report_flag"] = prior_report_flag

        if prior_report_flag == "Yes":
            raise RuntimeError(
                "Reviewer indicated a previously issued advisory report "
                f"exists for reference_id={reference_id}, but no "
                f"'_priorreport.pdf' was found in {prior_folder}. This "
                "looks like a missing upload rather than a normal "
                "no-prior-report case -- please confirm the prior report "
                "has been placed in that SharePoint folder before "
                "re-running."
            )

    # ---- Evaluator-hardcoded rows to append to a prior-report program's
    # ---- domain breakdown (for reports missing some of the 20 canonical
    # ---- B1 subjects). Only applies to the prior-report branch above --
    # ---- the SMApply gating question itself requires "Does this
    # ---- applicant have a previously issued advisory report?" = Yes
    # ---- before the append option is even shown to the evaluator, so
    # ---- this is skipped entirely on the manual-fallback path below.
    if dbg.get("education_source") == "prior_report_pdf":
        try:
            append_results = parse_appended_domain_rows(reviews)
            education_data = apply_appended_domain_rows(education_data, append_results)
            # Recompute TOTALS for every program, regardless of whether
            # it received any appended rows -- recomputing is idempotent
            # for a program nothing changed on (reproduces its original
            # total unchanged, see recompute_domain_totals), and correct
            # for every program that did receive appended rows. Simpler
            # and just as safe as tracking exactly which programs from
            # append_results actually got modified.
            for prog in education_data.get("programs") or []:
                recompute_domain_totals(prog.get("b1") or [])
            dbg["appended_domain_rows_blocks"] = [
                {"program_index": idx, "row_count": len(rows)} for idx, rows in append_results
            ]
        except Exception as exc:
            dbg["appended_domain_rows_error"] = str(exc)

        # ---- Evaluator can omit specific prior-report education
        # ---- programs from this report. MUST run after the append-rows
        # ---- block above -- both features address programs via the
        # ---- same original "Education Document N" numbering, and
        # ---- omitting first would shift indices out from under
        # ---- append's own targets. See apply_omitted_education_
        # ---- documents' docstring for the full ordering rationale.
        try:
            omit_indices = parse_omitted_education_document_indices(reviews)

            # Grab each omitted program's school name BEFORE removing it
            # from education_data["programs"] below -- once it's gone,
            # there's nothing left to match against when filtering the
            # corresponding bullets out of Part C's "C2: Nursing
            # Education Documents" section further down in main().
            programs_before_omit = education_data.get("programs") or []
            omitted_school_names = [
                _collapse_ws((programs_before_omit[i].get("a2") or {}).get("school_name", ""))
                for i in omit_indices
                if 0 <= i < len(programs_before_omit)
            ]
            omitted_school_names = [n for n in omitted_school_names if n]

            education_data = apply_omitted_education_documents(education_data, omit_indices)
            dbg["omitted_education_document_indices"]     = omit_indices
            dbg["omitted_education_document_school_names"] = omitted_school_names
        except Exception as exc:
            omitted_school_names = []
            dbg["omitted_education_documents_error"] = str(exc)
    else:
        omitted_school_names = []

    if not normalize_education_programs(education_data):
        try:
            education_data = parse_manual_education_from_reviews(reviews)
            dbg["education_source"] = "smapply_manual_review"
        except Exception as exc:
            education_data = {}
            dbg["manual_education_error"] = str(exc)

    education_programs = normalize_education_programs(education_data)
    dbg["education_program_count"] = len(education_programs)

    # ---- Part C: port the "Nursing Education Documents" list over from
    # ---- the prior report verbatim (when one was found above), filed as
    # ---- "C2: Nursing Education Documents" in this report. No
    # ---- renumbering of the review-derived C1/C3/C4 sections is needed
    # ---- since that adapter never produced a C2. See
    # ---- apply_part_c_education_documents for the reasoning on why this
    # ---- section is NOT trimmed to the newer one-document convention.
    payload = apply_part_c_education_documents(payload, prior_pdf_path)

    # ---- Drop C2 bullets belonging to any program the evaluator chose
    # ---- to omit (see omitted_school_names above). C2 is scraped
    # ---- verbatim from the prior report's own Part C text and keyed by
    # ---- school name in its own bullet lines, not by education_data
    # ---- index, so it doesn't stay in sync with an omitted program on
    # ---- its own -- without this step, a removed program's rows would
    # ---- disappear from A2/B1 while its documents kept showing up in
    # ---- Part C, contradicting the rest of the report.
    if omitted_school_names:
        for sec in payload.get("part_c_sections") or []:
            if str(sec.get("header", "")).startswith("C2"):
                sec["bullets"] = filter_out_omitted_school_bullets(
                    sec.get("bullets") or [], omitted_school_names
                )
        dbg["part_c_c2_bullets_filtered_for_omitted_schools"] = omitted_school_names

    dbg["part_c_education_documents_ported"] = sum(
        len(s.get("bullets") or [])
        for s in (payload.get("part_c_sections") or [])
        if str(s.get("header", "")).startswith("C2")
    )

    # ---- Part C: C4 ("Other Documents") is built entirely from the
    # ---- "Additional documents" SharePoint folder -- step through each
    # ---- of its direct subfolders, and for any subfolder that actually
    # ---- contains files, list those files. `prior_folder` is the
    # ---- application's root folder (applications/<provider>/
    # ---- <reference_id>), the same path already used above for the
    # ---- prior-report lookup; "Additional documents" is a sibling of
    # ---- "Credential reports" under that same root.
    c4_file_bullets = list_additional_documents_files_by_subfolder(prior_folder)
    payload = apply_c4_additional_documents_files(payload, c4_file_bullets)
    dbg["c4_additional_documents_files"] = c4_file_bullets

    dbg.update({
        "appid_type":           str(type(appid)),
        "appid_used":           app_id,
        "reference_id_used":    reference_id,
        "application_title":    application_title,
        "application_number":   application_number,
        "nnas_id":              nnas_id_str,
        "provider_resolved":    resolved_provider,
        "program_id":           program_id,
        "rb_count":             len(app_info_list),
        "rb_list":              [a["regulatory_body"] for a in app_info_list],
        "province_list":        [a["province"]        for a in app_info_list],
        "has_assessment_notes": bool(payload.get("assessment_notes")),
        "has_personal_info":    bool(payload.get("personal_info")),
        "has_registration":     bool(payload.get("registration")),
        "pi_keys_sample":       sorted(list((payload.get("personal_info") or {}).keys()))[:12],
        "reg_keys_sample":      sorted(list((payload.get("registration") or {}).keys()))[:12],
        "reg_entries_found":    len(extract_registration_entries(payload)),
        "part_c_headers":       [safe_str(s.get("header")) for s in (payload.get("part_c_sections") or [])],
    })

    template_filename = TEMPLATE_FILENAME_FULL if education_programs else TEMPLATE_FILENAME
    template  = download_template(template_filename)
    logo_path = download_logo()
    dbg["template_used"]   = template_filename
    dbg["logo_downloaded"] = logo_path is not None

    ts = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    os.makedirs("tmp", exist_ok=True)

    application_folder_id = reference_id
    _smapply_lock         = threading.Lock()

    # ---------------------------------------------------------
    # RESTORED FROM SCRIPT 1: clear any files already on the SMApply
    # upload task (per-provider task id via cfg["smapply_task_id"]) so a
    # re-run replaces rather than accumulates. Done ONCE here, not inside
    # _process_rb, because every RB uploads to the same task and clearing
    # per-upload would delete files uploaded by earlier RBs in the same
    # run.
    # ---------------------------------------------------------
    dbg["smapply_files_cleared"] = clear_smapply_task_files(app_id, cfg["smapply_task_id"])

    def _process_rb(idx_app_info):
        idx, app_info = idx_app_info
        rb_full    = app_info.get("regulatory_body", "")
        shortname  = rb_short_name(rb_full)
        group_code = app_info.get("nursing_group_code", "")

        group_tag        = f"_{group_code}" if group_code else ""
        base_name        = f"{reference_id}_{shortname}{group_tag}_{ts}"
        pdf_name_plain   = f"{base_name}.pdf"
        docx_name_plain  = f"{base_name}.docx"
        pdf_name_latest  = f"{base_name}_latest.pdf"
        docx_name_latest = f"{base_name}_latest.docx"

        pdf_path  = os.path.join("tmp", pdf_name_plain)
        rb_writer = populate_pdf(
            template, payload, app_info,
            evaluator_name=evaluator_name, date_str=date_str,
            education_data=education_data,
        )
        with open(pdf_path, "wb") as fh:
            rb_writer.write(fh)

        docx_path = os.path.join("tmp", docx_name_plain)
        build_docx(
            payload, app_info, logo_path, docx_path,
            evaluator_name=evaluator_name, date_str=date_str,
            education_data=education_data,
        )

        folder_path = (
            f"{cfg['applications_base']}/{application_folder_id}/"
            f"{CRED_REPORTS_SUBFOLDER}/{shortname}"
        )
        sp_ctx = ctx()
        _ensure_sp_folder_tree(sp_ctx, folder_path)
        _demote_latest_in_folder(sp_ctx, folder_path,
                                 name_filter=group_tag if group_tag else None)

        sp_pdf  = upload_to_sharepoint_folder(pdf_path,  pdf_name_latest,  folder_path,
                                              sp_ctx=sp_ctx, ensure_folder=False)
        sp_docx = upload_to_sharepoint_folder(docx_path, docx_name_latest, folder_path,
                                              sp_ctx=sp_ctx, ensure_folder=False)

        with _smapply_lock:
            sm_pdf = upload_pdf_to_smapply_task(
                appid=app_id, task_id=cfg["smapply_task_id"],
                pdf_filepath=pdf_path, upload_filename=pdf_name_plain,
            )
            sm_docx = upload_docx_to_smapply_task(
                appid=app_id, task_id=cfg["smapply_task_id"],
                docx_filepath=docx_path, upload_filename=docx_name_plain,
            )

        return idx, {
            "regulatory_body":   rb_full,
            "short_name":        shortname,
            "province":          app_info.get("province", ""),
            "sharepoint_folder": _sp_url(folder_path),
            "sharepoint_pdf":    sp_pdf,
            "sharepoint_docx":   sp_docx,
            "smapply_pdf":       sm_pdf,
            "smapply_docx":      sm_docx,
        }

    workers   = min(SP_MAX_WORKERS, len(app_info_list)) or 1
    indexed   = list(enumerate(app_info_list))
    collected: List[Tuple[int, Dict[str, Any]]] = []
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = [pool.submit(_process_rb, item) for item in indexed]
        for fut in as_completed(futures):
            collected.append(fut.result())
    reports = [r for _, r in sorted(collected, key=lambda t: t[0])]

    smapply_task_completed = mark_smapply_task_complete(
        appid=app_id, task_id=cfg["smapply_task_id"]
    )
    dbg["smapply_patch_ok"]    = smapply_task_completed.get("ok")
    dbg["smapply_patch_error"] = (
        None if smapply_task_completed.get("ok")
        else smapply_task_completed.get("response")
    )

    return {
        "status":                 "ok" if smapply_task_completed.get("ok") else "task_complete_failed",
        "reference_id":           reference_id,
        "report_count":           len(reports),
        "reports":                reports,
        "smapply_task_completed": smapply_task_completed,
        "debug":                  dbg,
    }