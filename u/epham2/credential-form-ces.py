# py: 3.12
# requirements:
# pypdf
# reportlab
# requests
# wmill
# python-docx
# pillow==10.4.0 --no-binary=:all:

import os
import json
import re
import requests
from io import BytesIO
from typing import Dict, List, Any, Optional, Tuple, Set
from urllib.parse import quote
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading

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
# SHAREPOINT / MICROSOFT GRAPH CONFIG
# =========================================================

TENANT_ID     = wmill.get_variable("u/epham2/AZURE_TENANT_ID")
CLIENT_ID     = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_ID")
CLIENT_SECRET = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_SECRET")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SHAREPOINT_HOSTNAME = "nnasca.sharepoint.com"
SITE_PATH = "/sites/NNAS-Portal-Sharepoint"
DOCUMENT_LIBRARY_NAME = "Applications"
SITE_WEB_URL = "https://nnasca.sharepoint.com/sites/NNAS-Portal-Sharepoint"

TEMPLATE_FOLDER   = "Applications/Credential form"
TEMPLATE_FILENAME = "CredentialReportBaseV1.pdf"
LOGO_FILENAME     = "Credential_header.png"

# --- CES-specific config (kept from the original CES script) ---
APPLICATIONS_BASE      = "applications/ces"
CRED_REPORTS_SUBFOLDER = "Credential reports"

SP_MAX_WORKERS = 8

PAGE_NOTES = 0
PAGE_A1    = 1
PAGE_A2    = 2

_GRAPH_TOKEN = None
_SP_SITE_ID = None
_SP_DRIVE_ID = None
_SP_LIST_ID = None
_GRAPH_INIT_LOCK = threading.Lock()
_GRAPH_TOKEN_LOCK = threading.Lock()


# =========================================================
# SMAPPLY CONFIG
# =========================================================

# --- CES-specific config (kept from the original CES script) ---
SMAPPLY_TASK_ID = 2053943

SMAPPLY_BASE = "https://nnas.smapply.io/api"
SMAPPLY_REVIEWS_TOKEN_VAR = "u/epham2/SMAPPLY_ACCESS_TOKEN"


# =========================================================
# SHAREPOINT / MICROSOFT GRAPH HELPERS
# =========================================================

def ctx():
    return None


def _library_relative_path(path: str) -> str:
    p = (path or "").strip().strip("/")
    if p.lower().startswith("sites/nnas-portal-sharepoint/"):
        p = p[len("sites/NNAS-Portal-Sharepoint/"):]
    if p.lower().startswith("applications/"):
        p = p[len("Applications/"):]
    return p.strip("/")


def _sp_url(folder_path: str) -> str:
    return f"/sites/NNAS-Portal-Sharepoint/Applications/{_library_relative_path(folder_path)}"


def _sp_web_url(folder_path: str) -> str:
    return f"{SITE_WEB_URL}/Applications/{_encode_graph_path(_library_relative_path(folder_path))}"


def _encode_graph_path(path: str) -> str:
    return quote(path.strip("/"), safe="/")


def _get_graph_token() -> str:
    global _GRAPH_TOKEN

    if _GRAPH_TOKEN:
        return _GRAPH_TOKEN

    with _GRAPH_TOKEN_LOCK:
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
        timeout=180,
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
        timeout=180,
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
    global _SP_SITE_ID, _SP_DRIVE_ID, _SP_LIST_ID

    if _SP_SITE_ID and _SP_DRIVE_ID and _SP_LIST_ID:
        return

    with _GRAPH_INIT_LOCK:
        if _SP_SITE_ID and _SP_DRIVE_ID and _SP_LIST_ID:
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

        lists = _graph_request(
            "GET",
            f"/sites/{_SP_SITE_ID}/lists?$select=id,displayName",
            expected_statuses=(200,),
        )["value"]

        sp_list = next((l for l in lists if l.get("displayName") == DOCUMENT_LIBRARY_NAME), None)
        if not sp_list:
            available = [l.get("displayName") for l in lists]
            raise RuntimeError(
                f"Could not find SharePoint list '{DOCUMENT_LIBRARY_NAME}'. "
                f"Available lists: {available}"
            )
        _SP_LIST_ID = sp_list["id"]


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


def _get_drive_item_content(item_id: str) -> bytes:
    _init_sharepoint_graph()
    response = requests.get(
        f"{GRAPH_BASE}/drives/{_SP_DRIVE_ID}/items/{item_id}/content",
        headers={"Authorization": f"Bearer {_get_graph_token()}"},
        timeout=300,
        allow_redirects=True,
    )
    if response.status_code != 200:
        raise RuntimeError(
            f"Failed downloading SharePoint item content.\n"
            f"Item ID: {item_id}\n"
            f"Status: {response.status_code}\n"
            f"Response body:\n{response.text[:2000]}"
        )
    return response.content


def _patch_driveitem_fields(drive_item_id: str, fields: dict) -> None:
    _init_sharepoint_graph()

    item = _graph_request(
        "GET",
        f"/drives/{_SP_DRIVE_ID}/items/{drive_item_id}?$expand=listItem",
        expected_statuses=(200,),
    )

    list_item = item.get("listItem")
    if not list_item or not list_item.get("id"):
        raise RuntimeError(f"Drive item {drive_item_id} has no listItem information.")

    _graph_request(
        "PATCH",
        f"/sites/{_SP_SITE_ID}/lists/{_SP_LIST_ID}/items/{list_item['id']}/fields",
        headers={"Content-Type": "application/json"},
        json=fields,
        expected_statuses=(200, 204),
    )


def _create_folder(parent_path: str, folder_name: str):
    _init_sharepoint_graph()

    parent_path = _library_relative_path(parent_path)
    full_path = f"{parent_path}/{folder_name}".strip("/")

    existing = _get_drive_item_by_path(full_path)
    if existing:
        return existing

    if parent_path:
        endpoint = f"/drives/{_SP_DRIVE_ID}/root:/{_encode_graph_path(parent_path)}:/children"
    else:
        endpoint = f"/drives/{_SP_DRIVE_ID}/root/children"

    response = requests.post(
        f"{GRAPH_BASE}{endpoint}",
        headers={
            "Authorization": f"Bearer {_get_graph_token()}",
            "Content-Type": "application/json",
        },
        json={
            "name": folder_name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "fail",
        },
        timeout=180,
    )

    if response.status_code == 201:
        return response.json()

    if response.status_code == 409:
        existing_after_conflict = _get_drive_item_by_path(full_path)
        if existing_after_conflict:
            return existing_after_conflict

    raise RuntimeError(
        f"Microsoft Graph folder create failed.\n"
        f"Parent path: {parent_path}\n"
        f"Folder name: {folder_name}\n"
        f"Full path: {full_path}\n"
        f"Status: {response.status_code}\n"
        f"Response body:\n{response.text[:5000]}"
    )


def _ensure_sp_folder_tree(sp_ctx, folder_path: str):
    parts = [p for p in _library_relative_path(folder_path).split("/") if p]
    current_path = ""
    current_item = None

    for part in parts:
        current_item = _create_folder(current_path, part)
        current_path = f"{current_path}/{part}".strip("/")

    return current_item


def _upload_bytes_to_sharepoint(content: bytes, filename: str, folder_path: str,
                                file_type: str = "Report") -> str:
    _init_sharepoint_graph()
    folder_path = _library_relative_path(folder_path)
    upload_path = f"{folder_path}/{filename}".strip("/")

    uploaded_file = _graph_request(
        "PUT",
        f"/drives/{_SP_DRIVE_ID}/root:/{_encode_graph_path(upload_path)}:/content",
        headers={"Content-Type": "application/octet-stream"},
        data=content,
        expected_statuses=(200, 201),
    )

    if file_type:
        _patch_driveitem_fields(uploaded_file["id"], {"FileType": file_type})

    return uploaded_file.get("webUrl") or f"{SITE_WEB_URL}/Applications/{_encode_graph_path(upload_path)}"


def download_template() -> str:
    drive_path = f"{_library_relative_path(TEMPLATE_FOLDER)}/{TEMPLATE_FILENAME}"
    item = _get_drive_item_by_path(drive_path)
    if not item:
        raise RuntimeError(f"SharePoint template not found: {drive_path}")

    os.makedirs("tmp", exist_ok=True)
    path = os.path.join("tmp", TEMPLATE_FILENAME)
    with open(path, "wb") as f:
        f.write(_get_drive_item_content(item["id"]))
    if os.path.getsize(path) == 0:
        raise RuntimeError(f"Downloaded template is empty: {path}")
    return path


def download_logo() -> Optional[str]:
    try:
        drive_path = f"{_library_relative_path(TEMPLATE_FOLDER)}/{LOGO_FILENAME}"
        item = _get_drive_item_by_path(drive_path)
        if not item:
            return None

        os.makedirs("tmp", exist_ok=True)
        path = os.path.join("tmp", LOGO_FILENAME)
        with open(path, "wb") as f:
            f.write(_get_drive_item_content(item["id"]))
        if os.path.getsize(path) == 0:
            return None
        return path
    except Exception:
        return None


def _demote_latest_in_folder(sp_ctx, folder_path: str,
                             name_filter: Optional[str] = None):
    folder_path = _library_relative_path(folder_path)
    folder = _get_drive_item_by_path(folder_path)
    if not folder:
        return

    try:
        children = _graph_request(
            "GET",
            f"/drives/{_SP_DRIVE_ID}/items/{folder['id']}/children?$select=id,name,file,folder",
            expected_statuses=(200,),
        ).get("value", [])
    except Exception:
        return

    for item in list(children):
        name = item.get("name", "")
        if not item.get("file"):
            continue
        if "." not in name:
            continue
        if name_filter is not None and name_filter not in name:
            continue
        stem, _, ext = name.rpartition(".")
        if not stem.endswith("_latest"):
            continue

        new_name = f"{stem[:-len('_latest')]}.{ext}"
        try:
            content = _get_drive_item_content(item["id"])
            demoted_url = _upload_bytes_to_sharepoint(
                content,
                new_name,
                folder_path,
                file_type="Report",
            )
        except Exception:
            continue

        try:
            _graph_request(
                "DELETE",
                f"/drives/{_SP_DRIVE_ID}/items/{item['id']}",
                expected_statuses=(204,),
            )
        except Exception:
            pass


def upload_to_sharepoint_folder(local_path: str, filename: str, folder_path: str,
                                file_type: str = "Report",
                                sp_ctx: Optional[Any] = None,
                                ensure_folder: bool = True) -> str:
    if ensure_folder:
        _ensure_sp_folder_tree(sp_ctx, folder_path)

    with open(local_path, "rb") as fh:
        content = fh.read()

    return _upload_bytes_to_sharepoint(
        content,
        filename,
        folder_path,
        file_type=file_type,
    )


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
            "Authorization":    f"Bearer {access_token}",
            "Content-Type":     "application/json",
            "apply-automations": "full",
        },
    )

    try:
        body = resp.json()
    except Exception:
        body = {"text": resp.text}

    result = {
        "status_code":  resp.status_code,
        "ok":           resp.ok,
        "url":          url,
        "sent_payload": payload,
        "response":     body,
    }
    return result


# ---------------------------------------------------------
# Clear existing files on the upload task before re-uploading.
# The .../tasks/{id}/files/ collection forbids GET, so files are read from the
# task detail. Called ONCE per delivery (before the per-RB upload loop),
# because every RB uploads to the same task; clearing per-upload would delete
# files uploaded by earlier RBs in the same run.
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


# =========================================================
# NAME FORMATTING
# =========================================================

def _title_case_name(name):
    """
    Title-case a person's name for display on the generated report.

    Applicant names arrive in inconsistent casing (all caps, all lowercase,
    etc.), so this normalizes to a conventional title case while keeping
    hyphenated names (e.g. "Smith-Jones") and names with apostrophes
    (e.g. "O'Brien") capitalized correctly on each segment, instead of
    naively lowercasing everything after the first letter the way
    str.title() alone would with something like "MARY-JANE O'BRIEN".
    """
    if not name or not isinstance(name, str):
        return name

    def _cap_segment(segment):
        return segment[:1].upper() + segment[1:].lower() if segment else segment

    def _cap_word(word):
        # Preserve hyphens/apostrophes as separators, capitalize each piece.
        return re.sub(r"[^\-']+", lambda m: _cap_segment(m.group(0)), word)

    return " ".join(_cap_word(w) for w in name.split(" "))


# =========================================================
# ADDITIONAL_CONTENT PARSING
# ---------------------------------------------------------
# additional_content is built upstream by inserting free-text values into a
# JSON string. Reviewer notes can contain unescaped double quotes or raw line
# breaks, which makes the JSON invalid. Parsing is therefore attempted in
# stages: a normal (lenient) parse first, then a retry after repairing
# unescaped inner quotes. The method that succeeded is recorded in
# dbg["parse_method"] so malformed upstream payloads remain visible.
#
# Only personal_info and application_info are used from this payload; all
# other sections are rebuilt from the SMApply reviews endpoint.
# =========================================================

_JSON_CLOSE_FOLLOWERS = set(",}]:")


def _repair_unescaped_quotes(s: str) -> str:
    """Escape double quotes that appear inside a JSON string value.

    A quote inside a string is treated as the real closing quote only when
    the next non-whitespace character is one of , } ] : (or end of input).
    Any other quote is escaped. This is a best-effort repair used only
    after a normal parse has failed.

    Known limitation: a quote inside a value that is immediately followed by
    a comma, colon or closing bracket (e.g. 'the term "x", which') will still
    be read as a closing quote, and the retry will fail with a clear error.
    """
    out = []
    in_string = False
    escaped = False
    n = len(s)
    for i, ch in enumerate(s):
        if not in_string:
            out.append(ch)
            if ch == '"':
                in_string = True
            continue
        if escaped:
            out.append(ch)
            escaped = False
            continue
        if ch == "\\":
            out.append(ch)
            escaped = True
            continue
        if ch == '"':
            j = i + 1
            while j < n and s[j] in " \t\r\n":
                j += 1
            if j >= n or s[j] in _JSON_CLOSE_FOLLOWERS:
                out.append(ch)
                in_string = False
            else:
                out.append('\\"')
            continue
        out.append(ch)
    return "".join(out)


def _loads_lenient(text: str) -> Any:
    """json.loads with strict=False (allows raw control characters such as
    line breaks inside strings) and support for double-encoded JSON."""
    raw = json.loads(text, strict=False)
    if isinstance(raw, str):
        raw = json.loads(raw, strict=False)
    return raw


def parse_additional_content(additional_content):
    # The raw payload contains applicant personal details, so only its type
    # and length are recorded here (not a text preview), to keep PII out of
    # the Windmill job logs.
    dbg = {
        "received_additional_content_type": str(type(additional_content)),
        "received_additional_content_length": (
            len(additional_content) if isinstance(additional_content, str) else None
        ),
    }
    if not additional_content or not isinstance(additional_content, str):
        raise ValueError("additional_content must be a non-empty JSON string.")

    s = additional_content.strip()
    candidates = [("as_is", s)]
    if len(s) >= 2 and s[0] == s[-1] and s[0] in "\"'":
        candidates.append(("outer_quotes_stripped", s[1:-1]))

    first_err = None
    parsed = False
    raw = None
    for label, cand in candidates:
        for mode, text in (("plain", cand),
                           ("quote_repaired", _repair_unescaped_quotes(cand))):
            try:
                raw = _loads_lenient(text)
                dbg["parse_method"] = f"{label}/{mode}"
                parsed = True
                break
            except json.JSONDecodeError as e:
                if first_err is None:
                    first_err = e
        if parsed:
            break

    if not parsed:
        raise ValueError(
            f"additional_content is not valid JSON "
            f"({first_err.msg} at char {first_err.pos}) and could not be repaired."
        ) from first_err

    if not isinstance(raw, dict):
        raise ValueError(f"Parsed additional_content is not a dict. Got: {type(raw)}")

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
# SMAPPLY REVIEWS -> PAYLOAD ADAPTER
# ---------------------------------------------------------
# Replaces the applicant-submitted `sa_` fields with the REVIEWED values
# pulled from the SMApply reviews endpoint. Returns a `payload` dict in the
# SAME shape the rest of the script already consumes, so no rendering code
# changes.
#
# NOT available from reviews (kept from additional_content):
#   - personal_info   (A1: name, dob, address, phone, email)
#   - application_info (regulatory_body_rn / _lpn selection)
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
# Confirm each map against pages A2 of a known-good report.
# -----------------------------------------------------------------
LIFETIME_FLAG_MAP   = {0: "Yes", 1: "No"}   # SMApply choice order [Yes, No] -> index 1 = No
DISCIPLINE_FLAG_MAP = {0: "Yes", 1: "No"}   # SMApply choice order [Yes, No] -> index 1 = No

# -----------------------------------------------------------------
# CONFIRMED per-task "Meets Requirements" Document Status index, provided
# directly (not inferred): each task has its own choice ordering, so a single
# global accepted-status set (the earlier TM-derived assumption) was wrong.
#   Identity Document Review     -> Document Status == 2 means Meets Requirements
#   Registration Document Review -> Document Status == 1 means Meets Requirements
#   Education Document Review    -> Document Status == 1 means Meets Requirements
# Only documents at the "Meets Requirements" index for their task are kept;
# everything else (including blocks with no Document Status at all) is
# dropped. This is in addition to the "Add another?" gating and the
# board/number or eca/recv presence checks.
# -----------------------------------------------------------------
FILTER_BY_STATUS = True
IDENTITY_ACCEPTED_STATUS     = {2}
REGISTRATION_ACCEPTED_STATUS = {1}
EDUCATION_ACCEPTED_STATUS    = {1}

# -----------------------------------------------------------------
# CONFIRMED: full Category Entry choice order for Identity Document Review,
# provided directly (0-indexed):
#   0 Passport
#   1 Driver's License
#   2 Government-Issued Photo ID
#   3 Marriage Certificate
#   4 Birth Certificate
#   5 Legal Name Affidavit
#   6 Other
# -----------------------------------------------------------------
IDENTITY_CATEGORY_MAP = {
    0: "Passport",
    1: "Driver's License",
    2: "Government-Issued Photo ID",
    3: "Marriage Certificate",
    4: "Birth Certificate",
    5: "Legal Name Affidavit",
    6: "Other",
}


def _map_choice(v: Any, m: Dict[int, str]) -> str:
    if v is None or v == "":
        return ""
    if isinstance(v, int):
        return m.get(v, f"[unmapped category index {v}]")
    return safe_str(v)


def _rev_headers():
    token = wmill.get_variable(SMAPPLY_REVIEWS_TOKEN_VAR)
    return {"Authorization": f"Bearer {token}", "Accept": "application/json"}


def fetch_all_reviews(application_id: int) -> List[Dict[str, Any]]:
    """All reviews for an application, following DRF pagination via `next`."""
    url = f"{SMAPPLY_BASE}/reviews/"
    params: Optional[dict] = {"application": application_id}
    out: List[Dict[str, Any]] = []
    while url:
        resp = requests.get(url, headers=_rev_headers(), params=params, timeout=120)
        if not resp.ok:
            raise RuntimeError(
                f"SMApply reviews fetch failed {resp.status_code} "
                f"app={application_id}: {resp.text[:2000]}"
            )
        data = resp.json()
        out.extend(data.get("results", []))
        url = data.get("next")   # absolute URL or None
        params = None            # `next` already carries the query
    return out


def _nl(s: Any) -> str:
    return re.sub(r"\s+", " ", (s or "").strip()).lower()


def _fmt_review_date(v: Any) -> str:
    """'2026-06-30T00:00:00' -> '2026/06/30' to match the report's date style."""
    s = safe_str(v)
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", s)
    return f"{m.group(1)}/{m.group(2)}/{m.group(3)}" if m else s


def _split_blocks(responses: List[dict], sentinel_prefix: str) -> List[Tuple[List[Tuple[str, Any]], Any]]:
    """Segment a task's flat responses list into per-document blocks on the
    'Add another ...' delimiter. Returns (block_fields, terminator_value) per
    block, where terminator_value is the answer to the 'Add another ...?'
    question that closed the block (None for the final, unterminated block)."""
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


def _status_accepted(bm: Dict[str, Any], accepted_status: Set[int]) -> bool:
    """True only if Document Status exactly matches the "Meets Requirements"
    index for this task. Missing/blank status is treated as NOT accepted --
    the requirement is to render only "Meets Requirements" docs, so an
    unreviewed or incomplete block should not pass through by default."""
    if not FILTER_BY_STATUS:
        return True
    v = _get(bm, "Document Status")
    if v is None or v == "":
        return False
    try:
        return int(v) in accepted_status
    except (TypeError, ValueError):
        return False


def _is_stop(add_another: Any) -> bool:
    """'Add another ...?' answered No -> the block that immediately follows is
    dropped. No is choice index 1 ([Yes, No]); also tolerate string forms.
    NOTE: this gates only the NEXT block, not the rest of the list. A later
    Yes re-enables ingestion (skip-and-resume)."""
    if add_another is None or add_another == "":
        return False
    try:
        return int(add_another) == 1
    except (TypeError, ValueError):
        return _nl(add_another) in ("no", "n", "false")


def _is_completed_review(rv: Dict[str, Any]) -> bool:
    """True only for reviews that have actually been completed. An
    in-progress or not-started review (no completed_at) must never be
    selected as the authoritative review for a task, even if every field
    inside it happens to be blank rather than raising an error -- a blank
    in-progress review and a genuinely-empty completed review look the same
    at the field level, so completion status is the only reliable signal."""
    return bool(safe_str(rv.get("completed_at")))


def _reviews_by_task(reviews: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    """Task name (normalized) -> the review that should represent that task.

    Only completed reviews are eligible. When more than one reviewer has
    completed the same task (e.g. two evaluators both reviewed Registration
    Document Review), the most recently completed one wins, rather than
    whichever the API happens to list first."""
    out: Dict[str, Dict[str, Any]] = {}
    for rv in reviews:
        if not _is_completed_review(rv):
            continue
        name = _nl((rv.get("task") or {}).get("name"))
        existing = out.get(name)
        if existing is None or safe_str(rv.get("completed_at")) > safe_str(existing.get("completed_at")):
            out[name] = rv
    return out


def build_payload_from_reviews(application_id: int,
                               base_payload: Dict[str, Any]) -> Tuple[Dict[str, Any], Dict[str, Any]]:
    """Build the render `payload` from SMApply reviews.

    `base_payload` is the parsed additional_content, used ONLY for
    personal_info and application_info (which reviews do not carry).

    Two dependent-field fixes are applied below, both scoped to the render
    payload only (they never write back to SMApply, so the raw review data
    is untouched - only what appears on the generated report changes):

      1. "Date registration/license expired/expires" is blanked whenever
         "Lifetime practice registration/license" resolves to "Yes". If a
         reviewer had entered an expiry date, then later flipped the
         lifetime flag back to Yes without clearing the date field in
         SMApply, the stale expiry date would otherwise still render next
         to "Yes" on the report.

      2. "Language of examination" is blanked whenever "Method by which
         this applicant was registered" is anything other than
         "Examination" - it only makes sense in that case, but was
         previously rendered unconditionally from whatever the field
         happened to hold.

    Returns (payload, review_debug).
    """
    reviews = fetch_all_reviews(application_id)
    by = _reviews_by_task(reviews)

    payload: Dict[str, Any] = {
        "application_info": base_payload.get("application_info") or {},
        "personal_info":    base_payload.get("personal_info") or {},
        "registration":     {},
        "assessment_notes": {},
        "part_c_sections":  [],
    }

    # ---- Registration Document Review -> A2 (registration.entry_N) + C3 ----
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
            if not _status_accepted(bm, REGISTRATION_ACCEPTED_STATUS):  # keep only Meets Requirements
                continue

            # Fix 1: lifetime flag gates expiry date.
            lifetime_flag = _map_yesno(_get(bm, "Lifetime practice registration/license"), LIFETIME_FLAG_MAP)
            expiry_date   = _fmt_review_date(_get(bm, "Date license/registration expired/expires"))
            if lifetime_flag == "Yes":
                expiry_date = ""

            # Fix 2: registration method gates exam language.
            registration_method = safe_str(_get(bm, "Method by which this applicant was registered"))
            exam_language        = safe_str(_get(bm, "Language of examination"))
            if registration_method.strip().lower() != "examination":
                exam_language = ""

            i += 1
            payload["registration"][f"entry_{i}"] = {
                f"sa_jurisdiction_name_{i}":            board,
                f"sa_jurisdiction_address_{i}":         safe_str(_get(bm, "Address of Registration Board/Authority")),
                f"sa_registrant_name_issued_{i}":       safe_str(_get(bm, "Applicant Name when registered/licensed")),
                f"sa_registration_title_{i}":           safe_str(_get(bm, "Title of registration/license")),
                f"sa_registration_number_{i}":          num,
                f"sa_registration_method_{i}":          registration_method,
                f"sa_exam_language_{i}":                exam_language,
                f"sa_registration_issued_date_{i}":     _fmt_review_date(_get(bm, "Date registration/license was issued")),
                f"sa_registration_lifetime_flag_{i}":   lifetime_flag,
                f"sa_registration_expiry_date_{i}":     expiry_date,
                f"sa_registration_status_{i}":          safe_str(_get(bm, "Status of license/registration")),
                f"sa_registration_discipline_flag_{i}": _map_yesno(_get(bm, "Has the license/registration ever been suspended"), DISCIPLINE_FLAG_MAP),
            }
            recv = _fmt_review_date(_get(bm, "Document Received Date"))
            line = board or ""
            if num:  line = f"{line}: {num}" if line else num
            if recv: line = f"{line} - {recv}" if line else recv
            if line: c3_bullets.append(line)

    # ---- Identity Document Review -> C1 ----
    id_rv = by.get(TASK_NAME_IDENTITY)
    c1_bullets: List[str] = []
    if id_rv:
        for block, _add in _split_blocks(id_rv.get("responses"), "add another identity"):
            bm = _bmap(block)
            dtype = _map_choice(_get(bm, "Category Entry"), IDENTITY_CATEGORY_MAP)
            dnum  = safe_str(_get(bm, "Identity Document Number"))
            recv  = _fmt_review_date(_get(bm, "Document Received Date"))
            if not (dtype or dnum):          # empty / abandoned block
                continue
            if not _status_accepted(bm, IDENTITY_ACCEPTED_STATUS):  # keep only Meets Requirements
                continue
            parts = [p for p in [dtype, dnum, recv] if p]
            if parts:
                c1_bullets.append(" - ".join(parts))

    # ---- Education Document Review -> C4 ----
    ed_rv = by.get(TASK_NAME_EDUCATION)
    c4_bullets: List[str] = []
    if ed_rv:
        prev_allows = True   # first block is always eligible
        for block, add_another in _split_blocks(ed_rv.get("responses"), "add another education"):
            allowed = prev_allows
            prev_allows = not _is_stop(add_another)   # gates ONLY the next block
            if not allowed:
                continue

            bm = _bmap(block)
            eca  = safe_str(_get(bm, "ECA ID"))
            recv = _fmt_review_date(_get(bm, "Document Received Date"))
            if not (eca or recv):        # empty / abandoned block
                continue
            if not _status_accepted(bm, EDUCATION_ACCEPTED_STATUS):  # keep only Meets Requirements
                continue

            line = "ECA Report"
            if eca:  line = f"{line} ({eca})"
            if recv: line = f"{line} - {recv}"
            c4_bullets.append(line)

    if c1_bullets:
        payload["part_c_sections"].append({"header": "C1: Identity Documents", "bullets": c1_bullets})
    if c3_bullets:
        payload["part_c_sections"].append({"header": "C3: Nursing Registration/License Documents", "bullets": c3_bullets})
    if c4_bullets:
        payload["part_c_sections"].append({"header": "C4: Other Documents", "bullets": c4_bullets})

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

    def _selected_review_info(rv: Optional[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
        if rv is None:
            return None
        reviewer = rv.get("reviewer") or {}
        name = " ".join(
            p for p in [safe_str(reviewer.get("first_name")), safe_str(reviewer.get("last_name"))] if p
        ).strip()
        return {
            "reviewer":     name,
            "completed_at": safe_str(rv.get("completed_at")),
        }

    review_dbg = {
        "reviews_fetched":            len(reviews),
        "review_task_names":          [ (rv.get("task") or {}).get("name") for rv in reviews ],
        "completed_review_count":     sum(1 for rv in reviews if _is_completed_review(rv)),
        "matched_registration":       reg_rv is not None,
        "matched_identity":           id_rv is not None,
        "matched_education":          ed_rv is not None,
        "matched_notes":              n_rv is not None,
        "selected_registration_review": _selected_review_info(reg_rv),
        "selected_identity_review":     _selected_review_info(id_rv),
        "selected_education_review":    _selected_review_info(ed_rv),
        "selected_notes_review":        _selected_review_info(n_rv),
        "registration_entries":       len(payload["registration"]),
        "part_c_headers":             [s.get("header") for s in payload["part_c_sections"]],
    }
    return payload, review_dbg


# =========================================================
# APPLICATION INFO / REGULATORY BODY LOOKUP
# =========================================================

NURSING_GROUP_DISPLAY = {
    "RN":  "Registered Nurse (RN)",
    "LPN": "Licensed Practical Nurse (LPN)",
}

RB_PROVINCE_MAP = {
    "NL College of Nurses":                                         "Newfoundland and Labrador",
    "PEI College of Nursing & Midwifery":                           "Prince Edward Island",
    "College of Nursing of New Brunswick":                          "New Brunswick",
    "Nova Scotia College of Nursing":                               "Nova Scotia",
    "College of Nurses of Ontario":                                 "Ontario",
    "College of Registered Nurses of Manitoba":                     "Manitoba",
    "College of Registered Nurses of Saskatchewan":                 "Saskatchewan",
    "College of Registered Nurses of Alberta":                      "Alberta",
    "BC College of Nurses and Midwives":                            "British Columbia",
    "Association of New Brunswick Licensed Practical Nurses":       "New Brunswick",
    "College of Licensed Practical Nurses of Manitoba":             "Manitoba",
    "College of Licensed Practical Nurses of Saskatchewan":         "Saskatchewan",
    "College of Licensed Practical Nurses and Health Care Aides of Alberta": "Alberta",
}

RB_SHORTNAME_MAP = {
    "NL College of Nurses":                                          "NLCN",
    "PEI College of Nursing & Midwifery":                           "PEICNM",
    "College of Nursing of New Brunswick":                          "CNNB",
    "Nova Scotia College of Nursing":                               "NSCN",
    "College of Nurses of Ontario":                                 "CNO",
    "College of Registered Nurses of Manitoba":                     "CRNM",
    "College of Registered Nurses of Saskatchewan":                 "CRNS",
    "College of Registered Nurses of Alberta":                      "CRNA",
    "BC College of Nurses and Midwives":                            "BCCNM",
    "Association of New Brunswick Licensed Practical Nurses":       "ANBLPN",
    "College of Licensed Practical Nurses of Manitoba":             "CLPNM",
    "College of Licensed Practical Nurses of Saskatchewan":         "CLPNS",
    "College of Licensed Practical Nurses and Health Care Aides of Alberta": "CLHA",
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
        p for p in [_title_case_name(safe_str(pi.get("sa_first_name"))),
                    _title_case_name(safe_str(pi.get("sa_middle_name"))),
                    _title_case_name(safe_str(pi.get("sa_last_name")))] if p
    ).strip()
    same_flag    = normalize_yes_no(pi.get("sa_mailing_same_as_current_flag"))
    current_addr = join_address_from_flat("sa_current_", pi)
    mailing_addr = join_address_from_flat("sa_mailing_", pi)
    return {
        "name":                      full_name,
        "other_names":               _title_case_name(safe_str(pi.get("sa_other_names"))),
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
            "registrant_name_issued":       _title_case_name(safe_str(e.get(f"sa_registrant_name_issued_{i}"))),
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
    rows = [
        ("Name of jurisdiction / registering board / authority:",                reg.get("jurisdiction_name", "")),
        ("Current address of jurisdiction / registering board / authority:",     reg.get("jurisdiction_address", "")),
        ("Applicant name when registration / license / diploma was issued:",     reg.get("registrant_name_issued", "")),
        ("Title of registration / license:",                                     reg.get("registration_title", "")),
        ("Registration / license number:",                                       reg.get("registration_number", "")),
        ("Method by which this applicant was registered:",                       reg.get("registration_method", "")),
    ]
    # "Language of examination" only applies when the registration method was
    # Examination - the row is omitted entirely (not just blanked) when it
    # doesn't apply, so a stale value from an earlier method selection never
    # surfaces on the report.
    if reg.get("registration_method", "").strip().lower() == "examination":
        rows.append(("Language of examination:", reg.get("exam_language", "")))
    rows += [
        ("Date registration / license issued:",                                  reg.get("registration_issued_date", "")),
        ("Is this a lifetime practice registration / license:",                  reg.get("registration_lifetime_flag", "")),
        ("Date registration / license expired / expires:",                       reg.get("registration_expiry_date", "")),
        ("Status of registration / license:",                                    reg.get("registration_status", "")),
        ("Has the applicant's registration / license ever been suspended, "
         "restricted, surrendered, revoked, or subject to individual terms "
         "and conditions:",                                                       reg.get("registration_discipline_flag", "")),
    ]
    return rows


def build_part_c_sections(payload: Dict[str, Any]) -> List[Dict[str, Any]]:
    sections = payload.get("part_c_sections")
    if isinstance(sections, list) and sections:
        return sections
    sections = []
    csecs    = payload.get("c_sections") or {}
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
    notes               = payload.get("assessment_notes") or {}
    base_public         = safe_str(notes.get("public"))
    identity_note       = safe_str(notes.get("identity_section_note"))
    education_note      = safe_str(notes.get("education_section_note"))
    registration_note   = safe_str(notes.get("registration_section_note"))

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
# PDF - PAGE 2: PERSONAL INFO
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
# PDF - PAGE 3+: REGISTRATION
# =========================================================

def render_registration_tables(writer, registrations, *, y_top):
    styles      = getSampleStyleSheet()
    label_style = ParagraphStyle("label", parent=styles["Normal"],
                                 fontName="Helvetica-Bold", fontSize=FONT_LABEL,
                                 textColor=COLOR_RED, leading=12)
    value_style = ParagraphStyle("value", parent=styles["Normal"],
                                 fontName="Helvetica", fontSize=FONT_BODY,
                                 textColor=COLOR_TEXT, leading=12)
    page_index = PAGE_A2
    for i, reg in enumerate(registrations):
        if i == 0:
            # first registration reuses the template's existing A2 page
            if page_index >= len(writer.pages):
                writer.insert_blank_page(width=612, height=792, index=page_index)
        else:
            # every subsequent registration ALWAYS gets a brand-new page,
            # pushing Part B / evaluator / Part C pages down instead of
            # overwriting them
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
                 evaluator_name: str = "", date_str: str = "") -> PdfWriter:
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

    regs = extract_registration_entries(payload)
    if regs:
        render_registration_tables(writer, regs, y_top=REGISTRATION_HEADER_Y - HALF_INCH)

    evaluator_page_index = len(writer.pages) - 2
    if evaluator_page_index >= 0 and (evaluator_name or date_str):
        page_ev  = writer.pages[evaluator_page_index]
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
                color: RGBColor = DOCX_TEXT, size: int = 10):
    cell.text = ""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    _cell_margins(cell)
    lines = str(text).split("\n") if text else [""]
    for idx, line in enumerate(lines):
        p = cell.paragraphs[0] if idx == 0 else cell.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)
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
          "This report does not determine an applicant's eligibility to become registered or "
          "licensed in the jurisdiction of application. A regulatory body reserves the right to "
          "request additional information pertaining to the findings provided in this report and "
          "has the authority to require further testing and assessment, in accordance with the "
          "regulatory body's legislation.",
          size=9, color=DOCX_TEXT, space_after=6)

    _para(doc,
          "Disclaimer: The purpose of this Advisory Report is to support the Canadian nursing "
          "registration decisions of NNAS member nursing regulatory bodies. NNAS does not intend "
          "it for any other use nor does it endorse it for any other use. NNAS is not responsible "
          "or liable for anything that results from using this Report for any other purpose or for "
          "the use of the Report by a person who is not an NNAS member nursing regulatory body.",
          size=9, color=DOCX_RED, space_after=16)

    notes               = payload.get("assessment_notes") or {}
    base_public         = safe_str(notes.get("public"))
    identity_note       = safe_str(notes.get("identity_section_note"))
    education_note      = safe_str(notes.get("education_section_note"))
    registration_note   = safe_str(notes.get("registration_section_note"))

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
        hp      = hc.paragraphs[0]
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
        _para(doc, "A2: REGISTRATION/LICENSE OVERVIEW",
              bold=True, size=14, color=DOCX_RED,
              align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
        rows = registration_rows(reg)
        _build_two_col_table(doc, rows, all_grey=True)


def docx_page4_part_b(doc: Document):
    _banner(doc, "Part B: Nursing Education Analysis")


def docx_page5_evaluator(doc: Document, evaluator_name: str = "", date_str: str = ""):
    body_text = (
        "This report has been generated based on the compilation of applicant's data and the "
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
        "This report does not determine an applicant's eligibility to become registered or "
        "licensed in the jurisdiction of application. Regardless of the report content, a "
        "regulatory body reserves the right to request additional information pertaining to the "
        "findings provided in this report and has the authority to require further testing and "
        "assessment, in accordance with the regulatory body's legislation."
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
               evaluator_name: str = "", date_str: str = "") -> str:
    doc = Document()
    for sec in doc.sections:
        sec.top_margin    = Cm(1.8)
        sec.bottom_margin = Cm(1.8)
        sec.left_margin   = Cm(1.8)
        sec.right_margin  = Cm(1.8)

    docx_page1_cover(doc, payload, app_info, logo_path)

    _page_break(doc)
    docx_page2_personal_info(doc, payload)

    regs = extract_registration_entries(payload)
    if regs:
        _page_break(doc)
        docx_page3_registration(doc, payload)

    _page_break(doc)
    docx_page4_part_b(doc)

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
         nnas_id: Optional[int] = None):

    # additional_content still supplies personal_info + application_info;
    # education / registration / identity / assessment-notes now come from
    # the SMApply reviews endpoint.
    app_id             = coerce_appid(appid)
    base_payload, dbg  = parse_additional_content(additional_content)
    payload, rev_dbg   = build_payload_from_reviews(app_id, base_payload)
    dbg["reviews"]     = rev_dbg

    application_number = parse_application_number(application_title)
    nnas_id_str        = str(nnas_id) if nnas_id is not None else ""
    evaluator_name     = "NNAS Expedited Service"
    date_str           = datetime.utcnow().strftime("%Y-%m-%d")

    if not reference_id:
        reference_id = safe_str(payload.get("reference_id")) or str(app_id)

    app_info_list = build_application_info(
        payload,
        application_number=application_number,
        nnas_id=nnas_id_str
    )
    if not app_info_list:
        raise ValueError("No regulatory bodies found in application_info.")

    part_c = payload.get("part_c_sections") or []

    dbg.update({
        "appid_type":           str(type(appid)),
        "appid_used":           app_id,
        "reference_id_used":    reference_id,
        "application_title":    application_title,
        "application_number":   application_number,
        "nnas_id":              nnas_id_str,
        "rb_count":             len(app_info_list),
        "rb_list":              [a["regulatory_body"] for a in app_info_list],
        "province_list":        [a["province"]        for a in app_info_list],
        "has_assessment_notes": isinstance(payload.get("assessment_notes"), dict)
                                and bool(payload.get("assessment_notes")),
        "has_personal_info":    isinstance(payload.get("personal_info"), dict)
                                and bool(payload.get("personal_info")),
        "has_registration":     isinstance(payload.get("registration"), dict)
                                and bool(payload.get("registration")),
        "pi_keys_sample":       sorted(list((payload.get("personal_info") or {}).keys()))[:12],
        "reg_entries_found":    len(extract_registration_entries(payload)),
        "part_c_headers":       [safe_str(s.get("header")) for s in part_c],
        "part_c_bullet_counts": [len(s.get("bullets") or []) for s in part_c],
    })

    template  = download_template()
    logo_path = download_logo()
    dbg["logo_downloaded"] = logo_path is not None

    ts = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    os.makedirs("tmp", exist_ok=True)

    application_folder_id = reference_id

    reports = []
    _smapply_lock = threading.Lock()

    def _process_rb(idx_app_info):
        idx, app_info = idx_app_info
        rb_full    = app_info.get("regulatory_body", "")
        shortname  = rb_short_name(rb_full)
        group_code = app_info.get("nursing_group_code", "")

        group_tag = f"_{group_code}" if group_code else ""
        base_name        = f"{reference_id}_{shortname}{group_tag}_{ts}"
        pdf_name_plain   = f"{base_name}.pdf"
        docx_name_plain  = f"{base_name}.docx"
        pdf_name_latest  = f"{base_name}_latest.pdf"
        docx_name_latest = f"{base_name}_latest.docx"

        pdf_path  = os.path.join("tmp", pdf_name_plain)
        rb_writer = populate_pdf(template, payload, app_info,
                                 evaluator_name=evaluator_name, date_str=date_str)
        with open(pdf_path, "wb") as fh:
            rb_writer.write(fh)
        docx_path = os.path.join("tmp", docx_name_plain)
        build_docx(payload, app_info, logo_path, docx_path,
                   evaluator_name=evaluator_name, date_str=date_str)

        folder_path = (
            f"{APPLICATIONS_BASE}/{application_folder_id}/"
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
                appid=app_id, task_id=SMAPPLY_TASK_ID,
                pdf_filepath=pdf_path, upload_filename=pdf_name_plain
            )
            sm_docx = upload_docx_to_smapply_task(
                appid=app_id, task_id=SMAPPLY_TASK_ID,
                docx_filepath=docx_path, upload_filename=docx_name_plain
            )

        return idx, {
            "regulatory_body":      rb_full,
            "short_name":           shortname,
            "province":             app_info.get("province", ""),
            "sharepoint_folder":    _sp_url(folder_path),
            "sharepoint_pdf":       sp_pdf,
            "sharepoint_docx":      sp_docx,
            "smapply_pdf":          sm_pdf,
            "smapply_docx":         sm_docx,
        }

    # Before uploading, clear any files already on the upload task so a
    # re-run replaces rather than accumulates. Done ONCE here, not inside
    # _process_rb, because every RB uploads to the same task and clearing
    # per-upload would delete earlier RBs' new files.
    dbg["smapply_files_cleared"] = clear_smapply_task_files(app_id, SMAPPLY_TASK_ID)

    workers = min(SP_MAX_WORKERS, len(app_info_list)) or 1
    indexed = list(enumerate(app_info_list))
    collected: List[Tuple[int, Dict[str, Any]]] = []
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = [pool.submit(_process_rb, item) for item in indexed]
        for fut in as_completed(futures):
            collected.append(fut.result())
    reports = [r for _, r in sorted(collected, key=lambda t: t[0])]

    smapply_task_completed = mark_smapply_task_complete(
        appid=app_id, task_id=SMAPPLY_TASK_ID
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