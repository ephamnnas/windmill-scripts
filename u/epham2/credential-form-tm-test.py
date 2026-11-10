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

# DOCX colour constants
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


# =========================================================
# SHAREPOINT CONFIG
# =========================================================

TENANT_ID     = wmill.get_variable("u/epham2/AZURE_TENANT_ID")
CLIENT_ID     = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_ID")
CLIENT_SECRET = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_SECRET")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SHAREPOINT_HOSTNAME = "nnasca.sharepoint.com"
SITE_PATH = "/sites/NNAS-Portal-Sharepoint"
DOCUMENT_LIBRARY_NAME = "Applications"

TEMPLATE_FOLDER   = "Applications/Credential form"
TEMPLATE_FILENAME = "CredentialReportBaseV1.pdf"
LOGO_FILENAME     = "Credential_header.png"   # logo lives in same folder as the PDF template

# Output destination: <APPLICATIONS_BASE>/<application id>/<CRED_REPORTS_SUBFOLDER>/<RB short name>/
APPLICATIONS_BASE      = "applications/tm/test"
CRED_REPORTS_SUBFOLDER = "Credential reports"

# Max concurrent per-RB workers (each does doc gen + SharePoint + SMApply uploads).
SP_MAX_WORKERS = 8

PAGE_NOTES = 0
PAGE_A1    = 1
PAGE_A2    = 2


# =========================================================
# SMAPPLY CONFIG
# =========================================================

# TM mailroomstorage task (file-upload task).
SMAPPLY_TASK_ID = 2075140


# =========================================================
# SHAREPOINT HELPERS
# =========================================================

_GRAPH_TOKEN = None
_SP_SITE_ID = None
_SP_DRIVE_ID = None
_SP_LIST_ID = None
_GRAPH_LOCK = threading.Lock()


def ctx():
    return None


def _site_web_url() -> str:
    return f"https://{SHAREPOINT_HOSTNAME}{SITE_PATH}"


def _get_graph_token() -> str:
    global _GRAPH_TOKEN

    if _GRAPH_TOKEN:
        return _GRAPH_TOKEN

    with _GRAPH_LOCK:
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
    token = _get_graph_token()

    response = requests.get(
        f"{GRAPH_BASE}{endpoint}",
        headers={"Authorization": f"Bearer {token}"},
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


def _encode_graph_path(path: str) -> str:
    return quote(path.strip("/"), safe="/")


def _init_sharepoint_graph() -> None:
    global _SP_SITE_ID, _SP_DRIVE_ID, _SP_LIST_ID

    if _SP_SITE_ID and _SP_DRIVE_ID and _SP_LIST_ID:
        return

    with _GRAPH_LOCK:
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

        lists = _graph_request(
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


def _drive_path(folder_path: str) -> str:
    path = (folder_path or "").strip().strip("/")
    library_name = DOCUMENT_LIBRARY_NAME.strip("/")

    if path.lower() == library_name.lower():
        return ""

    if path.lower().startswith(library_name.lower() + "/"):
        return path[len(library_name) + 1:]

    if path.lower().startswith("applications/"):
        return path[len("applications/"):]

    return path


def _sp_url(folder_path: str) -> str:
    drive_path = _drive_path(folder_path)
    if drive_path:
        return f"{_site_web_url()}/{DOCUMENT_LIBRARY_NAME}/{drive_path}"
    return f"{_site_web_url()}/{DOCUMENT_LIBRARY_NAME}"


def _get_drive_item_by_path(drive_relative_path: str):
    _init_sharepoint_graph()

    clean_path = _drive_path(drive_relative_path)

    if not clean_path:
        return _graph_request(
            "GET",
            f"/drives/{_SP_DRIVE_ID}/root",
            expected_statuses=(200,),
        )

    return _graph_get_or_none(
        f"/drives/{_SP_DRIVE_ID}/root:/{_encode_graph_path(clean_path)}"
    )


def _download_drive_item_content(item_id: str) -> bytes:
    _init_sharepoint_graph()

    response = requests.get(
        f"{GRAPH_BASE}/drives/{_SP_DRIVE_ID}/items/{item_id}/content",
        headers={"Authorization": f"Bearer {_get_graph_token()}"},
        timeout=300,
        allow_redirects=True,
    )

    if response.status_code != 200:
        raise RuntimeError(
            f"Failed downloading SharePoint drive item.\n"
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

    list_item_id = list_item["id"]

    _graph_request(
        "PATCH",
        f"/sites/{_SP_SITE_ID}/lists/{_SP_LIST_ID}/items/{list_item_id}/fields",
        headers={"Content-Type": "application/json"},
        json=fields,
        expected_statuses=(200, 204),
    )


def _create_folder(parent_path: str, folder_name: str):
    _init_sharepoint_graph()

    parent_path = _drive_path(parent_path)
    full_path = f"{parent_path}/{folder_name}".strip("/")

    existing = _get_drive_item_by_path(full_path)
    if existing:
        return existing

    if parent_path:
        endpoint = f"/drives/{_SP_DRIVE_ID}/root:/{_encode_graph_path(parent_path)}:/children"
    else:
        endpoint = f"/drives/{_SP_DRIVE_ID}/root/children"

    return _graph_request(
        "POST",
        endpoint,
        headers={"Content-Type": "application/json"},
        json={
            "name": folder_name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "fail",
        },
        expected_statuses=(201,),
    )


def download_template() -> str:
    drive_path = f"{_drive_path(TEMPLATE_FOLDER)}/{TEMPLATE_FILENAME}".strip("/")
    item = _get_drive_item_by_path(drive_path)
    if not item:
        raise RuntimeError(f"SharePoint template not found: {drive_path}")

    os.makedirs("tmp", exist_ok=True)
    path = os.path.join("tmp", TEMPLATE_FILENAME)

    with open(path, "wb") as f:
        f.write(_download_drive_item_content(item["id"]))

    if os.path.getsize(path) == 0:
        raise RuntimeError(f"Downloaded template is empty: {path}")

    return path


def download_logo() -> Optional[str]:
    try:
        drive_path = f"{_drive_path(TEMPLATE_FOLDER)}/{LOGO_FILENAME}".strip("/")
        item = _get_drive_item_by_path(drive_path)
        if not item:
            return None

        os.makedirs("tmp", exist_ok=True)
        path = os.path.join("tmp", LOGO_FILENAME)

        with open(path, "wb") as f:
            f.write(_download_drive_item_content(item["id"]))

        if os.path.getsize(path) == 0:
            return None

        return path
    except Exception:
        return None


def _ensure_sp_folder_tree(sp_ctx, folder_path: str):
    drive_folder_path = _drive_path(folder_path)
    current = ""

    for part in [p for p in drive_folder_path.split("/") if p]:
        _create_folder(current, part)
        current = f"{current}/{part}".strip("/")


def _demote_latest_in_folder(sp_ctx, folder_path: str,
                             name_filter: Optional[str] = None):
    drive_folder_path = _drive_path(folder_path)

    folder_item = _get_drive_item_by_path(drive_folder_path)
    if not folder_item:
        return

    try:
        children = _graph_request(
            "GET",
            f"/drives/{_SP_DRIVE_ID}/items/{folder_item['id']}/children",
            expected_statuses=(200,),
        )["value"]
    except Exception:
        return

    for f in list(children):
        try:
            name = f.get("name", "")
            item_id = f.get("id")
        except Exception:
            continue

        if not item_id or "." not in name:
            continue
        if name_filter is not None and name_filter not in name:
            continue

        stem, _, ext = name.rpartition(".")
        if not stem.endswith("_latest"):
            continue

        new_name = f"{stem[:-len('_latest')]}.{ext}"
        new_path = f"{drive_folder_path}/{new_name}".strip("/")

        try:
            content = _download_drive_item_content(item_id)
            demoted = _graph_request(
                "PUT",
                f"/drives/{_SP_DRIVE_ID}/root:/{_encode_graph_path(new_path)}:/content",
                headers={"Content-Type": "application/octet-stream"},
                data=content,
                expected_statuses=(200, 201),
            )
            _patch_driveitem_fields(demoted["id"], {"FileType": "Report"})
        except Exception:
            continue

        try:
            _graph_request(
                "DELETE",
                f"/drives/{_SP_DRIVE_ID}/items/{item_id}",
                expected_statuses=(204,),
            )
        except Exception:
            pass


def upload_to_sharepoint_folder(local_path: str, filename: str, folder_path: str,
                                file_type: str = "Report",
                                sp_ctx=None,
                                ensure_folder: bool = True) -> str:
    if ensure_folder:
        _ensure_sp_folder_tree(sp_ctx, folder_path)

    drive_folder_path = _drive_path(folder_path)
    folder_item = _get_drive_item_by_path(drive_folder_path)
    if not folder_item:
        raise RuntimeError(f"SharePoint folder not found: {drive_folder_path}")

    with open(local_path, "rb") as fh:
        content = fh.read()

    upload_path = f"{drive_folder_path}/{filename}".strip("/")

    uploaded_file = _graph_request(
        "PUT",
        f"/drives/{_SP_DRIVE_ID}/root:/{_encode_graph_path(upload_path)}:/content",
        headers={"Content-Type": "application/octet-stream"},
        data=content,
        expected_statuses=(200, 201),
    )

    _patch_driveitem_fields(uploaded_file["id"], {"FileType": file_type})

    return uploaded_file.get("webUrl") or _sp_url(upload_path)



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
        # Surface WHAT the server rejected (raise_for_status discards the body).
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


# SMApply task status codes: 1 = Not started, 2 = In progress, 3 = Completed.
SMAPPLY_STATUS_COMPLETED = 3


def mark_smapply_task_complete(*, appid, task_id) -> Dict[str, Any]:
    """PATCH a SMApply task to status 3 (Completed).

    The files are already attached to the task by the upload POSTs, so completing
    the task is just the status change. We send ONLY {"status": 3}: the endpoint
    rejects `data` ('invalid property') and `completed_at` ('readonly property').
    Sends 'apply-automations: full' so completion automations fire. Requires the
    token's 'View Applications' permission.
    """
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

    # Parse the response body regardless of status (this is where the API explains
    # WHAT it rejected on a 400 — raise_for_status would otherwise discard it).
    try:
        body = resp.json()
    except Exception:
        body = {"text": resp.text}

    result = {
        "status_code":  resp.status_code,
        "ok":           resp.ok,
        "url":          url,
        "sent_payload": payload,   # exactly what we PATCHed, for side-by-side debugging
        "response":     body,      # server's reply, incl. validation errors on a 400
    }
    # Do NOT raise: returning the body lets the caller see the validation message.
    # Flip raise_on_error=True (below) once the shape is confirmed if you want hard failures.
    return result


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
    try:
        raw = json.loads(s)
        if isinstance(raw, str): raw = json.loads(raw)
    except json.JSONDecodeError:
        if (s.startswith('"') and s.endswith('"')) or (s.startswith("'") and s.endswith("'")):
            s2 = s[1:-1]
            raw = json.loads(s2)
            if isinstance(raw, str): raw = json.loads(raw)
        else:
            raise
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
# APPLICATION INFO / REGULATORY BODY LOOKUP
# =========================================================

NURSING_GROUP_DISPLAY = {
    "RN":  "Registered Nurse (RN)",
    "LPN": "Licensed Practical Nurse (LPN)",
}

RB_PROVINCE_MAP = {
    # RN
    "NL College of Nurses":                                         "Newfoundland and Labrador",
    "PEI College of Nursing & Midwifery":                           "Prince Edward Island",
    "College of Nursing of New Brunswick":                          "New Brunswick",
    "Nova Scotia College of Nursing":                               "Nova Scotia",
    "College of Nurses of Ontario":                                 "Ontario",
    "College of Registered Nurses of Manitoba":                     "Manitoba",
    "College of Registered Nurses of Saskatchewan":                 "Saskatchewan",
    "College of Registered Nurses of Alberta":                      "Alberta",
    "BC College of Nurses and Midwives":                            "British Columbia",
    # LPN
    "Association of New Brunswick Licensed Practical Nurses":       "New Brunswick",
    "College of Licensed Practical Nurses of Manitoba":             "Manitoba",
    "College of Licensed Practical Nurses of Saskatchewan":         "Saskatchewan",
    "College of Licensed Practical Nurses and Health Care Aides of Alberta": "Alberta",
}

# Full regulatory-body name -> SFTP/folder acronym (matches the sftp-* short names).
RB_SHORTNAME_MAP = {
    # RN
    "NL College of Nurses":                                          "NLCN",
    "PEI College of Nursing & Midwifery":                           "PEICNM",
    "College of Nursing of New Brunswick":                          "CNNB",
    "Nova Scotia College of Nursing":                               "NSCN",
    "College of Nurses of Ontario":                                 "CNO",
    "College of Registered Nurses of Manitoba":                     "CRNM",
    "College of Registered Nurses of Saskatchewan":                 "CRNS",
    "College of Registered Nurses of Alberta":                      "CRNA",
    "BC College of Nurses and Midwives":                            "BCCNM",
    # LPN (shared entries already covered above via same string)
    "Association of New Brunswick Licensed Practical Nurses":       "ANBLPN",
    "College of Licensed Practical Nurses of Manitoba":             "CLPNM",
    "College of Licensed Practical Nurses of Saskatchewan":         "CLPNS",
    "College of Licensed Practical Nurses and Health Care Aides of Alberta": "CLHA",
}


def rb_short_name(rb_full: str) -> str:
    """Resolve a full regulatory-body name to its folder acronym, with safe fallbacks."""
    rb_full = (rb_full or "").strip()
    if rb_full in RB_SHORTNAME_MAP:
        return RB_SHORTNAME_MAP[rb_full]
    for k, v in RB_SHORTNAME_MAP.items():
        if k.lower() == rb_full.lower():
            return v
    # Last resort: derive an acronym from word initials, sanitized for a folder name.
    acronym = "".join(w[0] for w in re.split(r"\s+", rb_full) if w and w[0].isalpha()).upper()
    acronym = re.sub(r"[^A-Za-z0-9_-]", "", acronym)
    return acronym or "UNKNOWN_RB"


def build_application_info(payload: Dict[str, Any],
                            application_number: str = "",
                            nnas_id: str = "") -> List[Dict[str, str]]:
    """Build one app-info entry per regulatory body.

    Both regulatory_body_rn and regulatory_body_lpn may be populated at once. We
    process ALL RN bodies first, then ALL LPN bodies, generating a distinct entry
    per body. Each entry carries its OWN nursing group (not the top-level
    nursing_group_selection), so its acronym and group label resolve correctly.
    """
    ai = payload.get("application_info") or {}
    results = []
    # (group code, raw comma-separated list of bodies) processed in this order.
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
        "first_language":            safe_str(pi.get("sa_first_language")),
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
        ("First language:",                                             d["first_language"]),
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
        ("Name of jurisdiction / registering board / authority:",                reg.get("jurisdiction_name", "")),
        ("Current address of jurisdiction / registering board / authority:",     reg.get("jurisdiction_address", "")),
        ("Applicant name when registration / license / diploma was issued:",     reg.get("registrant_name_issued", "")),
        ("Title of registration / license:",                                     reg.get("registration_title", "")),
        ("Registration / license number:",                                       reg.get("registration_number", "")),
        ("Method by which this applicant was registered:",                       reg.get("registration_method", "")),
        ("Language of examination:",                                             reg.get("exam_language", "")),
        ("Date registration / license issued:",                                  reg.get("registration_issued_date", "")),
        ("Is this a lifetime practice registration / license:",                  reg.get("registration_lifetime_flag", "")),
        ("Date registration / license expired / expires:",                       reg.get("registration_expiry_date", "")),
        ("Status of registration / license:",                                    reg.get("registration_status", "")),
        ("Has the applicant's registration / license ever been suspended, "
         "restricted, surrendered, revoked, or subject to individual terms "
         "and conditions:",                                                       reg.get("registration_discipline_flag", "")),
    ]


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
    for reg in registrations:
        while page_index >= len(writer.pages):
            writer.add_blank_page(width=612, height=792)
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

def populate_pdf(template_path, payload, app_info: Dict[str, str]) -> PdfWriter:
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


def docx_page5_evaluator(doc: Document):
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
    for i, label in enumerate(("Evaluator:", "Date:")):
        cell = ev_tbl.rows[0].cells[i]
        _set_cell_borders(cell, size=0, color="FFFFFF")
        _cell_margins(cell, top=40, bottom=40, left=0, right=0)
        cell.text = ""
        p   = cell.paragraphs[0]
        run = p.add_run(label)
        run.font.size      = Pt(10)
        run.font.color.rgb = DOCX_RED
        run.font.name      = "Calibri"
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
               logo_path: Optional[str], out_path: str) -> str:
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
    docx_page5_evaluator(doc)

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

    payload, dbg       = parse_additional_content(additional_content)
    app_id             = coerce_appid(appid)
    application_number = parse_application_number(application_title)
    nnas_id_str        = str(nnas_id) if nnas_id is not None else ""

    if not reference_id:
        reference_id = safe_str(payload.get("reference_id")) or str(app_id)

    app_info_list = build_application_info(
        payload,
        application_number=application_number,
        nnas_id=nnas_id_str
    )
    if not app_info_list:
        raise ValueError("No regulatory bodies found in application_info.")

    csecs = payload.get("c_sections") if isinstance(payload.get("c_sections"), dict) else {}
    c1    = csecs.get("c1_identity_documents") if isinstance(csecs.get("c1_identity_documents"), dict) else {}
    c1_entries_present = [k for k in ("entry_1", "entry_2") if isinstance(c1.get(k), dict)]
    c3    = csecs.get("c3_registration_documents") if isinstance(csecs.get("c3_registration_documents"), dict) else {}
    c3_entries_present = [
        f"entry_{i}" for i in range(1, 8)
        if isinstance(c3.get(f"entry_{i}"), dict)
        and any(safe_str((c3.get(f"entry_{i}") or {}).get(k))
                for k in ("registration_board_authority",
                          "registration_license_number",
                          "document_received_date"))
    ]

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
        "has_assessment_notes": isinstance(payload.get("assessment_notes"), dict),
        "has_personal_info":    isinstance(payload.get("personal_info"), dict),
        "has_registration":     isinstance(payload.get("registration"), dict),
        "has_c_sections":       isinstance(payload.get("c_sections"), dict),
        "pi_keys_sample":       sorted(list((payload.get("personal_info") or {}).keys()))[:12],
        "reg_keys_sample":      sorted(list((payload.get("registration") or {}).keys()))[:12],
        "reg_entries_found":    len(extract_registration_entries(payload)),
        "c1_entries_present":   c1_entries_present,
        "c3_entries_present":   c3_entries_present,
    })

    template  = download_template()
    logo_path = download_logo()
    dbg["logo_downloaded"] = logo_path is not None

    ts = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    os.makedirs("tmp", exist_ok=True)

    # Application folder keyed by reference_id (matches applications/tm/<reference_id>/...).
    application_folder_id = reference_id

    reports = []
    _smapply_lock = threading.Lock()  # serialize SMApply uploads (shared task = no concurrency)

    def _process_rb(idx_app_info):
        """Generate + upload one RB's reports."""
        idx, app_info = idx_app_info
        rb_full    = app_info.get("regulatory_body", "")
        shortname  = rb_short_name(rb_full)
        group_code = app_info.get("nursing_group_code", "")  # "RN" or "LPN"

        # Tag the group into the name so RN and LPN reports never overwrite each
        # other when they share an acronym folder (NLCN / NSCN / PEICNM). Both can
        # then carry _latest independently.
        group_tag = f"_{group_code}" if group_code else ""
        base_name        = f"{reference_id}_{shortname}{group_tag}_{ts}"
        pdf_name_plain   = f"{base_name}.pdf"
        docx_name_plain  = f"{base_name}.docx"
        pdf_name_latest  = f"{base_name}_latest.pdf"
        docx_name_latest = f"{base_name}_latest.docx"

        # --- Build this RB's PDF + DOCX (CPU; ~140ms total per RB) ---
        pdf_path  = os.path.join("tmp", pdf_name_plain)
        rb_writer = populate_pdf(template, payload, app_info)
        with open(pdf_path, "wb") as fh:
            rb_writer.write(fh)
        docx_path = os.path.join("tmp", docx_name_plain)
        build_docx(payload, app_info, logo_path, docx_path)

        # --- SharePoint: <app id>/Credential reports/<RB short name>/ ---
        folder_path = (
            f"{APPLICATIONS_BASE}/{application_folder_id}/"
            f"{CRED_REPORTS_SUBFOLDER}/{shortname}"
        )
        sp_ctx = ctx()
        _ensure_sp_folder_tree(sp_ctx, folder_path)
        # Only demote prior _latest files of the SAME group (RN won't touch LPN's
        # _latest in a shared acronym folder, and vice versa).
        _demote_latest_in_folder(sp_ctx, folder_path,
                                 name_filter=group_tag if group_tag else None)

        sp_pdf  = upload_to_sharepoint_folder(pdf_path,  pdf_name_latest,  folder_path,
                                              sp_ctx=sp_ctx, ensure_folder=False)
        sp_docx = upload_to_sharepoint_folder(docx_path, docx_name_latest, folder_path,
                                              sp_ctx=sp_ctx, ensure_folder=False)

        # --- SMApply: upload this RB's PDF + DOCX (clean names) ---
        # Serialized across threads: the SMApply task is one shared resource and
        # rejects concurrent uploads (400). SharePoint above stays parallel because
        # each RB writes to its own folder. These two POSTs are quick, so the lock
        # costs little while keeping uploads one-at-a-time as before.
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

    # Run RBs concurrently — the work is network-bound, so overlapping the round
    # trips is the dominant speedup. Results are reordered back to RN-then-LPN.
    workers = min(SP_MAX_WORKERS, len(app_info_list)) or 1
    indexed = list(enumerate(app_info_list))
    collected: List[Tuple[int, Dict[str, Any]]] = []
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = [pool.submit(_process_rb, item) for item in indexed]
        for fut in as_completed(futures):
            collected.append(fut.result())
    reports = [r for _, r in sorted(collected, key=lambda t: t[0])]

    # Mark the shared SMApply task complete ONCE, after all RB files are uploaded.
    # Files are already attached by the upload POSTs; this is just the status change.
    # NOTE: requires the token to have the 'View Applications' permission on SMApply.
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
    