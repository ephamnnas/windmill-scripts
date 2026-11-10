#requirements:
#requests
#wmill
#json_repair
#python-docx

import wmill
import os
import io
import re
import csv
import datetime
import requests
import json
import json_repair
import time
from urllib.parse import quote
from docx import Document


# =========================================================
# CONFIGURATION
# =========================================================

TENANT_ID = wmill.get_variable("u/epham2/AZURE_TENANT_ID")
CLIENT_ID = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_ID")
CLIENT_SECRET = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_SECRET")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SHAREPOINT_HOSTNAME = "nnasca.sharepoint.com"
SITE_PATH = "/sites/NNAS-Portal-Sharepoint"
DOCUMENT_LIBRARY_NAME = "Applications"

# Application folders live under /Applications/CES/<reference_id>/
FOLDER_RELATIVE_PATH = "CES"

# Location of the Word letter template (with letterhead) in the Applications
# library. The template should contain these placeholder tokens wherever the
# values should appear (body, tables, headers or footers):
#   {{APPLICANT_NAME}}  {{APPLICATION_ID}}  {{NNAS_ID}}  {{DATE}}
TEMPLATE_FOLDER = "IEN-Forms"
TEMPLATE_FILENAME = "Placeholder_Letter_Template.docx"

# Set to True to keep the filled Word file in the application folder alongside
# the PDF. When False, the Word file is only used for PDF conversion and is
# deleted afterwards.
KEEP_FILLED_DOCX = False

# File name used when the placeholder is delivered. Kept identical to the
# generated licensure form so downstream processes treat it the same way.
DELIVERED_FILENAME_TEMPLATE = "%s_IENForm.pdf"
DELIVERED_FILE_TYPE = "Registration"

# When True, the placeholder is only uploaded if the applicant has declared
# at least one jurisdiction (mirrors the original licensure script).
# Set to False to always upload the placeholder.
REQUIRE_JURISDICTION = True

# SMApply task that receives the uploaded form (CES program).
SMAPPLY_TASK_ID = "2053930"

_GRAPH_TOKEN = None
_SP_SITE_ID = None
_SP_DRIVE_ID = None
_SP_LIST_ID = None


# =========================================================
# SMAPPLY CUSTOM FIELD STATUS CONFIG
# =========================================================
# These custom field IDs are shared across both the TM and CES SMApply programs.

IDENTITY_STATUS_FIELD_IDS = [153220, 153221]

REGISTRATION_STATUS_FIELD_IDS = {
    1: 153222,
    2: 153223,
    3: 153224,
    4: 153225,
    5: 153226,
    6: 153227,
    7: 153228,
}

EDUCATION_STATUS_FIELD_IDS = {
    1: 153229,
    2: 153230,
    3: 153231,
    4: 153232,
    5: 153233,
    6: 153234,
    7: 153235,
}

STATUS_UNDER_REVIEW = "Under Review"
STATUS_WAITING_TO_BE_RECEIVED = "Waiting to be Received"


# =========================================================
# SMAPPLY LABEL CONFIG
# =========================================================

SMAPPLY_BASE = "https://nnas.smapply.io/api"

# Label added to each processed application. SMApply has no labels list
# endpoint on this instance, so the label ID must be set here. To find it,
# apply the label to any application in the UI, then GET
# /api/applications/{id}/ and read the "id" of that entry in "labels".
SMAPPLY_LABEL_NAME = "Under vendor review"
SMAPPLY_LABEL_ID = 187213


# =========================================================
# SHAREPOINT / MICROSOFT GRAPH
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
    global _SP_SITE_ID, _SP_DRIVE_ID, _SP_LIST_ID

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


def patch_driveitem_fields(drive_item_id: str, fields: dict) -> None:
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


def download_sharepoint_file(drive_path):
    """Download a file from the Applications library and return its bytes."""
    init_sharepoint_graph()

    item = get_drive_item_by_path(drive_path)
    if not item:
        raise RuntimeError(f"SharePoint file not found: {drive_path}")

    response = requests.get(
        f"{GRAPH_BASE}/drives/{_SP_DRIVE_ID}/items/{item['id']}/content",
        headers={"Authorization": f"Bearer {get_graph_token()}"},
        timeout=120,
        allow_redirects=True,
    )

    if response.status_code != 200:
        raise RuntimeError(
            f"Failed downloading SharePoint file: {drive_path}\n"
            f"Status: {response.status_code}\n"
            f"Response body:\n{response.text[:2000]}"
        )

    print("  Downloaded %s (%d bytes)" % (drive_path, len(response.content)))
    return response.content


# =========================================================
# WORD TEMPLATE FILLING
# =========================================================

def _replace_in_paragraph(paragraph, replacements):
    """
    Replace placeholder tokens in a paragraph while keeping the formatting
    of the surrounding text.

    Word often splits a token like {{APPLICATION_ID}} across several runs
    (for example after spell check or an edit), so a simple per-run replace
    can miss it. This locates each token in the paragraph's combined run
    text, writes the value into the first run the token touches, and trims
    the remaining pieces of the token from the following runs.
    """
    for token, value in replacements.items():
        guard = 0
        while guard < 50:
            guard += 1
            runs = paragraph.runs
            full_text = "".join(run.text for run in runs)
            start = full_text.find(token)
            if start == -1:
                break
            end = start + len(token)

            pos = 0
            wrote_value = False
            for run in runs:
                run_start = pos
                run_end = pos + len(run.text)
                pos = run_end
                if run_end <= start or run_start >= end:
                    continue

                local_start = max(start, run_start) - run_start
                local_end = min(end, run_end) - run_start
                if not wrote_value:
                    run.text = run.text[:local_start] + value + run.text[local_end:]
                    wrote_value = True
                else:
                    run.text = run.text[:local_start] + run.text[local_end:]


def _replace_in_container(container, replacements):
    """Replace tokens in every paragraph and table (including nested tables)."""
    for paragraph in container.paragraphs:
        _replace_in_paragraph(paragraph, replacements)
    for table in container.tables:
        for row in table.rows:
            for cell in row.cells:
                _replace_in_container(cell, replacements)


def fill_docx_template(template_bytes, replacements):
    """
    Fill placeholder tokens in a .docx template and return the new file bytes.
    Covers the document body, tables, and all headers and footers, which is
    where letterhead details usually live.
    """
    document = Document(io.BytesIO(template_bytes))

    _replace_in_container(document, replacements)

    for section in document.sections:
        for part in (
            section.header, section.first_page_header, section.even_page_header,
            section.footer, section.first_page_footer, section.even_page_footer,
        ):
            if part is not None and not part.is_linked_to_previous:
                _replace_in_container(part, replacements)

    output = io.BytesIO()
    document.save(output)
    filled_bytes = output.getvalue()

    # Flag any tokens that were not filled (typos in the template, or tokens
    # placed inside text boxes, which python-docx does not edit).
    leftover = set(re.findall(r"\{\{[A-Z_]+\}\}", "\n".join(p.text for p in document.paragraphs)))
    if leftover:
        print("  Warning: unfilled placeholders remain in the letter: %s" % ", ".join(sorted(leftover)))

    return filled_bytes


def convert_docx_to_pdf_via_graph(docx_bytes, reference_id):
    """
    Convert a .docx to PDF using Microsoft Graph, so no office software is
    needed on the Windmill worker. The Word file is uploaded to the
    application folder, downloaded back with ?format=pdf, and then deleted
    unless KEEP_FILLED_DOCX is True.
    """
    init_sharepoint_graph()

    application_folder_path = "%s/%s" % (FOLDER_RELATIVE_PATH.rstrip("/"), reference_id)
    if not get_drive_item_by_path(application_folder_path):
        raise RuntimeError(
            f"SharePoint application folder not found. "
            f"Expected library-relative path: {application_folder_path}"
        )

    if KEEP_FILLED_DOCX:
        docx_name = "%s_PlaceholderLetter.docx" % reference_id
    else:
        docx_name = "%s_PlaceholderLetter_temp.docx" % reference_id
    docx_path = "%s/%s" % (application_folder_path, docx_name)

    uploaded = graph_request(
        "PUT",
        f"/drives/{_SP_DRIVE_ID}/root:/{encode_graph_path(docx_path)}:/content",
        headers={"Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"},
        data=docx_bytes,
        expected_statuses=(200, 201),
    )

    try:
        response = requests.get(
            f"{GRAPH_BASE}/drives/{_SP_DRIVE_ID}/items/{uploaded['id']}/content?format=pdf",
            headers={"Authorization": f"Bearer {get_graph_token()}"},
            timeout=180,
            allow_redirects=True,
        )

        if response.status_code != 200 or not response.content.startswith(b"%PDF"):
            raise RuntimeError(
                "Microsoft Graph PDF conversion failed.\n"
                f"Status: {response.status_code}\n"
                f"Response body:\n{response.text[:2000]}"
            )

        print("  Converted letter to PDF (%d bytes)" % len(response.content))
        return response.content
    finally:
        if not KEEP_FILLED_DOCX:
            try:
                graph_request(
                    "DELETE",
                    f"/drives/{_SP_DRIVE_ID}/items/{uploaded['id']}",
                    expected_statuses=(204, 404),
                )
            except RuntimeError as exc:
                print("  Warning: could not delete temporary Word file %s: %s" % (docx_path, exc))


def _format_letter_date(when=None):
    when = when or datetime.date.today()
    return "%s %d, %d" % (when.strftime("%B"), when.day, when.year)


def build_placeholder_pdf(reference_id, user_id, user, local_path):
    """
    Pull the letterhead template from SharePoint, fill in the applicant's
    details, convert it to PDF, and write it to local_path.
    """
    template_bytes = download_sharepoint_file("%s/%s" % (TEMPLATE_FOLDER, TEMPLATE_FILENAME))

    replacements = {
        "{{APPLICANT_NAME}}": _title_case_name(user),
        "{{APPLICATION_ID}}": str(reference_id),
        "{{NNAS_ID}}": str(user_id),
        "{{DATE}}": _format_letter_date(),
    }

    filled_docx = fill_docx_template(template_bytes, replacements)
    pdf_bytes = convert_docx_to_pdf_via_graph(filled_docx, reference_id)

    os.makedirs(os.path.dirname(local_path), exist_ok=True)
    with open(local_path, "wb") as f:
        f.write(pdf_bytes)

    return local_path


def upload_to_sharepoint(local_filepath, reference_id):
    init_sharepoint_graph()

    application_folder_path = "%s/%s" % (FOLDER_RELATIVE_PATH.rstrip("/"), reference_id)
    target_folder = get_drive_item_by_path(application_folder_path)

    if not target_folder:
        raise RuntimeError(
            f"SharePoint application folder not found. "
            f"Expected library-relative path: {application_folder_path}"
        )

    with open(local_filepath, "rb") as f:
        file_data = f.read()

    sharepoint_filename = DELIVERED_FILENAME_TEMPLATE % reference_id
    upload_path = "%s/%s" % (application_folder_path, sharepoint_filename)

    uploaded_file = graph_request(
        "PUT",
        f"/drives/{_SP_DRIVE_ID}/root:/{encode_graph_path(upload_path)}:/content",
        headers={"Content-Type": "application/pdf"},
        data=file_data,
        expected_statuses=(200, 201),
    )

    patch_driveitem_fields(uploaded_file["id"], {"FileType": DELIVERED_FILE_TYPE})
    print("  Uploaded placeholder form to SharePoint: %s" % upload_path)


def upload_csv_to_sharepoint(csv_bytes, filename, reference_id):
    init_sharepoint_graph()

    application_folder_path = "%s/%s" % (FOLDER_RELATIVE_PATH.rstrip("/"), reference_id)
    target_folder = get_drive_item_by_path(application_folder_path)

    if not target_folder:
        raise RuntimeError(
            f"SharePoint application folder not found. "
            f"Expected library-relative path: {application_folder_path}"
        )

    upload_path = "%s/%s" % (application_folder_path, filename)

    graph_request(
        "PUT",
        f"/drives/{_SP_DRIVE_ID}/root:/{encode_graph_path(upload_path)}:/content",
        headers={"Content-Type": "text/csv; charset=utf-8"},
        data=csv_bytes,
        expected_statuses=(200, 201),
    )

    print("  Uploaded CSV: %s (%d bytes)" % (filename, len(csv_bytes)))


# =========================================================
# NAME FORMATTING
# =========================================================

def _title_case_name(name):
    """
    Title-case a person's name, keeping hyphenated names and names with
    apostrophes capitalized on each segment (e.g. "Smith-Jones", "O'Brien").
    """
    if not name or not isinstance(name, str):
        return name

    def _cap_segment(segment):
        return segment[:1].upper() + segment[1:].lower() if segment else segment

    def _cap_word(word):
        return re.sub(r"[^\-']+", lambda m: _cap_segment(m.group(0)), word)

    return " ".join(_cap_word(w) for w in name.split(" "))


# =========================================================
# PAYLOAD PARSING
# =========================================================

def _parse_additional_content(raw):
    """
    Parse the additional_content payload into a dict.

    SMApply free-text fields sometimes contain unescaped double quotes,
    which produces invalid JSON. A strict parse is tried first, and a
    tolerant repair pass is used only if the strict parse fails.
    """
    if raw is None:
        raise ValueError("additional_content / ADDITIONAL_CONTENT was not provided.")

    if isinstance(raw, dict):
        return raw

    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError as exc:
        print("  additional_content is not valid JSON (%s). Attempting repair." % exc)
        parsed = json_repair.loads(raw)
        if not isinstance(parsed, dict):
            raise RuntimeError(
                "additional_content could not be repaired into a JSON object "
                "(got %s)." % type(parsed).__name__
            )
        print("  additional_content repaired successfully.")

    # Occasionally the payload is a JSON string that itself contains JSON.
    if isinstance(parsed, str):
        try:
            parsed = json.loads(parsed)
        except json.JSONDecodeError:
            parsed = json_repair.loads(parsed)

    if not isinstance(parsed, dict):
        raise RuntimeError(
            "additional_content did not parse into a JSON object (got %s)."
            % type(parsed).__name__
        )

    return parsed


def has_any_jurisdiction(parsed):
    return any(
        (parsed.get("sa_jurisdiction_name_%d" % i) or "").strip()
        for i in range(1, 8)
    )


# =========================================================
# SMAPPLY
# =========================================================

def _extract_task_files(task_json):
    """
    Locate uploaded-file objects in a task detail payload so they can be
    deleted before a fresh upload. The .../tasks/{id}/files/ collection does
    not allow GET, so files are read from the task detail instead.
    Returns a list of file dicts, de-duplicated by id.
    """
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

    # Recursive fallback.
    signature_keys = {
        "file", "filename", "original_filename", "download_url", "url",
        "content_type", "mimetype", "size", "filesize", "extension",
    }
    found = []
    seen = set()

    def walk(obj):
        if isinstance(obj, dict):
            if "id" in obj and any(k in obj for k in signature_keys):
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


def update_smapply_task(appid, pdf_filepath, reference_id):
    task_id = SMAPPLY_TASK_ID
    access_token = wmill.get_variable("u/epham2/SMAPPLY_ACCESS_TOKEN")
    headers = {"Authorization": "Bearer %s" % access_token}
    base_url = "https://nnas.smapply.io/api/applications/%s/tasks/%s/files/" % (appid, task_id)
    task_url = "https://nnas.smapply.io/api/applications/%s/tasks/%s/" % (appid, task_id)

    # Step 1: Remove any existing files on the task.
    task_resp = requests.get(task_url, headers=headers, timeout=120)
    task_resp.raise_for_status()
    existing_files = _extract_task_files(task_resp.json())

    if existing_files:
        print("  Found %d existing file(s) on task %s. Deleting before upload." % (len(existing_files), task_id))

    for file in existing_files:
        file_id = file.get("id")
        if file_id is None:
            continue
        del_resp = requests.delete("%s%s/" % (base_url, file_id), headers=headers, timeout=120)
        # A 404 means the file is already gone; treat that as success.
        if del_resp.status_code not in (200, 202, 204, 404):
            del_resp.raise_for_status()
        print("  Deleted existing file %s from task %s" % (file_id, task_id))

    # Step 2: Upload the placeholder form.
    with open(pdf_filepath, "rb") as f:
        file_data = f.read()
    upload_resp = requests.post(
        base_url,
        files={"file": (DELIVERED_FILENAME_TEMPLATE % reference_id, file_data, "application/octet-stream")},
        headers=headers,
        timeout=120,
    )
    upload_resp.raise_for_status()
    print("  Uploaded placeholder form to SMApply task %s" % task_id)


def update_smapply_custom_field_statuses(appid, parsed):
    """
    Update per-application custom field values on SMApply:
      - Identity document fields are set to "Under Review".
      - Registration and Education status fields are set to
        "Waiting to be Received" for every declared slot (1-7).
    """
    field_values = {}

    for field_id in IDENTITY_STATUS_FIELD_IDS:
        if field_id is not None:
            field_values[field_id] = STATUS_UNDER_REVIEW

    for i in range(1, 8):
        jname = (parsed.get("sa_jurisdiction_name_%d" % i) or "").strip()
        lic   = (parsed.get("sa_registration_number_%d" % i) or "").strip()
        field_id = REGISTRATION_STATUS_FIELD_IDS.get(i)
        if (jname or lic) and field_id is not None:
            field_values[field_id] = STATUS_WAITING_TO_BE_RECEIVED

    for i in range(1, 8):
        school  = (parsed.get("ed_school_name_%d" % i) or "").strip()
        program = (parsed.get("ed_program_name_%d" % i) or "").strip()
        field_id = EDUCATION_STATUS_FIELD_IDS.get(i)
        if (school or program) and field_id is not None:
            field_values[field_id] = STATUS_WAITING_TO_BE_RECEIVED

    if not field_values:
        print("  No custom field statuses to update.")
        return

    access_token = wmill.get_variable("u/epham2/SMAPPLY_ACCESS_TOKEN")
    headers = {
        "Authorization": "Bearer %s" % access_token,
        "Content-Type": "application/json",
    }
    url = "https://nnas.smapply.io/api/applications/%s/" % appid
    body = {
        "custom_fields": [
            {"id": field_id, "value": value}
            for field_id, value in field_values.items()
        ]
    }

    resp = requests.patch(url, headers=headers, json=body, timeout=120)
    resp.raise_for_status()
    print("  Updated %d custom field(s) on application %s" % (len(field_values), appid))


# =========================================================
def _smapply_headers(json_body=False):
    access_token = wmill.get_variable("u/epham2/SMAPPLY_ACCESS_TOKEN")
    headers = {"Authorization": "Bearer %s" % access_token}
    if json_body:
        headers["Content-Type"] = "application/json"
    return headers


def get_smapply_application(appid):
    resp = requests.get(
        "%s/applications/%s/" % (SMAPPLY_BASE, appid),
        headers=_smapply_headers(),
        timeout=120,
    )
    resp.raise_for_status()
    return resp.json()


# ---------- Label assignment ----------

def _label_ids(labels):
    """
    Extract label IDs from an application's "labels" value. SMApply returns
    labels as [{"id": int, "name": str}], but bare IDs are tolerated too.
    """
    out = []
    for item in (labels or []):
        raw = item.get("id") if isinstance(item, dict) else item
        try:
            out.append(int(raw))
        except (TypeError, ValueError):
            continue
    return out


def add_smapply_label(appid):
    """
    Add the configured label to the application.

    SMApply reads labels as nested objects but writes them as a plain list of
    IDs, and the list replaces the full set. Existing labels are read first and
    merged so nothing applied by an automation is dropped. Success is
    confirmed by reading the application back, not by status code alone.
    """
    if SMAPPLY_LABEL_ID is None:
        raise RuntimeError(
            "SMAPPLY_LABEL_ID is not set. Add the numeric ID of the '%s' label "
            "at the top of the script." % SMAPPLY_LABEL_NAME
        )
    label_id = int(SMAPPLY_LABEL_ID)
    url = "%s/applications/%s/" % (SMAPPLY_BASE, appid)

    existing = _label_ids(get_smapply_application(appid).get("labels"))
    if label_id in existing:
        print("  Label '%s' is already on application %s" % (SMAPPLY_LABEL_NAME, appid))
        return

    merged = existing + [label_id]
    resp = requests.patch(
        url,
        headers=_smapply_headers(json_body=True),
        json={"labels": merged},
        timeout=120,
    )
    if not resp.ok:
        raise RuntimeError(
            "PATCH %s labels=%s returned %s: %s"
            % (url, merged, resp.status_code, resp.text[:2000])
        )

    after = _label_ids(get_smapply_application(appid).get("labels"))
    if label_id not in after:
        raise RuntimeError(
            "Label %s is not on application %s after the update (labels now %s)."
            % (label_id, appid, after)
        )
    print("  Added label '%s' (id %s) to application %s" % (SMAPPLY_LABEL_NAME, label_id, appid))


# =========================================================
# DOCUMENT INDEX CSV
# =========================================================

def generate_document_index_csv(parsed):
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["Number", "Document Title", "Tracking Date", "Uploaded"])
    idx = 0

    # Education: ECA (always doc #1)
    eca_completed = (parsed.get("ed_eca_completed") or "").strip().lower()
    eca_agency    = (parsed.get("ed_eca_agency") or "").strip()
    if eca_completed == "yes" and eca_agency:
        eca_title = eca_agency
    else:
        eca_title = "Applicant has not indicated provider"
    idx += 1
    w.writerow([idx, "ECA - %s" % eca_title, "", ""])

    # Section A: Registration documents
    for i in range(1, 8):
        jname = (parsed.get("sa_jurisdiction_name_%d" % i) or "").strip()
        lic   = (parsed.get("sa_registration_number_%d" % i) or "").strip()
        if not jname and not lic:
            continue
        title = "%s: %s" % (jname, lic) if jname and lic else (jname or lic)
        idx += 1
        w.writerow([idx, "Registration - %s" % title, "", ""])

    # Education: School/program documents
    for i in range(1, 8):
        school  = (parsed.get("ed_school_name_%d" % i) or "").strip()
        program = (parsed.get("ed_program_name_%d" % i) or "").strip()
        if not school and not program:
            continue
        title = "%s: %s" % (school, program) if school and program else (school or program)
        idx += 1
        w.writerow([idx, "Education - %s" % title, "", ""])

    csv_bytes = buf.getvalue().encode("utf-8-sig")
    buf.close()
    return csv_bytes


# =========================================================
# MAIN
# =========================================================

def main(
    REFERENCE_ID=None,
    USER_ID=None,
    USER=None,
    appid=None,
    ADDITIONAL_CONTENT=None,
    reference_id=None,
    user_id=None,
    user=None,
    id=None,
    additional_content=None,
    **kwargs,
):
    if REFERENCE_ID is None:
        REFERENCE_ID = reference_id or kwargs.get("title")

    if USER_ID is None:
        USER_ID = user_id

    if USER is None:
        USER = user

    if appid is None:
        appid = id or kwargs.get("application_id") or kwargs.get("appid")

    if ADDITIONAL_CONTENT is None:
        ADDITIONAL_CONTENT = additional_content

    if REFERENCE_ID is None:
        raise ValueError("REFERENCE_ID / reference_id was not provided.")

    if USER_ID is None:
        raise ValueError("USER_ID / user_id was not provided.")

    if USER is None:
        raise ValueError("USER / user was not provided.")

    if appid is None:
        raise ValueError("appid / id was not provided.")

    parsed = _parse_additional_content(ADDITIONAL_CONTENT)

    pdf_dir = os.path.join(os.getcwd(), "tmp")
    pdf_path = os.path.join(pdf_dir, "%s_placeholder_form.pdf" % REFERENCE_ID)

    time.sleep(20)

    # Placeholder form delivery
    form_uploaded = False
    if REQUIRE_JURISDICTION and not has_any_jurisdiction(parsed):
        print("No jurisdictions found. Skipping placeholder form upload.")
    else:
        build_placeholder_pdf(REFERENCE_ID, USER_ID, USER, pdf_path)
        upload_to_sharepoint(pdf_path, REFERENCE_ID)
        update_smapply_task(appid, pdf_path, REFERENCE_ID)
        form_uploaded = True

    # Document index CSV
    try:
        csv_bytes = generate_document_index_csv(parsed)
        upload_csv_to_sharepoint(csv_bytes, "%s_DocumentIndex.csv" % REFERENCE_ID, REFERENCE_ID)
    except (json.JSONDecodeError, KeyError, TypeError, ValueError, RuntimeError) as e:
        print("  Warning: CSV generation failed due to unexpected data format: %s" % e)
    except requests.HTTPError as e:
        print("  Warning: CSV upload to SharePoint failed: %s" % e)

    # Custom field status updates
    try:
        update_smapply_custom_field_statuses(appid, parsed)
    except (json.JSONDecodeError, KeyError, TypeError, ValueError, RuntimeError) as e:
        print("  Warning: custom field status update failed due to unexpected data format: %s" % e)
    except requests.HTTPError as e:
        print("  Warning: custom field status update to SMApply failed: %s" % e)

    # Label assignment
    try:
        add_smapply_label(appid)
    except (KeyError, TypeError, ValueError, RuntimeError) as e:
        print("  Warning: label assignment failed: %s" % e)
    except requests.HTTPError as e:
        print("  Warning: label assignment to SMApply failed: %s" % e)

    # Cleanup
    try:
        os.remove(pdf_path)
    except OSError:
        pass

    return {
        "placeholder_uploaded": form_uploaded,
        "reference_id": REFERENCE_ID,
        "appid": appid,
    }