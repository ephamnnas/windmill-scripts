#requirements:
#pypdf
#reportlab
#pillow
#qrcode
#requests
#wmill
#json_repair

import wmill
import os
import io
import re
import csv
import qrcode
import requests
import json
import json_repair
import time
from urllib.parse import quote
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from PIL import Image, ImageDraw, ImageFont
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, BooleanObject


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

# Old SharePoint path:
# /sites/NNAS-Portal-Sharepoint/Applications/CES/
#
# In Graph, this is relative to the Applications document library root.
FOLDER_RELATIVE_PATH = "CES"

_GRAPH_TOKEN = None
_SP_SITE_ID = None
_SP_DRIVE_ID = None
_SP_LIST_ID = None


# =========================================================
# SMAPPLY CUSTOM FIELD STATUS CONFIG
# =========================================================
# Confirmed: these custom field IDs are shared across both the TM and CES
# SMApply programs, despite the two programs using different task IDs
# (2044278 for TM, 2053930 for CES) for the file upload step.

# Identity documents are uploaded directly by applicants, so we don't wait
# on any external confirmation before marking them for review.
IDENTITY_STATUS_FIELD_IDS = [153220, 153221]

# Registration / education status fields, keyed by the same slot index (1-7)
# used for sa_jurisdiction_name_<i> / ed_school_name_<i> throughout this script.
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


def download_licensure_form_from_sharepoint():
    folder_name = "IEN-Forms"
    filename = "Nursing_Registration_Form.pdf"
    drive_path = f"{folder_name}/{filename}"

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

    local_path = os.path.join("tmp", filename)
    os.makedirs("tmp", exist_ok=True)

    with open(local_path, "wb") as local_file:
        local_file.write(response.content)

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

    sharepoint_filename = "%s_IENForm.pdf" % reference_id
    upload_path = "%s/%s" % (application_folder_path, sharepoint_filename)

    uploaded_file = graph_request(
        "PUT",
        f"/drives/{_SP_DRIVE_ID}/root:/{encode_graph_path(upload_path)}:/content",
        headers={"Content-Type": "application/pdf"},
        data=file_data,
        expected_statuses=(200, 201),
    )

    patch_driveitem_fields(uploaded_file["id"], {"FileType": "Registration"})


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
# PAYLOAD PARSING
# =========================================================

def _parse_additional_content(raw):
    """
    Parse the additional_content payload into a dict.

    SMApply free-text fields (school names, program names, etc.) sometimes
    contain literal double quotes that were never escaped when the payload
    was serialized upstream, e.g.:

        "ed_school_name_1": "State Educational Institution "Eisk Medical
        College" of Krasnodar Regional Department of Public Healthcare"

    That produces invalid JSON and json.loads raises "Expecting ',' delimiter".
    Try a strict parse first (the normal, well-formed case) and fall back to a
    tolerant repair pass so a single stray quote does not kill the whole
    delivery. The strict path is preferred because it never guesses; the
    repair path only runs when strict parsing has already failed.
    """
    if raw is None:
        raise ValueError("additional_content / ADDITIONAL_CONTENT was not provided.")

    if isinstance(raw, dict):
        return raw

    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError as exc:
        print("  additional_content is not valid JSON (%s) — attempting repair." % exc)
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


# =========================================================
# NAME FORMATTING
# =========================================================

def _title_case_name(name):
    """
    Title-case a person's name for display on the generated form.

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
# QR CODE
# =========================================================

def _generate_qr_image(data):
    """Return a PIL Image of the QR code, resized to half scale."""
    qr = qrcode.QRCode(box_size=10, border=4)
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white").convert("RGB")

    try:
        resample_filter = Image.Resampling.LANCZOS
    except AttributeError:
        resample_filter = Image.LANCZOS
    img = img.resize((img.width // 2, img.height // 2), resample=resample_filter)
    return img


def _build_combined_header_image(header_text, qr_img):
    """
    Combine the text header and QR image side by side into one PIL Image.
    Returns the combined image and its scaled (width, height) in points
    for placement on a Letter page.
    """
    try:
        font = ImageFont.truetype("arial.ttf", 14)
    except IOError:
        font = ImageFont.load_default()

    # Measure text
    temp_img = Image.new("RGB", (1, 1), "white")
    draw = ImageDraw.Draw(temp_img)
    bbox = draw.multiline_textbbox((0, 0), header_text, font=font)
    text_w, text_h = bbox[2] - bbox[0], bbox[3] - bbox[1]
    padding, left_spacing = 10, 20

    header_img = Image.new(
        "RGB",
        (text_w + left_spacing + padding, text_h + 2 * padding),
        "white",
    )
    draw = ImageDraw.Draw(header_img)
    draw.multiline_text((left_spacing + padding, padding), header_text, fill="black", font=font)

    try:
        resample_filter = Image.Resampling.LANCZOS
    except AttributeError:
        resample_filter = Image.LANCZOS
    header_img = header_img.resize(
        (int(header_img.width * 0.8), int(header_img.height * 0.8)),
        resample=resample_filter,
    )

    header_w, header_h = header_img.size
    qr_w, qr_h = qr_img.size
    combined = Image.new("RGB", (header_w + qr_w, max(header_h, qr_h)), "white")
    combined.paste(header_img, (0, 0))
    combined.paste(qr_img, (header_w, 0))

    page_width, _ = letter
    img_w, img_h = combined.size
    scale = min(1.0, (page_width * 0.25) / img_w)
    draw_w = img_w * scale
    draw_h = img_h * scale

    return combined, draw_w, draw_h


# =========================================================
# OVERLAY BUILDER  (ReportLab + Pillow)
# =========================================================

CES_ADDRESS_LINES = [
    "Comparative Education Service (CES)",
    "University of Toronto School of Continuing Studies",
    "c/o Research Department",
    "158 St. George Street",
    "Toronto, ON, M5S 2V8",
    "Canada",
]


def _make_overlay_clause(combined_img, draw_w, draw_h, applicant_name, licensing_body):
    """
    Build a single-page Letter overlay PDF with the header/QR image and
    the consent clause + CES address. Returns raw PDF bytes.
    """
    page_width, page_height = letter

    clause_text = (
        "I, <b>%s</b>, hereby give my consent to "
        "<b>%s</b> "
        "to provide the information requested in Part B of this form related "
        "to my nurse registration, and to send this completed form directly "
        "to NNAS at the following address:" % (applicant_name, licensing_body)
    )

    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=letter)

    # Header + QR image at top-left
    c.drawImage(ImageReader(combined_img), 0, page_height - draw_h, width=draw_w, height=draw_h)

    # Consent clause paragraph
    styles = getSampleStyleSheet()
    style = styles["Normal"]
    style.fontName = "Times-Roman"
    style.fontSize = 10
    style.leading = 14

    text_x = 0.75 * inch
    text_y = 3.5 * inch
    p = Paragraph(clause_text, style)
    w, h = p.wrap(7 * inch, 2 * inch)
    p.drawOn(c, text_x, text_y)

    # CES address block
    c.setFont("Times-Roman", 10)
    start_y = text_y - h - 3
    for idx, line in enumerate(CES_ADDRESS_LINES):
        c.drawString(text_x, start_y - (idx * 12), line)

    c.save()
    packet.seek(0)
    return packet.read()


def _make_overlay_qr(combined_img, draw_w, draw_h):
    """
    Build a single-page Letter overlay PDF with only the header/QR image.
    Returns raw PDF bytes.
    """
    _, page_height = letter

    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=letter)
    c.drawImage(ImageReader(combined_img), 0, page_height - draw_h, width=draw_w, height=draw_h)
    c.save()
    packet.seek(0)
    return packet.read()


# =========================================================
# PDF GENERATION
# =========================================================

def create_pdf(reference_id, user_id, user, output_filename, additional_content):
    """
    Build the combined licensure PDF for all jurisdictions found in
    additional_content. Returns True if a PDF was written, False if
    there were no jurisdictions to process.
    """
    parsed = _parse_additional_content(additional_content)
    combined_writer = PdfWriter()

    if not any((parsed.get("sa_jurisdiction_name_%d" % i) or "").strip() for i in range(1, 8)):
        print("No jurisdictions found — skipping PDF creation.")
        return False

    # Use the passed-in user string directly (not constructed from sa_ name fields),
    # but normalize casing so it renders in Title Case on the generated form.
    applicant_name = _title_case_name(user)

    licensure_pdf_path = download_licensure_form_from_sharepoint()

    # Read the source form once into memory. Each jurisdiction still needs a
    # fresh PdfReader (merge_page mutates page objects), but reparsing from
    # these bytes avoids re-reading from disk on every iteration.
    with open(licensure_pdf_path, "rb") as _f:
        licensure_bytes = _f.read()

    header_text = "NNASID: %s\nAPPLICATIONID: %s\nApplicant Name: %s" % (user_id, reference_id, applicant_name)

    # Build the QR image and combined header once — reused across all jurisdictions.
    qr_img = _generate_qr_image(reference_id)
    combined_img, draw_w, draw_h = _build_combined_header_image(header_text, qr_img)

    # QR-only overlay bytes are the same for every jurisdiction — build once.
    overlay_qr_bytes = _make_overlay_qr(combined_img, draw_w, draw_h)

    for license_index in range(1, 8):
        jname_key = "sa_jurisdiction_name_%d" % license_index
        if not (parsed.get(jname_key) or "").strip():
            continue

        licensing_body = (parsed.get(jname_key) or "").strip() or "the regulatory body"

        overlay_clause_bytes = _make_overlay_clause(combined_img, draw_w, draw_h, applicant_name, licensing_body)

        # Parse fresh page objects every iteration — merge_page mutates them
        # and a reused page object produces corrupted output on the second
        # jurisdiction onward. The bytes are cheap to reparse.
        overlay_clause_page = PdfReader(io.BytesIO(overlay_clause_bytes)).pages[0]
        overlay_qr_page     = PdfReader(io.BytesIO(overlay_qr_bytes)).pages[0]

        reader_pdf = PdfReader(io.BytesIO(licensure_bytes))
        writer_pdf = PdfWriter()

        if "/AcroForm" in reader_pdf.trailer["/Root"]:
            writer_pdf._root_object.update({
                NameObject("/AcroForm"): reader_pdf.trailer["/Root"]["/AcroForm"]
            })
            writer_pdf._root_object["/AcroForm"].update({
                NameObject("/NeedAppearances"): BooleanObject(True)
            })

        def get_app(key, default="N/A"):
            val = parsed.get(key, "")
            if isinstance(val, str):
                val = val.strip()
            return val if val not in ("", None) else default

        def get_reg(base_key, default="N/A"):
            key = "%s_%d" % (base_key, license_index)
            val = parsed.get(key, "")
            if isinstance(val, str):
                val = val.strip()
            return val if val not in ("", None) else default

        def is_yes(val):
            return str(val or "").strip().lower() in ("yes", "y", "true", "1")

        mailing_same = is_yes(parsed.get("sa_mailing_same_as_current"))
        chosen_prefix = "sa_applicant_current_address" if mailing_same else "sa_applicant_mailing_address"

        def addr(suffix, default="N/A"):
            v = parsed.get("%s_%s" % (chosen_prefix, suffix), "")
            if isinstance(v, str):
                v = v.strip()
            return v if v not in ("", None) else default

        form_fields = {
            "sa_first_name":            _title_case_name(get_app("sa_first_name")),
            "sa_middle_name":           _title_case_name(get_app("sa_middle_name")),
            "sa_last_name":             _title_case_name(get_app("sa_last_name")),
            "sa_other_names":           _title_case_name(get_app("sa_other_names")),
            "sa_dob":                   get_app("sa_dob"),
            "sa_phone_number":          get_app("sa_phone_number"),
            "sa_email":                 get_app("sa_email"),
            "sa_applicant_address_1":   addr("1"),
            "sa_applicant_address_2":   addr("2"),
            "sa_applicant_city":        addr("city"),
            "sa_applicant_province":    addr("province"),
            "sa_applicant_postal_code": addr("postal_code"),
            "sa_applicant_country":     addr("country"),
            "sa_registrant_name":       _title_case_name(get_reg("sa_registrant_name")),
            "sa_jurisdiction_name":     get_reg("sa_jurisdiction_name"),
            "sa_nnas_id":               user_id,
            "sa_application_id":        reference_id,
            "sc_application_id":        reference_id,
        }

        num_pages = len(reader_pdf.pages)
        for i, page in enumerate(reader_pdf.pages):
            if i == 0:
                page.merge_page(overlay_clause_page)
            elif i == num_pages - 1:
                page.merge_page(overlay_qr_page)
            writer_pdf.add_page(page)

        for page in writer_pdf.pages:
            writer_pdf.update_page_form_field_values(page, form_fields)

        buffer = io.BytesIO()
        writer_pdf.write(buffer)
        buffer.seek(0)
        for pg in PdfReader(buffer).pages:
            combined_writer.add_page(pg)

    if len(combined_writer.pages) > 0:
        with open(output_filename, "wb") as f_out:
            combined_writer.write(f_out)
        print("Combined PDF created for %s" % reference_id)
        return True

    return False


# =========================================================
# SMAPPLY
# =========================================================

def _extract_task_files(task_json):
    """
    Locate uploaded-file objects in a task detail payload so they can be
    deleted before a fresh upload.

    The .../tasks/{id}/files/ collection does not allow GET (403), so files
    are read from the task detail instead. This tries the known 'data.files'
    location first, then a top-level 'files', then falls back to a recursive
    search for file-like objects (an 'id' alongside a file/name/url/size-type
    key) so existing files are found regardless of where the API nests them.
    Returns a list of file dicts, de-duplicated by id.
    """
    if not isinstance(task_json, dict):
        return []

    # Known/likely locations first.
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
    task_id = "2053930"
    access_token = wmill.get_variable("u/epham2/SMAPPLY_ACCESS_TOKEN")
    headers = {"Authorization": "Bearer %s" % access_token}
    base_url = "https://nnas.smapply.io/api/applications/%s/tasks/%s/files/" % (appid, task_id)
    task_url = "https://nnas.smapply.io/api/applications/%s/tasks/%s/" % (appid, task_id)

    # Step 1: Check for existing files on the task and delete each one before
    # uploading. Files are read from the task detail (the .../files/ collection
    # does not allow GET) and located robustly regardless of nesting.
    task_resp = requests.get(task_url, headers=headers, timeout=120)
    task_resp.raise_for_status()
    existing_files = _extract_task_files(task_resp.json())

    if existing_files:
        print("  Found %d existing file(s) on task %s — deleting before upload." % (len(existing_files), task_id))

    for file in existing_files:
        file_id = file.get("id")
        if file_id is None:
            continue
        del_resp = requests.delete("%s%s/" % (base_url, file_id), headers=headers, timeout=120)
        # A 404 means the file is already gone; treat that as success.
        if del_resp.status_code not in (200, 202, 204, 404):
            del_resp.raise_for_status()
        print("  Deleted existing file %s from task %s" % (file_id, task_id))

    # Step 2: Upload the new file
    with open(pdf_filepath, "rb") as f:
        file_data = f.read()
    upload_resp = requests.post(
        base_url,
        files={"file": ("%s_IENForm.pdf" % reference_id, file_data, "application/octet-stream")},
        headers=headers,
        timeout=120,
    )
    upload_resp.raise_for_status()
    print("  Uploaded new file to task %s" % task_id)


def update_smapply_custom_field_statuses(appid, parsed):
    """
    Update per-application custom field values on SMApply to reflect
    document review status:
      - Identity document fields are always set to "Under Review", since
        those are applicant-uploaded and don't require anything to arrive
        from a third party.
      - Registration and Education status fields are set to
        "Waiting to be Received" for every slot (1-7) the applicant has
        actually declared -- i.e. the same slots that appear in the
        document index CSV.

    Uses PATCH /api/applications/{appid}/ with a custom_fields array,
    which updates per-application field *values* (not the field
    definitions themselves) via a partial merge — confirmed against a
    test application, so fields not included here are left untouched.
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
# DOCUMENT INDEX CSV
# =========================================================

def generate_document_index_csv(parsed):
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["Number", "Document Title", "Tracking Date", "Uploaded"])
    idx = 0

    # Education — ECA  (always doc #1; placeholder title unless completed with a named agency)
    eca_completed = (parsed.get("ed_eca_completed") or "").strip().lower()
    eca_agency    = (parsed.get("ed_eca_agency") or "").strip()
    if eca_completed == "yes" and eca_agency:
        eca_title = eca_agency
    else:
        eca_title = "Applicant has not indicated provider"
    idx += 1
    w.writerow([idx, "ECA - %s" % eca_title, "", ""])

    # Section A — Registration documents
    for i in range(1, 8):
        jname = (parsed.get("sa_jurisdiction_name_%d" % i) or "").strip()
        lic   = (parsed.get("sa_registration_number_%d" % i) or "").strip()
        if not jname and not lic:
            continue
        title = "%s: %s" % (jname, lic) if jname and lic else (jname or lic)
        idx += 1
        w.writerow([idx, "Registration - %s" % title, "", ""])

    # Education — School/program documents
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
    os.makedirs(pdf_dir, exist_ok=True)
    pdf_path = os.path.join(pdf_dir, "%s_generated_form.pdf" % REFERENCE_ID)

    time.sleep(20)

    # PDF generation and delivery
    pdf_created = create_pdf(REFERENCE_ID, USER_ID, USER, pdf_path, parsed)
    if pdf_created:
        upload_to_sharepoint(pdf_path, REFERENCE_ID)
        update_smapply_task(appid, pdf_path, REFERENCE_ID)

    # Document index CSV
    try:
        csv_bytes = generate_document_index_csv(parsed)
        upload_csv_to_sharepoint(csv_bytes, "%s_DocumentIndex.csv" % REFERENCE_ID, REFERENCE_ID)
    except (json.JSONDecodeError, KeyError, TypeError, ValueError, RuntimeError) as e:
        print("  Warning: CSV generation failed — unexpected data format: %s" % e)
    except requests.HTTPError as e:
        print("  Warning: CSV upload to SharePoint failed: %s" % e)

    # Custom field status updates
    try:
        update_smapply_custom_field_statuses(appid, parsed)
    except (json.JSONDecodeError, KeyError, TypeError, ValueError, RuntimeError) as e:
        print("  Warning: custom field status update failed — unexpected data format: %s" % e)
    except requests.HTTPError as e:
        print("  Warning: custom field status update to SMApply failed: %s" % e)

    # Cleanup
    try:
        os.remove(pdf_path)
    except OSError:
        pass

    return {
        "pdf_created": pdf_created,
        "reference_id": REFERENCE_ID,
        "appid": appid,
    }