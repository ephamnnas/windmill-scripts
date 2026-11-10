#requirements:
#pypdf
#requests
#wmill
#json_repair

import wmill
import os
import io
import json
import json_repair
import requests
from urllib.parse import quote
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, BooleanObject


# =========================================================
# CONFIGURATION
#
# Fill these in for the specific SMApply task / PDF this webhook targets.
# =========================================================

# --- Where the blank PDF template lives ---------------------------------
# Option A (simplest): bundle the blank template as a Windmill resource /
# attach it to this script's folder and reference it by path.
# Option B (matches your existing script): pull it from SharePoint via
# Microsoft Graph. That block is included below, behind USE_SHAREPOINT_TEMPLATE,
# so you can delete it entirely if the template will live somewhere else.
USE_SHAREPOINT_TEMPLATE = True

TENANT_ID = wmill.get_variable("u/epham2/AZURE_TENANT_ID")
CLIENT_ID = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_ID")
CLIENT_SECRET = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_SECRET")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SHAREPOINT_HOSTNAME = "nnasca.sharepoint.com"
SITE_PATH = "/sites/NNAS-Portal-Sharepoint"
DOCUMENT_LIBRARY_NAME = "Applications"

TEMPLATE_FOLDER_NAME = "IEN-Forms"          # TODO: adjust per form
TEMPLATE_FILENAME = "TEMPLATE_FORM.pdf"      # TODO: adjust per form

_GRAPH_TOKEN = None
_SP_SITE_ID = None
_SP_DRIVE_ID = None
_SP_LIST_ID = None

# --- Where the filled PDFs get uploaded ----------------------------------
# IEN-Forms is a flat repository of blank templates only — it has no
# per-application subfolders. Filled PDFs go into the existing "TM" folder
# instead, which already has a <reference_id> subfolder per application
# (same structure script 2 uploads its Registration form into).
OUTPUT_FOLDER_NAME = "TM"
OUTPUT_FILE_TYPE = "Education"


# --- Field mapping --------------------------------------------------------
# This form is filled once *per school* the applicant declared
# (ed_school_name_1..7), so only the applicant-level fields that are the
# same on every copy live in FIELD_MAP. sa_nnas_id/sa_application_id come
# from the USER_ID/REFERENCE_ID webhook kwargs, the address fields branch
# on sa_mailing_same_as_current, and the school-specific fields are pulled
# per index — see build_address_fields() / build_school_fields() below.
#
# Left side: key as it appears in the parsed additional_content dict.
# Right side: the exact /AcroForm field name in the PDF (use
# pypdf's reader.get_fields().keys() to find these — see
# `inspect_form_fields()` below, run it once against your template).
FIELD_MAP = {
    "sa_first_name":  "sa_first_name",
    "sa_middle_name": "sa_middle_name",
    "sa_last_name":   "sa_last_name",
    "sa_other_names": "sa_other_names",
    "sa_dob":         "sa_dob",
    "sa_phone_number": "sa_phone_number",
    "sa_email":       "sa_email",
}

# Checkbox fields need explicit on/off values (found via inspect_form_fields()).
# Example: "sa_is_us_citizen": {"payload_key": "sa_is_us_citizen", "on": "/Yes", "off": "/Off"}
CHECKBOX_FIELD_MAP = {
    # "pdf_field_name": {"payload_key": "...", "on": "/Yes", "off": "/Off"},
}

# Per-school fields (indexed 1-7 in the payload, one PDF copy generated per
# non-empty school). Left side: payload key suffix (index gets appended in
# build_school_fields()). Right side: the PDF AcroForm field name.
SCHOOL_FIELD_MAP = {
    "ed_school_name":           "sa_institution_name",
    "ed_program_name":          "sa_program_name",
    "ed_name_used_at_school":   "sa_name_used_at_school",
    "ed_school_merged_name":    "sa_archive_institution_name",
    "ed_program_start_date":    "sa_start_date",
    "ed_program_end_date":      "sa_end_date",
    # sa_school_closed_or_merged is derived from ed_school_operating_<i> and
    # ed_school_merged_<i> (see build_school_fields()), not a direct 1:1 key.
}


# =========================================================
# SHAREPOINT / MICROSOFT GRAPH  (only needed if USE_SHAREPOINT_TEMPLATE)
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


def download_template_from_sharepoint() -> str:
    drive_path = f"{TEMPLATE_FOLDER_NAME}/{TEMPLATE_FILENAME}"

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

    os.makedirs("tmp", exist_ok=True)
    local_path = os.path.join("tmp", TEMPLATE_FILENAME)

    with open(local_path, "wb") as local_file:
        local_file.write(response.content)

    return local_path


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


def upload_to_sharepoint(local_filepath: str, reference_id) -> None:
    """
    Upload the filled PDF to TM/<reference_id>/ (the existing per-application
    folder structure, same one script 2 uploads into) and tag the resulting
    SharePoint list item with FileType=OUTPUT_FILE_TYPE. Requires the
    application's TM subfolder to already exist.
    """
    init_sharepoint_graph()

    application_folder_path = f"{OUTPUT_FOLDER_NAME.rstrip('/')}/{reference_id}"
    target_folder = get_drive_item_by_path(application_folder_path)

    if not target_folder:
        raise RuntimeError(
            f"SharePoint application folder not found. "
            f"Expected library-relative path: {application_folder_path}"
        )

    with open(local_filepath, "rb") as f:
        file_data = f.read()

    sharepoint_filename = f"{reference_id}_EducationForm.pdf"
    upload_path = f"{application_folder_path}/{sharepoint_filename}"

    uploaded_file = graph_request(
        "PUT",
        f"/drives/{_SP_DRIVE_ID}/root:/{encode_graph_path(upload_path)}:/content",
        headers={"Content-Type": "application/pdf"},
        data=file_data,
        expected_statuses=(200, 201),
    )

    patch_driveitem_fields(uploaded_file["id"], {"FileType": OUTPUT_FILE_TYPE})
    print(f"  Uploaded to SharePoint: {upload_path}")


# =========================================================
# INPUT NORMALIZATION
# =========================================================

def parse_additional_content_value(value):
    """
    Parse the SMApply additional_content payload into a dict.

    SMApply free-text fields sometimes contain literal unescaped double
    quotes, which breaks strict JSON parsing. Try strict json.loads first
    (the well-formed case) and fall back to json_repair only if that fails,
    so a single stray quote doesn't drop the whole webhook payload. The
    strict path is preferred because it never guesses; the repair path only
    runs when strict parsing has already failed.
    """
    if value is None:
        raise ValueError("additional_content / ADDITIONAL_CONTENT was not provided.")

    if isinstance(value, dict):
        return value

    if isinstance(value, (bytes, bytearray)):
        value = value.decode("utf-8")

    if isinstance(value, str):
        s = value.strip()
        if not s:
            raise ValueError("additional_content / ADDITIONAL_CONTENT was empty.")

        try:
            parsed = json.loads(s)
        except json.JSONDecodeError as exc:
            print(f"  additional_content is not valid JSON ({exc}) — attempting repair.")
            parsed = json_repair.loads(s)
            print("  additional_content repaired successfully.")

        if isinstance(parsed, str):
            try:
                parsed = json.loads(parsed)
            except json.JSONDecodeError:
                parsed = json_repair.loads(parsed)

        if not isinstance(parsed, dict):
            raise ValueError(
                f"additional_content parsed successfully but was not an object. Got: {type(parsed)}"
            )

        return parsed

    raise TypeError(
        f"additional_content / ADDITIONAL_CONTENT must be str, bytes, bytearray, or dict. Got: {type(value)}"
    )


# =========================================================
# PDF GENERATION
# =========================================================

def inspect_form_fields(pdf_path: str):
    """
    One-time helper: run this against your blank template (e.g. from a
    Windmill preview / test run) to print every AcroForm field name so you
    can populate FIELD_MAP / CHECKBOX_FIELD_MAP above accurately.
    """
    reader = PdfReader(pdf_path)
    fields = reader.get_fields() or {}
    for name, f in fields.items():
        print(f"field_name={name!r} type={f.get('/FT')} value={f.get('/V')!r}")
    return fields


def get_value(parsed: dict, payload_key: str, default=""):
    val = parsed.get(payload_key, default)
    if isinstance(val, str):
        val = val.strip()
    return val if val not in (None, "") else default


def is_yes(val) -> bool:
    return str(val or "").strip().lower() in ("yes", "y", "true", "1")


def build_shared_fields(parsed: dict, user_id, reference_id) -> dict:
    """
    Fields that are identical on every copy of the form regardless of
    which school it's for: applicant identity + the one address the form
    should show (current vs. mailing, per sa_mailing_same_as_current).
    """
    text_fields = {
        pdf_field: get_value(parsed, payload_key)
        for payload_key, pdf_field in FIELD_MAP.items()
    }
    text_fields["sa_nnas_id"] = user_id
    text_fields["sa_application_id"] = reference_id
    text_fields["sd_application_id_left"] = reference_id
    text_fields["sd_application_id_right"] = reference_id

    checkbox_fields = {}
    for pdf_field, cfg in CHECKBOX_FIELD_MAP.items():
        checkbox_fields[pdf_field] = cfg["on"] if is_yes(parsed.get(cfg["payload_key"])) else cfg["off"]

    # The applicant submits two full addresses (current + mailing) plus a
    # flag for whether they match. If they match, source from the current
    # address fields; otherwise use the explicit mailing address fields.
    mailing_same = is_yes(parsed.get("sa_mailing_same_as_current"))
    address_prefix = "sa_applicant_current_address" if mailing_same else "sa_applicant_mailing_address"

    def addr(suffix, default="N/A"):
        return get_value(parsed, f"{address_prefix}_{suffix}", default)

    address_fields = {
        "sa_applicant_address_1":   addr("1"),
        "sa_applicant_address_2":   addr("2"),
        "sa_applicant_city":        addr("city"),
        "sa_applicant_province":    addr("province"),
        "sa_applicant_postal_code": addr("postal_code"),
        "sa_applicant_country":     addr("country"),
    }

    return {**text_fields, **checkbox_fields, **address_fields}


def build_school_fields(parsed: dict, index: int) -> dict:
    """
    Fields specific to one school (index 1-7). sa_school_closed_or_merged
    is a single PDF field derived from two separate submitted answers:
    "is the school currently operating?" and "has it merged with another
    school?" — closed/merged = Yes if either condition is true.
    """
    fields = {
        pdf_field: get_value(parsed, f"{payload_prefix}_{index}")
        for payload_prefix, pdf_field in SCHOOL_FIELD_MAP.items()
    }

    is_operating = is_yes(parsed.get(f"ed_school_operating_{index}"))
    has_merged = is_yes(parsed.get(f"ed_school_merged_{index}"))
    fields["sa_school_closed_or_merged"] = "Yes" if (not is_operating or has_merged) else "No"

    return fields


def create_pdf(template_path: str, output_path: str, parsed: dict, user_id, reference_id) -> bool:
    """
    Build one combined PDF containing a filled copy of the Education form
    for every school the applicant declared (ed_school_name_1..7). Returns
    True if at least one school was found and a PDF was written, False if
    there were no schools to process.
    """
    if not any((parsed.get(f"ed_school_name_{i}") or "").strip() for i in range(1, 8)):
        print("No schools found — skipping PDF creation.")
        return False

    with open(template_path, "rb") as f:
        template_bytes = f.read()

    shared_fields = build_shared_fields(parsed, user_id, reference_id)
    combined_writer = PdfWriter()

    for index in range(1, 8):
        school_name = (parsed.get(f"ed_school_name_{index}") or "").strip()
        if not school_name:
            continue

        reader = PdfReader(io.BytesIO(template_bytes))
        writer = PdfWriter()
        writer.append(reader)

        # Preserve the AcroForm dict and force viewers to regenerate field
        # appearances, otherwise typed values can render blank in some readers.
        if "/AcroForm" in reader.trailer["/Root"]:
            writer._root_object.update({
                NameObject("/AcroForm"): reader.trailer["/Root"]["/AcroForm"]
            })
            writer._root_object["/AcroForm"].update({
                NameObject("/NeedAppearances"): BooleanObject(True)
            })

        form_fields = {**shared_fields, **build_school_fields(parsed, index)}

        for page in writer.pages:
            writer.update_page_form_field_values(page, form_fields)

        buffer = io.BytesIO()
        writer.write(buffer)
        buffer.seek(0)
        for page in PdfReader(buffer).pages:
            combined_writer.add_page(page)

    if len(combined_writer.pages) == 0:
        return False

    with open(output_path, "wb") as f_out:
        combined_writer.write(f_out)

    print(f"Combined Education form PDF created: {output_path} ({len(combined_writer.pages)} page(s))")
    return True


# =========================================================
# MAIN
#
# SMApply webhooks (and Windmill's own re-triggering of the flow) don't
# always send the same casing/keys, so accept the common variants and
# normalize once at the top, the same way the original script does.
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

    if appid is None:
        raise ValueError("appid / id was not provided.")

    if ADDITIONAL_CONTENT is None:
        raise ValueError("ADDITIONAL_CONTENT / additional_content was not provided.")

    parsed = parse_additional_content_value(ADDITIONAL_CONTENT)

    pdf_dir = os.path.join(os.getcwd(), "tmp")
    os.makedirs(pdf_dir, exist_ok=True)
    output_path = os.path.join(pdf_dir, "%s_filled_form.pdf" % REFERENCE_ID)

    if USE_SHAREPOINT_TEMPLATE:
        template_path = download_template_from_sharepoint()
    else:
        # TODO: point this at a bundled/local template if not using SharePoint
        template_path = os.path.join("tmp", TEMPLATE_FILENAME)

    pdf_created = create_pdf(template_path, output_path, parsed, USER_ID, REFERENCE_ID)
    if pdf_created:
        print("Filled PDF created: %s" % output_path)
        upload_to_sharepoint(output_path, REFERENCE_ID)
    else:
        print("No schools found in payload — no PDF was created.")

    # TODO: attach the PDF to the SMApply task too, if this webhook needs
    # that (see script 2's update_smapply_task() for the pattern).

    try:
        os.remove(template_path)
    except OSError:
        pass

    return {
        "pdf_created": pdf_created,
        "reference_id": REFERENCE_ID,
        "appid": appid,
        "output_path": output_path if pdf_created else None,
    }