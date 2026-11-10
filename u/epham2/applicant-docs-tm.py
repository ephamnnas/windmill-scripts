#requirements:
#pypdf
#reportlab
#pillow
#qrcode
#requests
#wmill

"""
NNAS Intake: ID documents to SharePoint  (unified TM + CES)
===========================================================

Single replacement for the two provider-specific intake scripts. The provider
("TM" or "CES") is read from the webhook's additional_content JSON
(application_info.provider) and selects ALL provider-specific configuration:

    provider  SMApply program  task IDs                      SharePoint base
    TM        295185           2044268 / 2044274 / 2044275   Applications/TM/
    CES       295865           2053922 / 2053926 / 2053927   Applications/CES/

An unknown or missing provider is a HARD ERROR (no default): guessing wrong
would query the wrong program's task IDs (silently zero files) and create the
application folder in the wrong tree. A manual PROVIDER input overrides the
webhook value for ad hoc runs.

Backward compatibility: if ADDITIONAL_CONTENT is NOT the JSON blob (the legacy
flows passed a comma-separated country list), the countries are taken from
that CSV as before and PROVIDER must be supplied explicitly.

Multi-file tasks: a task can carry several uploads (e.g. front/back of an ID).
The first file keeps the canonical name; later files get a page suffix:

    {ref}_ID1.pdf, {ref}_ID1-p2.jpg, {ref}_ID1-p3.pdf

The HYPHEN in -p{n} is load-bearing: the SFTP packager keys ID documents on
the regex _ID(\\d+)\\b, and \\b needs a non-word character after the digits.
_ID1-p2 resolves to ID document 1 (extra page of the same document); an
underscore suffix would fail the boundary and the page would be silently
skipped at packaging time. Requires the companion packager collision guard
(nnas_sftp_multipage_id_patch.py) so same-extension pages get unique numbers
inside the zip.

Special Task Documents (CES only): a single SM Apply task (ID 2095629)
carries four Alternative Assessment upload QUESTIONS (Final Award, Academic
Records/Transcript, Clinical Hours, Certificate of Registration/Licence) as
four separate fields within that one task's "data" object -- unlike the ID
documents above, which are four separate TASKS. Files are grouped by field
key (see get_task_files_grouped_by_field), mapped to a human label via
SPECIAL_TASK_FIELD_LABELS, and uploaded to their own "Special Task
Documents" subfolder as {ref}_{QuestionLabel}.ext (with -p{n} paging for
multi-file uploads within a question). This only runs for CES; TM's config
has no special_task_id, so main() skips it there.

IMPORTANT: the field keys inside SPECIAL_TASK_FIELD_LABELS were captured
from a live task 2095629 API response (application 35668917) and mapped by
JSON field order (matching the order the 4 questions appear in the task:
Final Award, Academic Records/Transcript, Clinical Hours, Certificate of
Registration/Licence). That test run uploaded the SAME file to all 4
fields, so the order-based mapping has NOT been independently verified
field-by-field. Before relying on this in production, run it once more
against an application with 4 DIFFERENT test files -- one per question --
and confirm each field key lands under the correct label using the log
line printed for every field found, e.g.:

    Special Task field discovered: 'gTwNhlsLw2' (sample file: xyz.pdf)

Any field key encountered that ISN'T in SPECIAL_TASK_FIELD_LABELS still
gets uploaded (labeled with its raw field key, with a WARNING in the log)
so nothing is silently dropped if SM Apply ever adds/reorders fields on
this task.
"""

import json
import wmill
import requests
from io import BytesIO
from urllib.parse import quote


# =========================================================
# CONFIGURATION
# =========================================================

BASE = "https://nnas.smapply.io"
SMAPPLY_API_TOKEN = wmill.get_variable("u/epham2/SMAPPLY_ACCESS_TOKEN")

# Same custom field ID in both source scripts (TM and CES programs).
SHAREPOINT_LINK_FIELD_ID = 152291

# Name of the subfolder that holds the Alternative Assessment / Special Task
# Documents uploads, created alongside "International Credentials" and
# "Additional documents" inside each application folder.
SPECIAL_TASK_DOCUMENTS_FOLDER = "Special Task Documents"

# Alternative Assessment pathway: single CES task, 4 upload questions living
# as separate fields inside that one task's "data" object.
SPECIAL_TASK_ID = "2095629"

# Field keys confirmed from a live task 2095629 API response (application
# 35668917) and mapped by JSON field order, which matches the order the
# questions appear in the task: Final Award, Academic Records/Transcript,
# Clinical Hours, Certificate of Registration/Licence. That test used the
# same file for all 4 fields, so the order-based mapping below hasn't been
# independently verified field-by-field -- confirm once with 4 distinct
# test files (one per question) before relying on this in production.
SPECIAL_TASK_FIELD_LABELS = {
    "gTwNhlsLw2": "Final Award",
    "IxxgKmmXo6": "Academic Records-Transcript",
    "QStTD5nX9K": "Clinical Hours",
    "U04O2nfebT": "Certificate of Registration-Licence",
}

# ---- everything provider-specific lives here ----
PROVIDER_CONFIG = {
    "tm": {
        "program_id": 295185,
        "library_base_folder": "TM",
        "task_name_map": {
            "2044268": "Proof of Name Change",
            "2044274": "ID1",
            "2044275": "ID2",
        },
        # TM has no Alternative Assessment / Special Task Documents flow.
        "special_task_id": None,
        "special_task_field_labels": {},
    },
    "ces": {
        "program_id": 295865,
        "library_base_folder": "CES",
        "task_name_map": {
            "2053922": "Proof of Name Change",
            "2053926": "ID1",
            "2053927": "ID2",
        },
        # Alternative Assessment pathway: 1 task (2095629), 4 questions
        # inside it. Empty/absent for providers without this pathway (e.g.
        # TM) -- main() skips special-task processing when this is None.
        "special_task_id": SPECIAL_TASK_ID,
        "special_task_field_labels": SPECIAL_TASK_FIELD_LABELS,
    },
}

# Microsoft Graph / Entra
TENANT_ID = wmill.get_variable("u/epham2/AZURE_TENANT_ID")
CLIENT_ID = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_ID")
CLIENT_SECRET = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_SECRET")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SHAREPOINT_HOSTNAME = "nnasca.sharepoint.com"
SITE_PATH = "/sites/NNAS-Portal-Sharepoint"

DOCUMENT_LIBRARY_NAME = "Applications"

HEADERS = {"Authorization": f"Bearer {SMAPPLY_API_TOKEN}"}

_GRAPH_TOKEN = None
_SP_SITE_ID = None
_SP_DRIVE_ID = None
_SP_LIST_ID = None


# =========================================================
# PROVIDER RESOLUTION
# =========================================================

def parse_additional_content(additional_content):
    """
    Tolerant parse of the webhook additional_content blob (handles the
    double-encoded string form Windmill sometimes passes). Returns a dict,
    or None when the input is not JSON (legacy country-CSV input).
    """
    if not additional_content:
        return None
    if isinstance(additional_content, dict):
        return additional_content
    s = str(additional_content).strip()
    try:
        raw = json.loads(s)
        if isinstance(raw, str):
            raw = json.loads(raw)
        return raw if isinstance(raw, dict) else None
    except (json.JSONDecodeError, ValueError):
        if (s.startswith('"') and s.endswith('"')) or (s.startswith("'") and s.endswith("'")):
            try:
                raw = json.loads(s[1:-1])
                if isinstance(raw, str):
                    raw = json.loads(raw)
                return raw if isinstance(raw, dict) else None
            except Exception:
                return None
        return None


def resolve_provider(provider_override: str, payload: dict | None) -> str:
    """
    Normalized provider key ('tm' / 'ces'). Precedence: explicit override,
    then additional_content.application_info.provider. Unknown or missing
    provider raises: NO default, because a wrong guess queries the wrong
    program's task IDs and creates folders in the wrong SharePoint tree.
    """
    candidate = (provider_override or "").strip().lower()
    if not candidate and payload:
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
        "or pass the PROVIDER input explicitly for manual runs."
    )


def education_countries_from(payload: dict | None, additional_content) -> list[str]:
    """
    EducationCountries metadata values. From the JSON blob:
    personal_info.sa_initial_education_country. From legacy CSV input: the
    comma-split list (original behavior).
    """
    if payload is not None:
        country = str(
            (payload.get("personal_info") or {}).get("sa_initial_education_country") or ""
        ).strip()
        return [country] if country else []
    return [c.strip() for c in str(additional_content or "").split(",") if c.strip()]


# =========================================================
# MICROSOFT GRAPH / SHAREPOINT HELPERS
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
    """Encode a Graph drive path while preserving folder slashes."""
    return quote(path.strip("/"), safe="/")


def init_sharepoint_graph() -> None:
    """Find SharePoint site, Applications document library drive, and backing list."""
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

    print("Connected to SharePoint via Microsoft Graph")
    print("Site ID:", _SP_SITE_ID)
    print("Drive ID:", _SP_DRIVE_ID)
    print("List ID:", _SP_LIST_ID)


def get_drive_item_by_path(drive_relative_path: str):
    """Return a SharePoint drive item by library-relative path, or None if not found."""
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


def create_folder(parent_path: str, folder_name: str):
    """Create a folder under parent_path if it does not already exist."""
    init_sharepoint_graph()

    parent_path = parent_path.strip("/")
    full_path = f"{parent_path}/{folder_name}".strip("/")

    existing = get_drive_item_by_path(full_path)
    if existing:
        print(f"Folder already exists: {full_path}")
        return existing

    if parent_path:
        endpoint = f"/drives/{_SP_DRIVE_ID}/root:/{encode_graph_path(parent_path)}:/children"
    else:
        endpoint = f"/drives/{_SP_DRIVE_ID}/root/children"

    created = graph_request(
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

    print(f"Created folder: {full_path}")
    return created


def ensure_folder_path(path: str):
    """Create each folder segment if missing and return the final drive item."""
    init_sharepoint_graph()

    parts = [p for p in path.strip("/").split("/") if p]
    current_path = ""
    current_item = None

    for part in parts:
        current_item = create_folder(current_path, part)
        current_path = f"{current_path}/{part}".strip("/")

    return current_item


def patch_driveitem_fields(drive_item_id: str, fields: dict) -> None:
    """
    Update SharePoint metadata fields for a driveItem using the backing list item.
    """
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
        headers={
            "Content-Type": "application/json",
            "Prefer": "apiversion=2.1",
        },
        json=fields,
        expected_statuses=(200, 204),
    )


# =========================================================
# SHAREPOINT
# =========================================================

def ensure_application_folder(library_base_folder: str, reference_id: str) -> None:
    """Creates the application folder and required subfolders if they do not exist."""
    application_folder_path = f"{library_base_folder}/{reference_id}"

    ensure_folder_path(library_base_folder)
    ensure_folder_path(application_folder_path)
    ensure_folder_path(f"{application_folder_path}/International Credentials")
    ensure_folder_path(f"{application_folder_path}/Additional documents")
    ensure_folder_path(f"{application_folder_path}/Additional documents/Communications")
    # NOTE: "Special Task Documents" is intentionally NOT created here.
    # Unlike the folders above, it should only exist when the Alternative
    # Assessment task actually has files -- most applicants aren't on that
    # pathway, so creating it unconditionally would leave an empty folder in
    # every application. It's created lazily in
    # process_special_task_documents(), only once files are confirmed.


def upload_file_to_sharepoint(application_folder_path: str, file_content: BytesIO, file_name: str) -> None:
    """Uploads a single file to the application folder in SharePoint via Microsoft Graph."""
    init_sharepoint_graph()

    destination_path = f"{application_folder_path.strip('/')}/{file_name}"

    uploaded_file = graph_request(
        "PUT",
        f"/drives/{_SP_DRIVE_ID}/root:/{encode_graph_path(destination_path)}:/content",
        headers={"Content-Type": "application/octet-stream"},
        data=file_content.getvalue(),
        expected_statuses=(200, 201),
    )

    # FileType drives downstream categorization; a metadata failure should not
    # lose the uploaded file, so it is reported but non-fatal.
    try:
        patch_driveitem_fields(uploaded_file["id"], {"FileType": "Identity"})
    except Exception as e:
        print(f"Uploaded file, but failed to set FileType metadata: {destination_path}")
        print(str(e))

    print(f"Uploaded to SharePoint: {destination_path}")


def update_folder_metadata(
    program_id: int,
    library_base_folder: str,
    application_id: str,
    reference_id: str,
    user: str,
    countries: list[str],
) -> None:
    try:
        application_folder_path = f"{library_base_folder}/{reference_id}"
        folder_item = get_drive_item_by_path(application_folder_path)

        if not folder_item:
            print(f"No matching folder metadata found for {reference_id}, skipping update.")
            return

        for field_name, field_value in [
            ("ApplicantName", user),
            ("NNASID", str(application_id)),
        ]:
            try:
                patch_driveitem_fields(folder_item["id"], {field_name: field_value})
                print(f"Updated {field_name} for folder: {reference_id}")
            except Exception as field_error:
                print(f"Failed to update {field_name} for folder {reference_id}: {field_error}")

        print(f"Metadata update pass complete for folder: {reference_id}")

        application_link = f"{BASE}/admin/awards/{program_id}/manage/applications/{application_id}/view/"

        try:
            patch_driveitem_fields(
                folder_item["id"],
                {"Link": {"Url": application_link, "Description": "Application Link"}},
            )
            print(f"Updated SharePoint Link field for folder: {reference_id}")
        except Exception as link_object_error:
            print("Link object format failed; trying plain URL format.")
            print(str(link_object_error))
            try:
                patch_driveitem_fields(folder_item["id"], {"Link": application_link})
                print(f"Updated SharePoint Link field using plain URL for folder: {reference_id}")
            except Exception as link_string_error:
                print("Core metadata updated, but the Link field failed.")
                print(str(link_string_error))

    except Exception as e:
        print(f"Error updating SharePoint metadata for {reference_id}: {str(e)}")


# =========================================================
# SMAPPLY
# =========================================================

def get_task_files(application_id: str, task_id: str) -> list[dict]:
    """Retrieves all file entries for a given task from the SM Apply API."""
    url = f"{BASE}/api/applications/{application_id}/tasks/{task_id}/"
    response = requests.get(url, headers=HEADERS, timeout=120)

    if response.status_code != 200:
        print(f"Error retrieving task {task_id}: {response.status_code}, {response.text}")
        return []

    task_data = response.json()
    files = []

    for value in task_data.get("data", {}).values():
        if isinstance(value, list):
            for entry in value:
                if isinstance(entry, dict) and "filename" in entry and "url" in entry:
                    files.append({
                        "id": entry["id"],
                        "filename": entry["filename"],
                        "url": entry["url"],
                        "task_id": task_id,
                    })

    return files


def get_task_files_grouped_by_field(application_id: str, task_id: str) -> dict:
    """
    Like get_task_files, but for tasks that carry MULTIPLE upload QUESTIONS
    as separate fields within one task's "data" object (e.g. the Alternative
    Assessment task, which has 4 distinct upload questions in a single task
    rather than 4 separate tasks). Returns {field_key: [file_dict, ...]},
    preserving which field each file came from so it can be labeled and
    filed correctly, instead of get_task_files' flat list which loses that
    distinction.

    Logs every field key found, with a sample filename, so real field keys
    can be copied into SPECIAL_TASK_FIELD_LABELS on first run against a live
    application.
    """
    url = f"{BASE}/api/applications/{application_id}/tasks/{task_id}/"
    response = requests.get(url, headers=HEADERS, timeout=120)

    if response.status_code != 200:
        print(f"Error retrieving task {task_id}: {response.status_code}, {response.text}")
        return {}

    task_data = response.json()
    grouped: dict = {}

    for field_key, value in task_data.get("data", {}).items():
        if not isinstance(value, list):
            continue

        field_files = []
        for entry in value:
            if isinstance(entry, dict) and "filename" in entry and "url" in entry:
                field_files.append({
                    "id": entry["id"],
                    "filename": entry["filename"],
                    "url": entry["url"],
                    "task_id": task_id,
                })

        if field_files:
            grouped[field_key] = field_files
            print(f"Special Task field discovered: '{field_key}' "
                  f"(sample file: {field_files[0]['filename']})")

    return grouped


def _file_extension(original_filename: str) -> str:
    """
    Extension WITHOUT the dot, lowercased. Safe for filenames with no dot
    (falls back to 'pdf' rather than treating the whole name as an extension,
    which is what `split('.')[-1]` did).
    """
    name = (original_filename or "").strip()
    if "." in name:
        ext = name.rsplit(".", 1)[1].strip().lower()
        if ext and len(ext) <= 8:
            return ext
    return "pdf"


def get_renamed_filename(task_name_map: dict, reference_id: str, task_id: str,
                         original_filename: str, index: int = 1) -> str:
    """
    Standardised filename based on reference ID and task ID mapping.

    A task can carry MULTIPLE uploads (e.g. front and back of an ID as two
    photos). The first file keeps the canonical name so single-file behavior
    is unchanged; subsequent files get a page suffix BEFORE the extension:

        index 1 -> {ref}_ID1.pdf
        index 2 -> {ref}_ID1-p2.jpg
        index 3 -> {ref}_ID1-p3.pdf

    The HYPHEN in `-p{n}` is load-bearing: downstream, the SFTP packager keys
    ID documents on the regex `_ID(\\d+)\\b`, and `\\b` needs a non-word char
    after the digits. `_ID1-p2` matches as ID document 1 (same logical
    document, extra page); an underscore suffix (`_ID1_p2`) would FAIL the
    boundary and the page would be silently skipped by the packager.
    """
    base_name = task_name_map.get(task_id, "Unknown")
    extension = _file_extension(original_filename)
    if index <= 1:
        return f"{reference_id}_{base_name}.{extension}"
    return f"{reference_id}_{base_name}-p{index}.{extension}"


def stream_file_to_sharepoint(
    task_name_map: dict,
    library_base_folder: str,
    application_id: str,
    reference_id: str,
    file_info: dict,
    index: int = 1,
    subfolder: str | None = None,
) -> bool:
    """
    Downloads one file from SM Apply and uploads it to the application's
    SharePoint folder (or a named subfolder of it, e.g. "Special Task
    Documents") under its standardised (and, for multi-file tasks,
    page-numbered) name. Returns True on success.
    """
    file_url = f"{BASE}{file_info['url']}"
    response = requests.get(file_url, headers=HEADERS, stream=True, timeout=180)

    if response.status_code != 200:
        print(f"Error downloading {file_info['filename']}: {response.status_code}, {response.text}")
        return False

    file_content = BytesIO(response.content)
    renamed_file = get_renamed_filename(
        task_name_map, reference_id, file_info["task_id"], file_info["filename"], index
    )

    application_folder_path = f"{library_base_folder}/{reference_id}"
    if subfolder:
        application_folder_path = f"{application_folder_path}/{subfolder}"

    upload_file_to_sharepoint(application_folder_path, file_content, renamed_file)
    return True


def process_special_task_documents(
    special_task_id: str | None,
    special_task_field_labels: dict,
    library_base_folder: str,
    application_id: str,
    reference_id: str,
) -> dict:
    """
    Pulls every file from every question (field) inside the single
    Alternative Assessment task and uploads them to the application's
    "Special Task Documents" subfolder, named {ref}_{QuestionLabel}.ext
    (with -p{n} paging for multi-file uploads within a single question,
    same convention as get_renamed_filename/stream_file_to_sharepoint).

    Unlike the ID documents (4 separate tasks), these 4 questions are 4
    fields inside ONE task, so this groups by field key first
    (get_task_files_grouped_by_field) rather than calling get_task_files
    once per question.

    Any field found in the task that ISN'T in special_task_field_labels
    still gets uploaded -- labeled with its raw field key -- so a
    not-yet-mapped field key never silently loses files; it just shows up
    under an unfriendly name until the mapping is filled in.

    A no-op (returns {}) when special_task_id is falsy, so this is safe to
    call unconditionally for providers that don't have this task (e.g. TM).
    """
    uploaded_summary: dict = {}

    if not special_task_id:
        return uploaded_summary

    print(f"\nProcessing Special Task Documents task: {special_task_id}")
    fields = get_task_files_grouped_by_field(application_id, special_task_id)

    if not fields:
        print(f"No files found in Special Task Documents task {special_task_id}, "
              f"skipping (folder will not be created).")
        return uploaded_summary

    # Files were found -- only now create the subfolder, since most
    # applicants aren't on the Alternative Assessment pathway and shouldn't
    # get an empty "Special Task Documents" folder.
    application_folder_path = f"{library_base_folder}/{reference_id}"
    ensure_folder_path(f"{application_folder_path}/{SPECIAL_TASK_DOCUMENTS_FOLDER}")

    for field_key, files in fields.items():
        question_label = special_task_field_labels.get(field_key)
        if not question_label:
            print(f"WARNING: field '{field_key}' has no entry in "
                  f"SPECIAL_TASK_FIELD_LABELS; using raw field key as the "
                  f"filename label. Update the mapping once the correct "
                  f"question name is confirmed.")
            question_label = field_key

        if len(files) > 1:
            print(f"Field '{field_key}' ({question_label}) has {len(files)} uploads; "
                  f"numbering pages -p2...-p{len(files)}.")

        # get_renamed_filename looks up base_name by task_id, but here we
        # want it keyed by field/question instead -- build a one-entry map
        # so the existing helper can be reused without changing its
        # (task_id -> label) contract used elsewhere for the ID documents.
        field_as_task_name_map = {special_task_id: question_label}

        uploaded = 0
        for index, file_info in enumerate(files, start=1):
            if stream_file_to_sharepoint(
                field_as_task_name_map, library_base_folder,
                application_id, reference_id,
                file_info, index,
                subfolder=SPECIAL_TASK_DOCUMENTS_FOLDER,
            ):
                uploaded += 1

        uploaded_summary[question_label] = f"{uploaded}/{len(files)}"

    print(f"\nSpecial Task Documents upload summary: {uploaded_summary}")
    return uploaded_summary


def update_smapply_sharepoint_link(library_base_folder: str,
                                   application_id: str, reference_id: str) -> None:
    """Fetches the real SharePoint folder URL and writes it to the SM Apply Website custom field."""
    try:
        application_folder_path = f"{library_base_folder}/{reference_id}"
        folder = get_drive_item_by_path(application_folder_path)

        if not folder:
            print(f"Error fetching SharePoint folder URL for {reference_id}: folder not found")
            return

        sharepoint_link = folder.get("webUrl")

        if not sharepoint_link:
            print(f"Error fetching SharePoint folder URL for {reference_id}: no webUrl returned")
            return

    except Exception as e:
        print(f"Error fetching SharePoint folder URL for {reference_id}: {str(e)}")
        return

    # Custom field values are written via PATCH on the application itself
    url = f"{BASE}/api/applications/{application_id}/"
    payload = {
        "custom_fields": [
            {"id": SHAREPOINT_LINK_FIELD_ID, "value": sharepoint_link}
        ]
    }

    response = requests.patch(url, json=payload, headers=HEADERS, timeout=120)

    if response.status_code in (200, 201, 204):
        print(f"SM Apply SharePoint link updated: {sharepoint_link}")
    else:
        print(f"Error updating SM Apply custom field [{response.status_code}]: {response.text}")


# =========================================================
# MAIN
# =========================================================

def main(APPLICATION_ID: str, REFERENCE_ID: str, USER: str,
         ADDITIONAL_CONTENT: str, PROVIDER: str = ""):
    """
    Same four inputs as the old TM/CES scripts, plus an optional PROVIDER
    override for manual runs. When ADDITIONAL_CONTENT is the webhook JSON
    blob, provider comes from application_info.provider and the education
    country from personal_info.sa_initial_education_country; when it is the
    legacy country CSV, that CSV populates EducationCountries and PROVIDER
    must be given explicitly.
    """
    payload = parse_additional_content(ADDITIONAL_CONTENT)
    provider = resolve_provider(PROVIDER, payload)
    cfg = PROVIDER_CONFIG[provider]

    program_id = cfg["program_id"]
    library_base_folder = cfg["library_base_folder"]
    task_name_map = cfg["task_name_map"]
    special_task_id = cfg.get("special_task_id")
    special_task_field_labels = cfg.get("special_task_field_labels", {})
    task_ids = list(task_name_map.keys())

    print(f"Provider: {provider.upper()}  "
          f"(program {program_id}, base folder Applications/{library_base_folder}/)")

    init_sharepoint_graph()

    ensure_application_folder(library_base_folder, REFERENCE_ID)

    uploaded_summary: dict = {}

    for task_id in task_ids:
        print(f"\nProcessing Task ID: {task_id} ({task_name_map.get(task_id, 'Unknown')})")
        files = get_task_files(APPLICATION_ID, task_id)

        if not files:
            print(f"No files found for Task ID {task_id}, skipping.")
            continue

        if len(files) > 1:
            print(f"Task {task_id} has {len(files)} uploads; "
                  f"numbering pages -p2...-p{len(files)}.")

        uploaded = 0
        # Index within THIS task, in the order SM Apply returns the entries.
        # Numbering restarts per task, so ID1 pages never collide with ID2's.
        for index, file_info in enumerate(files, start=1):
            if stream_file_to_sharepoint(task_name_map, library_base_folder,
                                         APPLICATION_ID, REFERENCE_ID,
                                         file_info, index):
                uploaded += 1

        uploaded_summary[task_name_map.get(task_id, task_id)] = f"{uploaded}/{len(files)}"

    print(f"\nUpload summary: {uploaded_summary}")

    # Alternative Assessment / Special Task Documents (CES only; no-op
    # elsewhere since special_task_id is None for other providers).
    special_task_summary = process_special_task_documents(
        special_task_id, special_task_field_labels,
        library_base_folder, APPLICATION_ID, REFERENCE_ID
    )
    if special_task_summary:
        uploaded_summary.update(
            {f"[Special] {label}": result for label, result in special_task_summary.items()}
        )

    countries = education_countries_from(payload, ADDITIONAL_CONTENT)
    update_folder_metadata(program_id, library_base_folder,
                           APPLICATION_ID, REFERENCE_ID, USER, countries)
    update_smapply_sharepoint_link(library_base_folder, APPLICATION_ID, REFERENCE_ID)

    return {
        "provider": provider.upper(),
        "program_id": program_id,
        "base_folder": f"Applications/{library_base_folder}/{REFERENCE_ID}",
        "uploads": uploaded_summary,
        "education_countries": countries,
    }