#requirements:
#wmill
#requests

import csv
import io
import re
import wmill
import requests
from html import unescape
from datetime import datetime
from urllib.parse import quote
from typing import Any, Dict, List, Optional, Tuple


# =========================================================
# CONFIGURATION
# =========================================================

# Program IDs, used only to label the returned application (CES vs TM).
# Routing is no longer guessed from the reference ID; it is read off the
# API response, since reference_id already resolves to exactly one program.
PROGRAM_CES = 295865
PROGRAM_TM  = 295185

BASE = "https://nnas.smapply.io"

# Section display order and the field-name pattern for status dropdowns.
# Matches identity_1_status, registration_3_status, education_7_status, etc.
SECTION_ORDER      = ["Identity", "Registration", "Education"]
STATUS_FIELD_RE    = re.compile(r"^(identity|registration|education)_(\d+)_status$", re.IGNORECASE)

# (connect_timeout, read_timeout) so a slow upstream fails fast.
REQUEST_TIMEOUT = (5, 15)


# =========================================================
# ISSUING BODY / DOCUMENT INFO — CONFIG
# ---------------------------------------------------------
# Three different sources per section, because SMApply models each
# differently:
#   - Identity:     ID1/ID2 task endpoint -> "Document Type" dropdown,
#                    decoded via that task's settings. Confirmed.
#   - Registration: the applicant's DocumentIndex.csv, already generated
#                    and sitting in SharePoint per application (see the
#                    IEN-form script's generate_document_index_csv()).
#                    Rows: "Registration - <board>: <license>". Confirmed.
#   - Education:    same DocumentIndex.csv. Rows:
#                    "Education - <school>: <program>". Confirmed.
# =========================================================

# --- SharePoint / Microsoft Graph (same auth pattern as the credential
#     report and IEN-form scripts) ---
TENANT_ID_VAR     = "u/epham2/AZURE_TENANT_ID"
CLIENT_ID_VAR     = "u/epham2/SHAREPOINT_CLIENT_ID"
CLIENT_SECRET_VAR = "u/epham2/SHAREPOINT_CLIENT_SECRET"

GRAPH_BASE            = "https://graph.microsoft.com/v1.0"
SHAREPOINT_HOSTNAME   = "nnasca.sharepoint.com"
SITE_PATH             = "/sites/NNAS-Portal-Sharepoint"
DOCUMENT_LIBRARY_NAME = "Applications"

# Program -> top-level folder under the Applications library. The
# DocumentIndex.csv lives at <FOLDER>/<reference_id>/<reference_id>_DocumentIndex.csv
DOCUMENT_INDEX_FOLDER_BY_PROGRAM: Dict[int, str] = {
    PROGRAM_TM:  "TM",
    PROGRAM_CES: "CES",
}

_GRAPH_TOKEN: Optional[str] = None
_SP_SITE_ID: Optional[str] = None
_SP_DRIVE_ID: Optional[str] = None


def _get_graph_token() -> str:
    global _GRAPH_TOKEN
    if _GRAPH_TOKEN:
        return _GRAPH_TOKEN
    tenant = (wmill.get_variable(TENANT_ID_VAR) or "").strip()
    client_id = (wmill.get_variable(CLIENT_ID_VAR) or "").strip()
    secret = (wmill.get_variable(CLIENT_SECRET_VAR) or "").strip()
    if not tenant or not client_id or not secret:
        raise RuntimeError(
            f"Missing one or more Windmill variables: {TENANT_ID_VAR}, "
            f"{CLIENT_ID_VAR}, {CLIENT_SECRET_VAR}"
        )
    resp = requests.post(
        f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token",
        data={
            "client_id": client_id,
            "client_secret": secret,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        },
        timeout=60,
    )
    if resp.status_code != 200:
        raise RuntimeError(f"Microsoft Graph token request failed {resp.status_code}: {resp.text[:1000]}")
    _GRAPH_TOKEN = resp.json()["access_token"]
    return _GRAPH_TOKEN


def _graph_request(method: str, endpoint: str, expected_statuses=(200,), **kwargs):
    headers = kwargs.pop("headers", {})
    headers["Authorization"] = f"Bearer {_get_graph_token()}"
    resp = requests.request(method, f"{GRAPH_BASE}{endpoint}", headers=headers, timeout=120, **kwargs)
    if resp.status_code not in expected_statuses:
        raise RuntimeError(f"Microsoft Graph {method} {endpoint} failed {resp.status_code}: {resp.text[:2000]}")
    return resp.json() if resp.content else None


def _init_sharepoint_graph() -> None:
    global _SP_SITE_ID, _SP_DRIVE_ID
    if _SP_SITE_ID and _SP_DRIVE_ID:
        return
    site = _graph_request("GET", f"/sites/{SHAREPOINT_HOSTNAME}:{SITE_PATH}")
    _SP_SITE_ID = site["id"]
    drives = _graph_request("GET", f"/sites/{_SP_SITE_ID}/drives")["value"]
    drive = next((d for d in drives if d.get("name") == DOCUMENT_LIBRARY_NAME), None)
    if not drive:
        raise RuntimeError(
            f"Document library '{DOCUMENT_LIBRARY_NAME}' not found. "
            f"Available: {[d.get('name') for d in drives]}"
        )
    _SP_DRIVE_ID = drive["id"]


def _encode_graph_path(path: str) -> str:
    return quote(path.strip("/"), safe="/")


def _get_drive_item_by_path(drive_relative_path: str):
    _init_sharepoint_graph()
    clean_path = drive_relative_path.strip("/")
    if not clean_path:
        return _graph_request("GET", f"/drives/{_SP_DRIVE_ID}/root")
    resp = requests.get(
        f"{GRAPH_BASE}/drives/{_SP_DRIVE_ID}/root:/{_encode_graph_path(clean_path)}",
        headers={"Authorization": f"Bearer {_get_graph_token()}"},
        timeout=120,
    )
    if resp.status_code == 404:
        return None
    if resp.status_code != 200:
        raise RuntimeError(f"Microsoft Graph GET failed {resp.status_code}: {resp.text[:2000]}")
    return resp.json()


def _get_drive_item_content(item_id: str) -> bytes:
    _init_sharepoint_graph()
    resp = requests.get(
        f"{GRAPH_BASE}/drives/{_SP_DRIVE_ID}/items/{item_id}/content",
        headers={"Authorization": f"Bearer {_get_graph_token()}"},
        timeout=120,
        allow_redirects=True,
    )
    if resp.status_code != 200:
        raise RuntimeError(f"Failed downloading SharePoint item content: {resp.status_code}: {resp.text[:1000]}")
    return resp.content


# =========================================================
# HELPERS
# =========================================================

def clean_status(raw) -> str:
    """
    Status values come back HTML-wrapped, e.g. '<ul><li>Ready for Evaluation</li></ul>'.
    Empty slots arrive as '<ul><li></li></ul>' or ''. Strip tags, unescape
    entities, collapse whitespace. Returns '' for anything that is effectively empty.
    """
    if not raw:
        return ""
    text = re.sub(r"<[^>]+>", " ", str(raw))
    text = unescape(text)
    return " ".join(text.split()).strip()


def format_date(raw: str | None) -> str:
    if not raw:
        return "—"
    try:
        dt = datetime.fromisoformat(str(raw).replace("Z", "+00:00"))
        return dt.strftime("%b %-d")   # e.g. "Apr 10"
    except Exception:
        return str(raw)[:10]


def program_label(program_id) -> str:
    if program_id == PROGRAM_TM:
        return "Trumerit"
    if program_id == PROGRAM_CES:
        return "CES"
    return "Unknown"


def _nl(s: Any) -> str:
    return re.sub(r"\s+", " ", (s or "").strip()).lower()


def safe_str(v) -> str:
    if v is None:
        return ""
    if isinstance(v, dict):
        for k in ("label", "value", "name", "text"):
            if k in v and v[k] is not None:
                return str(v[k]).strip()
        return ""
    return str(v).strip()


# =========================================================
# SMAPPLY — APPLICATION LOOKUP
# =========================================================

def fetch_application(headers: dict, reference_id: str) -> dict | None:
    """
    Look up a single application by reference ID using the Bearer-token REST API.
    reference_id is globally unique, so a match returns exactly one result with
    the applicant object, program, and all custom fields inline.
    """
    url = f"{BASE}/api/applications/"
    params = {"reference_id": reference_id}

    try:
        response = requests.get(url, headers=headers, params=params, timeout=REQUEST_TIMEOUT)
    except requests.exceptions.RequestException as exc:
        print(f"Request failed fetching application for '{reference_id}': {exc}")
        return None

    if response.status_code != 200:
        print(f"Error fetching '{reference_id}': {response.status_code}, {response.text[:300]}")
        return None

    data = response.json()
    results = data.get("results") if isinstance(data, dict) else None

    if not results:
        print(f"No application found for reference ID: {reference_id}")
        return None

    # Exact match guard in case the filter is treated as a partial match.
    ref_clean = reference_id.strip().upper()
    for app in results:
        if str(app.get("reference_id", "")).strip().upper() == ref_clean:
            return app

    print(f"Filter returned {len(results)} row(s) but none matched '{ref_clean}' exactly")
    return None


# =========================================================
# IDENTITY — DOCUMENT TYPE (from task endpoint)
# ---------------------------------------------------------
# Identity tasks (ID1 / ID2) carry no issuer/issuing-body field at all —
# only a "Document Type" dropdown and the uploaded file. Confirmed against
# a real task response. So for Identity, "issuing_body" in the output is
# really the decoded Document Type label (e.g. "Passport"), not an issuer.
#
# Task IDs are per program and per slot (ID1, ID2) — SMApply doesn't reuse
# one task across programs. Registration/Education task IDs aren't wired up
# yet; those sections still use their older sources (see further down).
# =========================================================

IDENTITY_TASK_IDS: Dict[int, Dict[int, int]] = {
    PROGRAM_TM:  {1: 2044274, 2: 2044275},
    PROGRAM_CES: {1: 2053926, 2: 2053927},
}
DOCUMENT_TYPE_LABEL_FRAGMENT = "document type"


def fetch_single_task(headers: dict, application_id: int, task_id: int) -> Dict[str, Any] | None:
    """GET /api/applications/{appid}/tasks/{task_id}/ — a single task object,
    not the paginated task list."""
    url = f"{BASE}/api/applications/{application_id}/tasks/{task_id}/"
    try:
        resp = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
    except requests.exceptions.RequestException as exc:
        print(f"Request failed fetching task {task_id} for application {application_id}: {exc}")
        return None
    if not resp.ok:
        print(f"Task fetch failed {resp.status_code} task={task_id} app={application_id}: {resp.text[:300]}")
        return None
    return resp.json()


_TASK_SETTINGS_CACHE: Dict[int, Dict[str, Any]] = {}


def fetch_task_settings(headers: dict, task_id: int) -> Dict[str, Any]:
    """GET /api/tasks/{task_id}/ — field definitions (labels + dropdown
    choices) for a task, cached per task_id since the definition is the
    same across every application that has this task."""
    if task_id in _TASK_SETTINGS_CACHE:
        return _TASK_SETTINGS_CACHE[task_id]
    url = f"{BASE}/api/tasks/{task_id}/"
    try:
        resp = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
    except requests.exceptions.RequestException as exc:
        print(f"Request failed fetching task settings {task_id}: {exc}")
        return {}
    if not resp.ok:
        print(f"Task settings fetch failed {resp.status_code} task={task_id}: {resp.text[:300]}")
        return {}
    settings = resp.json()
    _TASK_SETTINGS_CACHE[task_id] = settings
    return settings


def decode_choice(task_settings: dict, field_key: str, index: Any) -> str:
    """Translate a dropdown response index to its label using task settings'
    fields[field_key].choices, same shape as the sanctions-screening
    script's ChoiceMaps. Falls back to the raw index (stringified) if the
    field or index can't be resolved, so the caller can see something is
    missing rather than silently losing the value."""
    if index is None or index == "" or (isinstance(index, int) and index < 0):
        return ""
    fields = task_settings.get("fields") or {}
    field = fields.get(field_key) or {}
    choices = field.get("choices") or []
    for choice in choices:
        if str(choice.get("index")) == str(index):
            return safe_str(choice.get("label"))
    return f"[option {index}]"   # visible marker that decoding failed, not a silent blank


def extract_identity_document_types(headers: dict, application_id: int,
                                     program_id: int) -> Tuple[Dict[Tuple[str, int], str], dict]:
    """
    Returns ( {("Identity", slot): document_type_label}, debug ).

    Fetches the ID1/ID2 task for this application's program, finds the
    "Document Type" field by label match (safer than trusting the field key
    stays 7alRwLvL1H on every task), and decodes its dropdown response via
    that task's settings.
    """
    slot_task_ids = IDENTITY_TASK_IDS.get(program_id, {})
    issuer_map: Dict[Tuple[str, int], str] = {}
    debug: Dict[str, Any] = {"program_id": program_id, "slots": {}}

    for slot, task_id in slot_task_ids.items():
        slot_debug: Dict[str, Any] = {"task_id": task_id}
        task = fetch_single_task(headers, application_id, task_id)
        if not task:
            slot_debug["error"] = "task fetch failed or task not found"
            debug["slots"][slot] = slot_debug
            continue

        data = task.get("data") or {}
        doc_type_key = None
        doc_type_response = None
        for key, field in data.items():
            if isinstance(field, dict) and DOCUMENT_TYPE_LABEL_FRAGMENT in _nl(field.get("label")):
                doc_type_key = key
                doc_type_response = field.get("response")
                break

        slot_debug["field_labels_seen"] = {
            k: safe_str(v.get("label")) for k, v in data.items() if isinstance(v, dict) and "label" in v
        }

        if doc_type_key is None:
            slot_debug["error"] = "no field with a 'Document Type' label found on this task"
            debug["slots"][slot] = slot_debug
            continue

        settings = fetch_task_settings(headers, task_id)
        label = decode_choice(settings, doc_type_key, doc_type_response)
        slot_debug["field_key"] = doc_type_key
        slot_debug["raw_response"] = doc_type_response
        slot_debug["decoded_label"] = label

        if label and not label.startswith("[option "):
            issuer_map[("Identity", slot)] = label

        debug["slots"][slot] = slot_debug

    return issuer_map, debug


# =========================================================
# REGISTRATION / EDUCATION — from the DocumentIndex.csv in SharePoint
# ---------------------------------------------------------
# The IEN-form script already generates and uploads this CSV per
# application, at:
#   <FOLDER>/<reference_id>/<reference_id>_DocumentIndex.csv
# where FOLDER is "TM" or "CES" (DOCUMENT_INDEX_FOLDER_BY_PROGRAM above).
# Rows look like:
#   1,ECA - Applicant has not indicated provider,,
#   2,Registration - name of the regulatory body: 12345678,,
#   3,Education - name of the school: complete and formal program name,,
# Only "Registration - " and "Education - " rows are used here; each is
# split on the FIRST ": " to separate the issuing name from what follows
# (license number / program name), matching how generate_document_index_csv
# built the line: f"{jname}: {lic}" / f"{school}: {program}".
# =========================================================

DOCUMENT_INDEX_PREFIXES = {
    "Registration": "registration - ",
    "Education":    "education - ",
}


def fetch_document_index_csv(reference_id: str, program_id: int) -> Tuple[Optional[str], dict]:
    """Downloads <reference_id>_DocumentIndex.csv for this application.
    Returns (csv_text or None, debug)."""
    folder = DOCUMENT_INDEX_FOLDER_BY_PROGRAM.get(program_id)
    debug: Dict[str, Any] = {"program_id": program_id, "folder": folder}
    if not folder:
        debug["error"] = f"no DocumentIndex folder configured for program_id {program_id}"
        return None, debug

    folder_path = f"{folder}/{reference_id}"
    filename = f"{reference_id}_DocumentIndex.csv"
    file_path = f"{folder_path}/{filename}"
    debug["file_path"] = file_path

    try:
        item = _get_drive_item_by_path(file_path)
    except Exception as exc:
        debug["error"] = f"SharePoint lookup failed: {str(exc)[:500]}"
        return None, debug

    if not item:
        debug["error"] = "file not found at expected path"
        return None, debug

    try:
        content = _get_drive_item_content(item["id"])
    except Exception as exc:
        debug["error"] = f"SharePoint download failed: {str(exc)[:500]}"
        return None, debug

    debug["found"] = True
    debug["bytes"] = len(content)
    # Uploaded as utf-8-sig (see generate_document_index_csv); decode accordingly.
    return content.decode("utf-8-sig", errors="replace"), debug


def parse_document_index_csv(csv_text: str) -> Dict[str, List[str]]:
    """Returns {"Registration": [name, ...], "Education": [name, ...]}, in
    the row order the CSV lists them (which mirrors the slot order they
    were generated in, since generate_document_index_csv() loops slots
    1-7 in order and only writes populated ones)."""
    out: Dict[str, List[str]] = {"Registration": [], "Education": []}
    reader = csv.DictReader(io.StringIO(csv_text))
    for row in reader:
        title = safe_str(row.get("Document Title"))
        title_nl = _nl(title)
        for section, prefix in DOCUMENT_INDEX_PREFIXES.items():
            if not title_nl.startswith(prefix):
                continue
            rest = title[len(prefix):].strip()   # preserve original casing after the prefix
            # Split on the FIRST ": " only, so a colon inside the program
            # name (rare, but possible) doesn't truncate it.
            name = rest.split(": ", 1)[0].strip()
            if name:
                out[section].append(name)
            break
    return out


def match_csv_entries_to_slots(entries: List[str], slots: List[int]) -> Dict[Tuple[str, int], str]:
    """Positional zip: the Nth CSV entry for a section -> the Nth populated
    slot for that section (both already in ascending order). If counts
    differ, only the overlapping entries are matched — see debug for the
    mismatch rather than a silent misalignment."""
    return dict(zip(slots, entries))


def build_issuing_body_map(headers: dict, application_id: int, program_id: int,
                           reference_id: str,
                           populated_slots: Dict[str, List[int]]) -> Tuple[Dict[Tuple[str, int], str], dict]:
    """
    Returns ( {(section, slot): issuing_body}, debug ).

    - Identity:     ID1/ID2 task, "Document Type" dropdown, decoded via task
                     settings.
    - Registration / Education: DocumentIndex.csv from SharePoint, entries
                     zipped positionally against populated_slots[section]
                     (the already-known, ascending list of slots that have
                     a non-empty status field for that section).
    """
    issuer_map: Dict[Tuple[str, int], str] = {}

    identity_map, identity_dbg = extract_identity_document_types(headers, application_id, program_id)
    issuer_map.update(identity_map)

    csv_text, csv_dbg = fetch_document_index_csv(reference_id, program_id)
    csv_entries = parse_document_index_csv(csv_text) if csv_text else {"Registration": [], "Education": []}

    match_debug: Dict[str, Any] = {}
    for section in ("Registration", "Education"):
        entries = csv_entries.get(section, [])
        slots = populated_slots.get(section, [])
        matched = match_csv_entries_to_slots(entries, slots)   # {slot: name}
        for slot, name in matched.items():
            issuer_map[(section, slot)] = name
        match_debug[section] = {
            "csv_entries_found":  entries,
            "populated_slots":    slots,
            "count_match":        len(entries) == len(slots),
            "matched":            {str(k): v for k, v in matched.items()},
        }

    debug: Dict[str, Any] = {
        "identity":              identity_dbg,
        "document_index_csv":    csv_dbg,
        "registration_education_match": match_debug,
    }

    return issuer_map, debug


# =========================================================
# MAIN
# =========================================================

def main(REFERENCE_ID: str, EMAIL: str) -> dict:
    ref_clean = REFERENCE_ID.strip().upper()

    access_token = wmill.get_variable("u/epham2/SMAPPLY_ACCESS_TOKEN")
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type":  "application/json",
        "Accept":        "application/json",
    }

    # 1. Fetch the application by reference ID (single Bearer-token call)
    application = fetch_application(headers, ref_clean)
    if not application:
        return {"verified": False}

    # 2. Verify the email matches the applicant on record
    applicant_obj   = application.get("applicant") or {}
    applicant_email = (applicant_obj.get("email") or "").lower().strip()
    if applicant_email != EMAIL.lower().strip():
        print(f"Email mismatch for reference ID: {ref_clean}")
        return {"verified": False}

    # 3. Determine program from the response (no letter-based guessing)
    program_id  = (application.get("program") or {}).get("id")
    prog        = program_label(program_id)
    print(f"Resolved {ref_clean} to program {program_id} ({prog})")

    # 4. Serialize status fields by name, skipping unpopulated slots. Build
    #    this BEFORE the issuer map so we know each section's populated slot
    #    order (needed to zip DocumentIndex.csv entries onto the right slots).
    section_docs: dict[str, list] = {}
    for cf in application.get("custom_fields", []):
        name  = cf.get("name") or ""
        match = STATUS_FIELD_RE.match(name.strip())
        if not match:
            continue

        section = match.group(1).capitalize()   # Identity / Registration / Education
        slot    = int(match.group(2))
        status  = clean_status(cf.get("value"))
        if not status:
            continue

        section_docs.setdefault(section, []).append({
            "name":         f"{section} document {slot}",
            "section":      section,
            "slot":         slot,
            "hint":         "",
            "status":       status,
            "issuing_body": "",   # filled in below once the issuer map is built
            "updated":      format_date(application.get("updated_at")),
        })

    populated_slots = {
        section: sorted(d["slot"] for d in docs)
        for section, docs in section_docs.items()
    }

    # 5. Issuing body / document info: Identity from ID1/ID2 task; Registration
    #    and Education from the applicant's DocumentIndex.csv in SharePoint,
    #    zipped onto the populated slots found above.
    application_id = application.get("id")
    issuer_map, issuer_dbg = build_issuing_body_map(
        headers, application_id, program_id, ref_clean, populated_slots
    )

    for docs in section_docs.values():
        for doc in docs:
            doc["issuing_body"] = issuer_map.get((doc["section"], doc["slot"]), "")

    # 6. Emit in fixed section order, sorted by slot within each section
    documents = []
    ordered_sections = SECTION_ORDER + [s for s in section_docs if s not in SECTION_ORDER]
    for section in ordered_sections:
        for doc in sorted(section_docs.get(section, []), key=lambda d: d["slot"]):
            doc.pop("slot", None)   # internal sort key, not needed downstream
            documents.append(doc)

    # 7. Build response
    full_name = f"{applicant_obj.get('first_name', '')} {applicant_obj.get('last_name', '')}".strip() or "Applicant"
    submitted = format_date(application.get("last_submitted_at") or application.get("created_at"))

    print(f"Status lookup successful for {ref_clean} ({prog}) — {len(documents)} document(s) found")

    return {
        "verified":       True,
        "reference_id":   ref_clean,
        "applicant_name": full_name,
        "submitted_date": submitted,
        "documents":      documents,
        "debug": {
            "issuing_body": issuer_dbg,
        },
    }