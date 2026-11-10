#requirements:
#requests
#wmill

import wmill
import io
import csv
import json
import requests
from urllib.parse import quote


# =========================================================
# CONFIGURATION
# =========================================================

# Microsoft Graph / Entra
TENANT_ID = wmill.get_variable("u/epham2/AZURE_TENANT_ID")
CLIENT_ID = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_ID")
CLIENT_SECRET = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_SECRET")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SHAREPOINT_HOSTNAME = "nnasca.sharepoint.com"
SITE_PATH = "/sites/NNAS-Portal-Sharepoint"

DOCUMENT_LIBRARY_NAME = "Applications"

# Old SharePoint server-relative path:
# /sites/NNAS-Portal-Sharepoint/Applications/TM/
#
# In Graph, paths are relative to the document library root.
LIBRARY_BASE_FOLDER = "TM"

_GRAPH_TOKEN = None
_SP_SITE_ID = None
_SP_DRIVE_ID = None


# =========================================================
# MICROSOFT GRAPH / SHAREPOINT
# =========================================================

def _get_graph_token():
    """Gets a Microsoft Graph app-only access token."""
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


def _graph_request(method, endpoint, expected_statuses=(200,), **kwargs):
    """Calls Microsoft Graph and raises a useful error if the request fails."""
    token = _get_graph_token()

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


def _encode_graph_path(path):
    """Encodes a Graph drive path while preserving folder slashes."""
    return quote(path.strip("/"), safe="/")


def _init_sharepoint_graph():
    """Finds the SharePoint site and the Applications document library drive."""
    global _SP_SITE_ID, _SP_DRIVE_ID

    if _SP_SITE_ID and _SP_DRIVE_ID:
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
            f"Could not find SharePoint document library '{DOCUMENT_LIBRARY_NAME}'. "
            f"Available drives: {available}"
        )

    _SP_DRIVE_ID = drive["id"]

    print("Connected to SharePoint via Microsoft Graph")
    print("Site ID:", _SP_SITE_ID)
    print("Drive ID:", _SP_DRIVE_ID)


def _get_drive_item_by_path(drive_relative_path):
    """Returns a drive item by path, or None if it does not exist."""
    _init_sharepoint_graph()

    clean_path = drive_relative_path.strip("/")

    response = requests.get(
        f"{GRAPH_BASE}/drives/{_SP_DRIVE_ID}/root:/{_encode_graph_path(clean_path)}",
        headers={"Authorization": f"Bearer {_get_graph_token()}"},
        timeout=120,
    )

    if response.status_code == 404:
        return None

    if response.status_code != 200:
        raise RuntimeError(
            f"Failed checking SharePoint path.\n"
            f"Path: {drive_relative_path}\n"
            f"Status: {response.status_code}\n"
            f"Response body:\n{response.text[:5000]}"
        )

    return response.json()


def upload_document_index_csv(csv_bytes, reference_id):
    """Overwrite the DocumentIndex.csv for this reference_id using Microsoft Graph."""
    _init_sharepoint_graph()

    filename = "%s_DocumentIndex.csv" % reference_id
    application_folder_path = "%s/%s" % (LIBRARY_BASE_FOLDER.rstrip("/"), reference_id)

    target_folder = _get_drive_item_by_path(application_folder_path)
    if not target_folder:
        raise RuntimeError(
            f"SharePoint folder does not exist for reference_id={reference_id}. "
            f"Expected library-relative path: {application_folder_path}"
        )

    upload_path = "%s/%s" % (application_folder_path.rstrip("/"), filename)

    uploaded = _graph_request(
        "PUT",
        f"/drives/{_SP_DRIVE_ID}/root:/{_encode_graph_path(upload_path)}:/content",
        headers={"Content-Type": "text/csv; charset=utf-8"},
        data=csv_bytes,
        expected_statuses=(200, 201),
    )

    print("  Uploaded CSV: %s (%d bytes)" % (filename, len(csv_bytes)))
    print("  SharePoint URL: %s" % uploaded.get("webUrl"))

    return uploaded


# =========================================================
# CSV GENERATION
# =========================================================

def generate_document_index_csv(parsed, eca_waived):
    """
    Regenerate the document index CSV from scratch.

    eca_waived=True  (Yes) → Registration + Education rows (no ECA)
    eca_waived=False (No)  → ECA + Registration rows (no Education)
    """
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["Number", "Document Title", "Tracking Date", "Uploaded"])
    idx = 0

    # ECA row — included only when NOT waived
    if not eca_waived:
        eca_completed = (parsed.get("ed_eca_completed") or "").strip().lower()
        eca_agency    = (parsed.get("ed_eca_agency") or "").strip()
        if eca_completed == "yes" and eca_agency:
            eca_title = eca_agency
        else:
            eca_title = "Applicant has not indicated provider"
        idx += 1
        w.writerow([idx, "ECA - %s" % eca_title, "", ""])

    # Registration rows — always included
    for i in range(1, 8):
        jname = (parsed.get("sa_jurisdiction_name_%d" % i) or "").strip()
        lic   = (parsed.get("sa_registration_number_%d" % i) or "").strip()
        if not jname and not lic:
            continue
        title = "%s: %s" % (jname, lic) if jname and lic else (jname or lic)
        idx += 1
        w.writerow([idx, "Registration - %s" % title, "", ""])

    # Education rows — included only when waived
    if eca_waived:
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

def main(REFERENCE_ID, USER_ID, USER, appid, ADDITIONAL_CONTENT):
    """
    Same payload shape as the licensure form script. ADDITIONAL_CONTENT
    now includes ed_eca_waived, so no separate API call is needed.
    """
    parsed = json.loads(ADDITIONAL_CONTENT)
    if isinstance(parsed, str):
        parsed = json.loads(parsed)

    eca_waived = str(parsed.get("ed_eca_waived") or "").strip().lower() == "yes"

    csv_bytes = generate_document_index_csv(parsed, eca_waived)
    upload_document_index_csv(csv_bytes, REFERENCE_ID)

    return {
        "reference_id": REFERENCE_ID,
        "appid": appid,
        "eca_waived": eca_waived,
    }