#requirements:
#pypdf
#reportlab
#pillow
#qrcode
#requests
#wmill

import wmill
import requests
from io import BytesIO
from urllib.parse import quote


# =========================================================
# CONFIGURATION
# =========================================================

BASE = "https://nnas.smapply.io"
PROGRAM_ID = 297312
SMAPPLY_API_TOKEN = wmill.get_variable("u/epham2/SMAPPLY_ACCESS_TOKEN")

SHAREPOINT_LINK_FIELD_ID = 152291

TASK_IDS = ["2075128", "2075132", "2075133"]

TASK_NAME_MAP = {
    "2075128": "Proof of Name Change",
    "2075132": "ID1",
    "2075133": "ID2",
}

TENANT_ID = wmill.get_variable("u/epham2/AZURE_TENANT_ID")
CLIENT_ID = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_ID")
CLIENT_SECRET = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_SECRET")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SHAREPOINT_HOSTNAME = "nnasca.sharepoint.com"
SITE_PATH = "/sites/NNAS-Portal-Sharepoint"

DOCUMENT_LIBRARY_NAME = "Applications"

# Old SharePoint path:
# /sites/NNAS-Portal-Sharepoint/Applications/TM/TEST/
#
# In Graph, this is relative to the Applications document library root.
LIBRARY_BASE_FOLDER = "TM/TEST"

HEADERS = {"Authorization": f"Bearer {SMAPPLY_API_TOKEN}"}

_GRAPH_TOKEN = None
_SP_SITE_ID = None
_SP_DRIVE_ID = None
_SP_LIST_ID = None


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

    print("Connected to SharePoint via Microsoft Graph")
    print("Site ID:", _SP_SITE_ID)
    print("Drive ID:", _SP_DRIVE_ID)
    print("List ID:", _SP_LIST_ID)


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


def create_folder(parent_path: str, folder_name: str):
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
    init_sharepoint_graph()

    parts = [p for p in path.strip("/").split("/") if p]
    current_path = ""
    current_item = None

    for part in parts:
        current_item = create_folder(current_path, part)
        current_path = f"{current_path}/{part}".strip("/")

    return current_item


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


# =========================================================
# SHAREPOINT
# =========================================================

def ensure_application_folder(reference_id: str) -> None:
    """Creates the application folder and 'International Credentials' subfolder if they don't exist."""
    application_folder_path = f"{LIBRARY_BASE_FOLDER}/{reference_id}"

    ensure_folder_path(LIBRARY_BASE_FOLDER)
    ensure_folder_path(application_folder_path)
    ensure_folder_path(f"{application_folder_path}/International Credentials")


def upload_file_to_sharepoint(application_folder_path: str, file_content: BytesIO, file_name: str) -> None:
    """Uploads a single file to the application folder in SharePoint."""
    init_sharepoint_graph()

    target_folder = get_drive_item_by_path(application_folder_path)
    if not target_folder:
        raise RuntimeError(
            f"SharePoint application folder not found. "
            f"Expected library-relative path: {application_folder_path}"
        )

    destination_path = f"{application_folder_path.strip('/')}/{file_name}"

    uploaded_file = graph_request(
        "PUT",
        f"/drives/{_SP_DRIVE_ID}/root:/{encode_graph_path(destination_path)}:/content",
        headers={"Content-Type": "application/octet-stream"},
        data=file_content.getvalue(),
        expected_statuses=(200, 201),
    )

    patch_driveitem_fields(uploaded_file["id"], {"FileType": "Identity"})

    print(f"Uploaded to SharePoint: {destination_path}")


def update_folder_metadata(
    application_id: str,
    reference_id: str,
    user: str,
    additional_content: str,
) -> None:
    """Updates SharePoint folder metadata fields for the application."""
    try:
        application_folder_path = f"{LIBRARY_BASE_FOLDER}/{reference_id}"
        folder_item = get_drive_item_by_path(application_folder_path)

        if not folder_item:
            print(f"No matching folder metadata found for {reference_id}, skipping update.")
            return

        countries = [
            c.strip()
            for c in additional_content.split(",")
            if c.strip()
        ]

        metadata = {
            "ApplicantName": user,
            "NNASID": str(application_id),
            "EducationCountries": countries,
        }

        patch_driveitem_fields(folder_item["id"], metadata)
        print(f"Updated SharePoint metadata for folder: {reference_id}")

        application_link = f"{BASE}/admin/awards/{PROGRAM_ID}/manage/applications/{application_id}/view/"

        try:
            patch_driveitem_fields(
                folder_item["id"],
                {
                    "Link": {
                        "Url": application_link,
                        "Description": "Application Link",
                    }
                },
            )
            print(f"Updated SharePoint Link field for folder: {reference_id}")
        except Exception as link_object_error:
            print("Link object format failed; trying plain URL format.")
            print(str(link_object_error))

            try:
                patch_driveitem_fields(
                    folder_item["id"],
                    {
                        "Link": application_link
                    },
                )
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


def get_renamed_filename(reference_id: str, task_id: str, original_filename: str) -> str:
    """Returns a standardised filename based on reference ID and task ID mapping."""
    base_name = TASK_NAME_MAP.get(task_id, "Unknown")
    extension = original_filename.split(".")[-1]
    return f"{reference_id}_{base_name}.{extension}"


def stream_file_to_sharepoint(
    application_id: str,
    reference_id: str,
    file_info: dict,
) -> None:
    """Downloads a file from SM Apply and uploads it to the application's SharePoint folder."""
    file_url = f"{BASE}{file_info['url']}"
    response = requests.get(file_url, headers=HEADERS, stream=True, timeout=180)

    if response.status_code != 200:
        print(f"Error downloading {file_info['filename']}: {response.status_code}, {response.text}")
        return

    file_content = BytesIO(response.content)
    renamed_file = get_renamed_filename(reference_id, file_info["task_id"], file_info["filename"])

    application_folder_path = f"{LIBRARY_BASE_FOLDER}/{reference_id}"
    upload_file_to_sharepoint(application_folder_path, file_content, renamed_file)


def update_smapply_sharepoint_link(application_id: str, reference_id: str) -> None:
    """Fetches the real SharePoint folder URL and writes it to the SM Apply Website custom field."""
    try:
        application_folder_path = f"{LIBRARY_BASE_FOLDER}/{reference_id}"
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

def main(APPLICATION_ID: str, REFERENCE_ID: str, USER: str, ADDITIONAL_CONTENT: str):
    init_sharepoint_graph()

    ensure_application_folder(REFERENCE_ID)

    for task_id in TASK_IDS:
        print(f"\nProcessing Task ID: {task_id}")
        files = get_task_files(APPLICATION_ID, task_id)

        if not files:
            print(f"No files found for Task ID {task_id}, skipping.")
            continue

        for file_info in files:
            stream_file_to_sharepoint(APPLICATION_ID, REFERENCE_ID, file_info)

    update_folder_metadata(APPLICATION_ID, REFERENCE_ID, USER, ADDITIONAL_CONTENT)
    update_smapply_sharepoint_link(APPLICATION_ID, REFERENCE_ID)