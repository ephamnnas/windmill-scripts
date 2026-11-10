#requirements:
#requests
#wmill

import re
import time
import zipfile
from io import BytesIO
from datetime import datetime
from urllib.parse import urljoin, quote

import requests
import wmill

BASE = "https://nnas.smapply.io"
PROGRAM_ID = 297312

# SharePoint / Microsoft Graph
TENANT_ID = wmill.get_variable("u/epham2/AZURE_TENANT_ID")
CLIENT_ID = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_ID")
CLIENT_SECRET = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_SECRET")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SHAREPOINT_HOSTNAME = "nnasca.sharepoint.com"
SITE_PATH = "/sites/NNAS-Portal-Sharepoint"
DOCUMENT_LIBRARY_NAME = "Applications"

# Old SharePoint path:
# /sites/NNAS-Portal-Sharepoint/Applications/TM/TEST
#
# In Graph, this is relative to the Applications document library root.
FOLDER_RELATIVE_PATH = "TM/TEST"

_GRAPH_TOKEN = None
_SP_SITE_ID = None
_SP_DRIVE_ID = None
_SP_LIST_ID = None

# SM Apply (Windmill vars/secrets)
SMAPPLY_EMAIL = "epham@nnas.ca"
SMAPPLY_PASSWORD = wmill.get_variable("u/epham2/SMAPPLY_EDDIE_PASSWORD")


# -----------------------------
# Microsoft Graph / SharePoint
# -----------------------------

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


# -----------------------------
# Helpers
# -----------------------------

def _get_csrf_from_cookies(session: requests.Session) -> str:
    for c in session.cookies:
        if c.name == "csrftoken":
            return c.value
    raise RuntimeError("No csrftoken cookie found (CSRF cookie missing).")


def _extract_urls(obj) -> list[str]:
    urls = set()

    def walk(x):
        if isinstance(x, dict):
            for v in x.values():
                walk(v)
        elif isinstance(x, list):
            for v in x:
                walk(v)
        elif isinstance(x, str):
            for m in re.findall(r"https?://[^\s\"'<>]+", x):
                urls.add(m)
            for m in re.findall(r"(/(?:dlr|dl)/[A-Za-z0-9_\-\/\.\?\=&]+)", x):
                urls.add(urljoin(BASE, m))

    walk(obj)
    return list(urls)


def _safe_raise_for_status(r: requests.Response, label: str) -> None:
    try:
        r.raise_for_status()
    except Exception:
        snippet = r.text[:500] if hasattr(r, "text") else ""
        raise RuntimeError(
            f"{label} failed: status={r.status_code}, content-type={r.headers.get('content-type')}, snippet={snippet}"
        )


def extract_single_pdf_from_zip(zip_bytes: bytes) -> tuple[bytes, str]:
    if zip_bytes[:2] != b"PK":
        raise RuntimeError("Expected ZIP bytes but magic header does not match PK.")

    if zip_bytes[:4] == b"PK\x05\x06":
        raise RuntimeError("Downloaded ZIP is empty (no entries). The export job produced nothing.")

    with zipfile.ZipFile(BytesIO(zip_bytes), "r") as z:
        pdf_names = [n for n in z.namelist() if n.lower().endswith(".pdf") and not n.endswith("/")]
        if not pdf_names:
            raise RuntimeError(f"ZIP downloaded but no PDFs found. Files={z.namelist()}")
        if len(pdf_names) != 1:
            raise RuntimeError(f"Expected exactly 1 PDF in ZIP, found {len(pdf_names)}: {pdf_names}")

        pdf_name = pdf_names[0]
        return z.read(pdf_name), pdf_name


# -----------------------------
# SM Apply: auth + lookup + export
# -----------------------------

def smapply_login(session: requests.Session) -> None:
    session.get(f"{BASE}/acc/l/", timeout=30)
    csrf = _get_csrf_from_cookies(session)

    headers = {
        "Accept": "*/*",
        "Content-Type": "application/json",
        "Origin": BASE,
        "Referer": f"{BASE}/acc/l/",
        "X-CSRFToken": csrf,
    }
    payload = {"email": SMAPPLY_EMAIL, "password": SMAPPLY_PASSWORD, "nextUrl": ""}

    r = session.post(f"{BASE}/internal/login/", headers=headers, json=payload, timeout=60)
    _safe_raise_for_status(r, "Login")

    if not any(c.name == "sessionid" for c in session.cookies):
        raise RuntimeError("Login did not yield a sessionid cookie. Check creds or auth flow.")


def find_submission_id(session: requests.Session, reference_id: str) -> int:
    url = (
        f"{BASE}/api/applications/"
        f"?program={PROGRAM_ID}"
        f"&fields=reference_id,round_name,round_id,is_awarded,is_declined,is_locked,is_hidden,user_id,initials,full_name,"
        f"number_of_collaborators,organization_id,organization_name,number_of_organization_members,email,id,created_at,"
        f"updated_at,last_submitted_at,status,labels,round"
        f"&page=1&per_page=25&ordering=&base_query=active"
        f"&search={quote(reference_id)}"
    )
    headers = {
        "Accept": "application/smapply.table",
        "Referer": f"{BASE}/admin/awards/{PROGRAM_ID}/manage/",
        "X-Requested-With": "XMLHttpRequest",
    }

    r = session.get(url, headers=headers, timeout=60)
    _safe_raise_for_status(r, "Find submission id")

    data = r.json()
    rows = (data.get("data") or data.get("results") or data.get("rows")) if isinstance(data, dict) else data
    if not rows:
        raise RuntimeError(f"No applications found for reference_id={reference_id}")

    for row in rows:
        if str(row.get("reference_id", "")).strip() == reference_id.strip() and "id" in row:
            return int(row["id"])

    if "id" in rows[0]:
        return int(rows[0]["id"])

    raise RuntimeError(f"Could not find numeric id in /api/applications response for {reference_id}")


def dl_mdl(session: requests.Session, submission_id: int, round_ids: list[int]) -> dict:
    csrf = _get_csrf_from_cookies(session)
    headers = {
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "Content-Type": "application/json; charset=UTF-8",
        "Origin": BASE,
        "Referer": f"{BASE}/admin/awards/{PROGRAM_ID}/manage/applications/{submission_id}/view/",
        "X-Requested-With": "XMLHttpRequest",
        "X-CSRFToken": csrf,
    }
    payload = {"submission_ids": [str(submission_id)], "round_ids": round_ids, "from_round_id": None}
    r = session.post(f"{BASE}/dl/mdl/", headers=headers, json=payload, timeout=60)
    _safe_raise_for_status(r, "dl/mdl")
    return r.json()


def dlr(session: requests.Session, submission_id: int, stage_ids: list[int]) -> dict:
    csrf = _get_csrf_from_cookies(session)
    headers = {
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "Content-Type": "application/json; charset=UTF-8",
        "Origin": BASE,
        "Referer": f"{BASE}/admin/awards/{PROGRAM_ID}/manage/applications/{submission_id}/view/",
        "X-Requested-With": "XMLHttpRequest",
        "X-CSRFToken": csrf,
    }
    payload = {
        "submission_ids": [str(submission_id)],
        "from_round_id": None,
        "stage_ids": stage_ids,
        "download_format": "as_merged_applications",
        "filename_prefix": "application_id",
        "task_selection": True,
        "task_ids": [],
        "include_application_profiles": True,
        "include_application_notes": False,
        "include_application_custom_fields": False,
        "include_applicant_custom_fields": False,
    }
    r = session.post(f"{BASE}/dlr/", headers=headers, json=payload, timeout=60)
    _safe_raise_for_status(r, "dlr")
    return r.json()


def find_download_page_url(dlr_resp: dict) -> str:
    urls = _extract_urls(dlr_resp)
    for u in urls:
        if "/dl/?token=" in u:
            return u

    for k in ("url", "redirect", "redirect_url", "download_url", "href"):
        v = dlr_resp.get(k)
        if isinstance(v, str):
            return urljoin(BASE, v)

    for u in urls:
        if "/dl/" in u or "/dlr/" in u:
            return u

    raise RuntimeError(f"Could not find download page URL in /dlr/ response. keys={list(dlr_resp.keys())}")


def resolve_dl_key_url_once(session: requests.Session, download_page_url: str) -> str:
    if "key=" in download_page_url:
        return download_page_url

    r = session.get(download_page_url, timeout=60)
    _safe_raise_for_status(r, "download page")
    html = r.text

    m = re.search(r"[?&]key=([A-Za-z0-9]+)", html)
    if m:
        return f"{BASE}/dl/?key={m.group(1)}"

    m = re.search(r"(\/dl\/\?key=[A-Za-z0-9]+)", html)
    if m:
        return urljoin(BASE, m.group(1))

    for u in _extract_urls({"html": html}):
        if "/dl/?key=" in u:
            return u

    raise RuntimeError("Could not resolve /dl/?key=... from the download page HTML.")


def wait_for_ready_and_get_zip_url(session: requests.Session, key_url: str, referer_url: str) -> str:
    """
    Polls /dl/?key=... until it returns JSON {success:true, url:"/d/...zip"}.
    Returns absolute ZIP URL.
    """
    headers = {
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "X-Requested-With": "XMLHttpRequest",
        "Referer": referer_url,
    }

    last = None
    for attempt in range(1, 31):
        r = session.get(key_url, headers=headers, timeout=60, allow_redirects=True)
        _safe_raise_for_status(r, "poll /dl/?key")

        ct = (r.headers.get("content-type") or "").lower()
        if "application/json" not in ct:
            snippet = r.text[:200]
            raise RuntimeError(f"Expected JSON from key poll. ct={ct} snippet={snippet}")

        data = r.json()
        last = data

        if data.get("pending") is True:
            print(f"[attempt {attempt}] pending")
            time.sleep(2)
            continue

        if data.get("success") is True and data.get("url"):
            zip_url = urljoin(BASE, data["url"])
            print(f"[attempt {attempt}] ready -> {zip_url} (size={data.get('size')})")
            return zip_url

        if "current" in data and "total" in data:
            print(f"[attempt {attempt}] progress {data.get('current')}/{data.get('total')} comment={data.get('comment')}")
        else:
            print(f"[attempt {attempt}] unexpected json: {data}")

        time.sleep(2)

    raise RuntimeError(f"Timed out waiting for ZIP url from key poll. Last={last}")


def download_zip(session: requests.Session, zip_url: str, referer_url: str) -> bytes:
    """Downloads the ZIP bytes from /d/...zip"""
    headers = {
        "Accept": "*/*",
        "Referer": referer_url,
    }
    r = session.get(zip_url, headers=headers, timeout=300, allow_redirects=True)
    _safe_raise_for_status(r, "download zip")

    blob = r.content
    if not blob[:2] == b"PK":
        snippet = blob[:200].decode("utf-8", errors="replace")
        raise RuntimeError(f"Downloaded file is not a ZIP. bytes={len(blob)} snippet={snippet!r}")

    if blob[:4] == b"PK\x05\x06":
        raise RuntimeError(f"Downloaded ZIP is empty (no entries). The export job produced nothing. zip_url={zip_url}")

    return blob


# -----------------------------
# SharePoint upload
# -----------------------------

def upload_pdf_to_sharepoint(reference_id: str, pdf_bytes: bytes) -> str:
    application_folder_path = f"{FOLDER_RELATIVE_PATH.rstrip('/')}/{reference_id}"
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{reference_id}_{timestamp}.pdf"

    # Retry loop — applicant-docs-tm-test may still be creating the folder
    # when this script runs in parallel, so we poll until it exists.
    max_attempts = 10
    retry_delay = 5  # seconds between attempts

    app_folder = None
    for attempt in range(1, max_attempts + 1):
        try:
            app_folder = get_drive_item_by_path(application_folder_path)
            if app_folder:
                print(f"[attempt {attempt}] Folder found: {application_folder_path}")
                break

            print(f"[attempt {attempt}] Folder not ready yet, retrying in {retry_delay}s...")
            time.sleep(retry_delay)

        except Exception as e:
            print(f"[attempt {attempt}] Folder not ready yet ({e}), retrying in {retry_delay}s...")
            time.sleep(retry_delay)

    if app_folder is None:
        raise RuntimeError(
            f"Folder {application_folder_path} did not exist after {max_attempts} attempts. Giving up."
        )

    upload_path = f"{application_folder_path}/{filename}"

    uploaded_file = graph_request(
        "PUT",
        f"/drives/{_SP_DRIVE_ID}/root:/{encode_graph_path(upload_path)}:/content",
        headers={"Content-Type": "application/pdf"},
        data=pdf_bytes,
        expected_statuses=(200, 201),
    )

    patch_driveitem_fields(uploaded_file["id"], {"FileType": "Other"})

    return filename


# -----------------------------
# Windmill entrypoint
# -----------------------------

def main(REFERENCE_ID: str):
    round_ids = [
        1386304,
        1386305,
        1386306,
        1386307,
        1386308,
        1386309,
        1386310,
        1386311
    ]

    stage_ids = [1386311]

    s = requests.Session()

    smapply_login(s)

    submission_id = find_submission_id(s, REFERENCE_ID)
    print(f"Resolved {REFERENCE_ID} -> submission_id {submission_id}")

    _ = dl_mdl(s, submission_id, round_ids)
    dlr_resp = dlr(s, submission_id, stage_ids)

    download_page_url = find_download_page_url(dlr_resp)
    print(f"Download page URL: {download_page_url}")

    key_url = resolve_dl_key_url_once(s, download_page_url)
    print(f"Resolved key URL (fixed): {key_url}")

    zip_url = wait_for_ready_and_get_zip_url(s, key_url, referer_url=download_page_url)

    zip_blob = download_zip(s, zip_url, referer_url=download_page_url)
    print(f"Downloaded ZIP: {len(zip_blob)} bytes")

    pdf_bytes, pdf_name_in_zip = extract_single_pdf_from_zip(zip_blob)
    print(f"Extracted PDF from ZIP: {pdf_name_in_zip} ({len(pdf_bytes)} bytes)")

    uploaded = upload_pdf_to_sharepoint(REFERENCE_ID, pdf_bytes)
    return {
        "reference_id": REFERENCE_ID,
        "submission_id": submission_id,
        "download_page_url": download_page_url,
        "key_url": key_url,
        "zip_url": zip_url,
        "pdf_name_in_zip": pdf_name_in_zip,
        "uploaded_file": uploaded,
    }