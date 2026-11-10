#requirements:
#requests
#wmill
#playwright

import re
import time
import zipfile
from io import BytesIO
from datetime import datetime
from urllib.parse import urljoin, quote

import subprocess
import sys

import requests
import wmill
from playwright.sync_api import sync_playwright

BASE = "https://nnas.smapply.io"
PROGRAM_ID = 295865

# SharePoint / Microsoft Graph
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

# SM Apply (Windmill vars/secrets)
SMAPPLY_EMAIL = "epham@nnas.ca"
SMAPPLY_PASSWORD = wmill.get_variable("u/epham2/SMAPPLY_EDDIE_PASSWORD")

# -----------------------------
# Playwright login config
# -----------------------------
# Confirmed against the live site via a standalone test run (two-step
# email -> password form, with a cookie-consent banner that must be
# dismissed first). Same login flow as the TM program's script, since both
# authenticate against the same NNAS SmApply account. Update these if
# SmApply changes the form markup.

# Cookie-consent banner (Ethyca-style) that sits on top of the page on
# first load and must be dismissed before anything else is clickable.
CONSENT_BUTTON_TEXT = "Continue"

# Step 1: email-only screen.
EMAIL_SELECTOR = 'input[name="email"], input[type="email"], #id_email'
EMAIL_NEXT_BUTTON_TEXT = "Next"

# Step 2: password screen, appears only after step 1 submits. The submit
# button has no type="submit" attribute (a React onClick handler), so it
# must be targeted by its visible text.
PASSWORD_SELECTOR = 'input[name="password"], input[type="password"], #id_password'
PASSWORD_SUBMIT_BUTTON_TEXT = "Log in"

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

# Set True locally to watch the browser while testing; keep False in Windmill
# (Windmill workers have no display, so headless=False would fail there).
PLAYWRIGHT_HEADLESS = True


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
    if zip_bytes[:4] != b"PK\x03\x04":
        raise RuntimeError("Expected ZIP bytes but magic header does not match PK\\x03\\x04.")

    with zipfile.ZipFile(BytesIO(zip_bytes), "r") as z:
        pdf_names = [n for n in z.namelist() if n.lower().endswith(".pdf") and not n.endswith("/")]
        if not pdf_names:
            raise RuntimeError(f"ZIP downloaded but no PDFs found. Files={z.namelist()}")
        if len(pdf_names) != 1:
            raise RuntimeError(f"Expected exactly 1 PDF in ZIP, found {len(pdf_names)}: {pdf_names}")

        pdf_name = pdf_names[0]
        return z.read(pdf_name), pdf_name


# -----------------------------
# SM Apply: Playwright-based login
# -----------------------------

def _ensure_chromium_installed() -> None:
    """
    Windmill can spread jobs across multiple workers, and not every worker's
    image/cache has the Chromium binary Playwright needs (the Python package
    installing via #requirements is separate from the actual browser
    download). This installs it on demand if missing -- a no-op (fast) on a
    worker that already has it, a one-time ~150-300MB download on one that
    doesn't.
    """
    try:
        subprocess.run(
            [sys.executable, "-m", "playwright", "install", "--with-deps", "chromium"],
            check=True,
            capture_output=True,
            text=True,
            timeout=300,
        )
    except subprocess.CalledProcessError as e:
        raise RuntimeError(
            "Failed to install Chromium for Playwright. This usually means "
            "the Windmill worker's base image is missing system libraries "
            "Chromium needs, or lacks network access to download it. "
            f"stdout={e.stdout[-2000:]}\nstderr={e.stderr[-2000:]}"
        )


def smapply_login_playwright(email: str, password: str, headless: bool = True,
                              timeout_ms: int = 30000) -> requests.Session:
    """
    Logs into SmApply using a real (Playwright-driven) browser, then returns
    a requests.Session carrying the resulting cookies. Replaces the old
    plain-requests login (email+password POST to /internal/login/), which
    SmApply's bot-detection now blocks for non-browser clients.

    Flow confirmed against the live site (as of this writing, no hard
    captcha challenge was encountered -- see note below):
      1. Dismiss the cookie-consent banner, if present.
      2. Email-only screen: fill email, click "Next".
      3. Password screen: fill password, click "Log in" (matched by text,
         since this button has no type="submit" attribute).
      4. Confirm we actually left the login page (NOT just that a sessionid
         cookie exists -- SmApply sets one for anonymous visitors too).

    Everything downstream (find_submission_id, dl_mdl, dlr,
    wait_for_ready_and_get_zip_url, download_zip) keeps using the returned
    object as a normal requests.Session -- no other function needs to change.

    Raises RuntimeError if the login doesn't complete (e.g. a captcha
    challenge appears, credentials are rejected, or the page markup has
    changed and a selector no longer matches).
    """
    _ensure_chromium_installed()

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=headless,
            args=["--disable-blink-features=AutomationControlled"],
        )
        context = browser.new_context(
            user_agent=USER_AGENT,
            viewport={"width": 1280, "height": 800},
        )
        page = context.new_page()

        try:
            # NOTE: wait_until="networkidle" is deliberately NOT used anywhere
            # in this flow. This page loads continuous analytics/telemetry
            # (Amplitude, consent-management beacons, etc. -- visible as
            # AMP_*, cdp_seg, fides_consent cookies), which can keep sending
            # background requests indefinitely. "networkidle" requires ZERO
            # network activity for 500ms, so on a page like this it can
            # simply never be satisfied, causing intermittent timeouts that
            # have nothing to do with whether the page actually loaded.
            # Waiting for a specific, known element/URL is both faster and
            # far more reliable than waiting for the whole network to go
            # quiet.
            page.goto(f"{BASE}/acc/l/", wait_until="domcontentloaded", timeout=timeout_ms)

            # Dismiss cookie-consent banner if present (non-fatal if absent).
            try:
                consent_btn = page.get_by_role(
                    "button", name=CONSENT_BUTTON_TEXT, exact=False
                )
                if consent_btn.count() > 0:
                    consent_btn.first.click(timeout=5000)
                    page.wait_for_timeout(500)
            except Exception:
                pass

            # Step 1: email + Next. Wait for the email field itself rather
            # than the network going idle -- this is the actual signal that
            # the form is ready to use.
            page.wait_for_selector(EMAIL_SELECTOR, timeout=timeout_ms)
            page.fill(EMAIL_SELECTOR, email, timeout=timeout_ms)
            page.get_by_role("button", name=EMAIL_NEXT_BUTTON_TEXT, exact=False).first.click(
                timeout=timeout_ms
            )

            # Step 2: password + Log in. Wait for the password field to
            # appear (confirms step 1 actually advanced the form) instead of
            # networkidle.
            page.wait_for_selector(PASSWORD_SELECTOR, timeout=timeout_ms)
            page.fill(PASSWORD_SELECTOR, password, timeout=timeout_ms)
            page.get_by_role(
                "button", name=PASSWORD_SUBMIT_BUTTON_TEXT, exact=False
            ).first.click(timeout=timeout_ms)

            # Wait for the URL to actually change away from the login page
            # (the real success signal) rather than networkidle. Falls back
            # to a fixed settle buffer if the URL check itself times out, so
            # we still capture a final screenshot-equivalent state for
            # debugging via the final_url/cookie check below.
            try:
                page.wait_for_url(lambda url: "/acc/l/" not in url, timeout=timeout_ms)
            except Exception:
                pass
            page.wait_for_timeout(2000)

            final_url = page.url
            cookies = context.cookies()
        finally:
            browser.close()

    if "/acc/l/" in final_url:
        cookie_names = [c["name"] for c in cookies]
        raise RuntimeError(
            "Playwright login did not leave the login page. Likely causes: "
            "(a) a captcha challenge blocked it, (b) credentials were "
            "rejected, or (c) the page markup changed and a selector "
            "(CONSENT_BUTTON_TEXT / EMAIL_SELECTOR / EMAIL_NEXT_BUTTON_TEXT / "
            "PASSWORD_SELECTOR / PASSWORD_SUBMIT_BUTTON_TEXT) no longer "
            f"matches the live form. Final URL: {final_url}. "
            f"Cookies present: {cookie_names}"
        )

    session = requests.Session()
    for c in cookies:
        session.cookies.set(
            c["name"], c["value"], domain=c.get("domain"), path=c.get("path", "/")
        )

    return session


# -----------------------------
# SM Apply: lookup + export (unchanged)
# -----------------------------

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
            # Unexpected, but log a snippet to help debug
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

        # progress updates
        if "current" in data and "total" in data:
            print(f"[attempt {attempt}] progress {data.get('current')}/{data.get('total')} comment={data.get('comment')}")
        else:
            print(f"[attempt {attempt}] unexpected json: {data}")

        time.sleep(2)

    raise RuntimeError(f"Timed out waiting for ZIP url from key poll. Last={last}")


def download_zip(session: requests.Session, zip_url: str, referer_url: str) -> bytes:
    """
    Downloads the ZIP bytes from /d/...zip
    """
    headers = {
        "Accept": "*/*",
        "Referer": referer_url,
    }
    r = session.get(zip_url, headers=headers, timeout=300, allow_redirects=True)
    _safe_raise_for_status(r, "download zip")

    blob = r.content
    if blob[:4] != b"PK\x03\x04":
        snippet = blob[:200].decode("utf-8", errors="replace")
        raise RuntimeError(f"Downloaded file is not a ZIP. bytes={len(blob)} snippet={snippet!r}")
    return blob


# -----------------------------
# SharePoint upload
# -----------------------------

def upload_pdf_to_sharepoint(reference_id: str, pdf_bytes: bytes) -> str:
    application_folder_path = f"{FOLDER_RELATIVE_PATH.rstrip('/')}/{reference_id}"
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{reference_id}_{timestamp}.pdf"

    target_folder = get_drive_item_by_path(application_folder_path)
    if not target_folder:
        raise RuntimeError(
            f"SharePoint application folder not found. "
            f"Expected library-relative path: {application_folder_path}"
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
        1369714,
        1369713,
        1369716,
        1369715,
        1369711,
        1369710,
        1369709,
        1369712
    ]

    stage_ids = [1369714]

    s = smapply_login_playwright(SMAPPLY_EMAIL, SMAPPLY_PASSWORD, headless=PLAYWRIGHT_HEADLESS)

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