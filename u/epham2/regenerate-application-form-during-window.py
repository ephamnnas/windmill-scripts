# py: 3.12
# requirements:
# requests
# wmill
# playwright
# pypdf

"""
Find and regenerate application forms that were SUBMITTED DURING THE
INCIDENT WINDOW (2026-10-30 21:20 UTC to 2026-11-05 16:20 UTC) in program
295185.

Why: an application submitted in the window had its form PDF generated while
the registration questions used Countries (Custom). Even where the stored
answers were correct (and were restored by the switch back), the PDF shows
the wrong countries. Example: NNAS-T-0000001324 shows Uganda instead of the
United States.

For every application in program 295185 the script:
  1. Finds the newest application form copy in TM/<reference>/
     (<reference>_<YYYYMMDD>_<HHMMSS>.pdf).
  2. If that copy was generated in the window, reads the registration
     countries from it and compares them with what SM Apply stores now
     (read against Countries (Registration)).
  3. Flags it when any country differs. Forms that already match (for
     example, answers in positions 0 to 4, which read the same in both lists)
     are left alone.
  4. Lists any SFTP export folders created in the window for the flagged
     applications, since those packages carry the same wrong form.
  5. With dry_run off, regenerates each flagged form (same export as the
     submission process), checks the new PDF matches SM Apply, and only then
     uploads it as a new timestamped copy. Earlier copies are not changed.

The 28 applications already regenerated on November 6 have a newer copy, so
they are skipped automatically.

Inputs:
  dry_run (default True)  Scan and report only; nothing is generated.
  only                    Comma-separated references or application IDs to
                          limit the run (e.g. "NNAS-T-0000001324").
  max_regenerate          Stop after this many regenerations (0 = no limit).

Results are saved as window_forms_<timestamp>.csv in
Applications/Country list backups. No applicant details are written.
"""

import csv
import io
import re
import time
import traceback
import unicodedata
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import quote

import requests
import wmill
from pypdf import PdfReader

# Same values as the existing script's main().
ROUND_IDS = [1361209, 1362840, 1362866, 1362878, 1367326, 1367683, 1367705, 1367706]
STAGE_IDS = [1361209]

SMAPPLY_BASE = "https://nnas.smapply.io/api"
SMAPPLY_TOKEN_VAR = "u/epham2/SMAPPLY_ACCESS_TOKEN"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SHAREPOINT_HOSTNAME = "nnasca.sharepoint.com"
SITE_PATH = "/sites/NNAS-Portal-Sharepoint"
DOCUMENT_LIBRARY_NAME = "Applications"
LIBRARY_BASE_FOLDER = "TM"
RESULT_FOLDER = "Country list backups"
SUPERSEDED_FOLDER = "Superseded"
FORM_PATTERN = re.compile(r"^NNAS-T-\d+_\d{8}_\d{6}\.pdf$", re.I)
PAUSE_SECONDS = 5

# Every name each Countries (Registration) index can appear as in a PDF:
# Registration name, French name, Countries (Custom) name.
NAMES = {
    0: ['Afghanistan', 'Afghanistan'],
    1: ['Albania', 'Albanie'],
    2: ['Algeria', 'Algérie'],
    3: ['Andorra', 'Andorre'],
    4: ['Angola', 'Angola'],
    5: ['Antarctica', 'Antarctique'],
    6: ['Antigua and Barbuda', 'Antigua-et-Barbuda', 'Antigua & Deps'],
    7: ['Argentina', 'Argentine'],
    8: ['Armenia', 'Arménie'],
    9: ['Australia', 'Australie'],
    10: ['Austria', 'Autriche'],
    11: ['Azerbaijan', 'Azerbaïdjan'],
    12: ['Bahamas', 'Bahamas'],
    13: ['Bahrain', 'Bahreïn'],
    14: ['Bangladesh', 'Bangladesh'],
    15: ['Barbados', 'Barbade'],
    16: ['Belarus', 'Bélarus'],
    17: ['Belgium', 'Belgique'],
    18: ['Belize', 'Belize'],
    19: ['Benin', 'Bénin'],
    20: ['Bermuda', 'Bermudes'],
    21: ['Bhutan', 'Bhoutan'],
    22: ['Bolivia', 'Bolivie'],
    23: ['Bosnia and Herzegovina', 'Bosnie-Herzégovine', 'Bosnia Herzegovina'],
    24: ['Botswana', 'Botswana'],
    25: ['Brazil', 'Brésil'],
    26: ['Brunei', 'Brunéi Darussalam'],
    27: ['Bulgaria', 'Bulgarie'],
    28: ['Burkina Faso', 'Burkina Faso', 'Burkina'],
    29: ['Burma', 'Myanmar', 'Myanmar, {Burma}'],
    30: ['Burundi', 'Burundi'],
    31: ['Cambodia', 'Cambodge'],
    32: ['Cameroon', 'Cameroun'],
    33: ['Canada', 'Canada'],
    34: ['Cape Verde', 'Cabo Verde'],
    35: ['Central African Republic', 'République centrafricaine', 'Central African Rep'],
    36: ['Chad', 'Tchad'],
    37: ['Chile', 'Chili'],
    38: ['China', 'Chine'],
    39: ['Colombia', 'Colombie'],
    40: ['Comoros', 'Comores'],
    41: ['Congo, Democratic Republic', 'Congo (République démocratique)', 'Congo {Democratic Rep}'],
    42: ['Congo, Republic of the', 'Congo (République du)', 'Congo'],
    43: ['Costa Rica', 'Costa Rica'],
    44: ["Cote d'Ivoire", "Côte d'Ivoire", 'Ivory Coast'],
    45: ['Croatia', 'Croatie'],
    46: ['Cuba', 'Cuba'],
    47: ['Cyprus', 'Chypre'],
    48: ['Czech Republic', 'République tchèque'],
    49: ['Denmark', 'Danemark'],
    50: ['Djibouti', 'Djibouti'],
    51: ['Dominica', 'Dominique'],
    52: ['Dominican Republic', 'République dominicaine'],
    53: ['East Timor', 'Timor-Leste'],
    54: ['Ecuador', 'Équateur'],
    55: ['Egypt', 'Égypte'],
    56: ['El Salvador', 'El Salvador'],
    57: ['Equatorial Guinea', 'Guinée équatoriale'],
    58: ['Eritrea', 'Érythrée'],
    59: ['Estonia', 'Estonie'],
    60: ['Ethiopia', 'Éthiopie'],
    61: ['Fiji', 'Fidji'],
    62: ['Finland', 'Finlande'],
    63: ['France', 'France'],
    64: ['Gabon', 'Gabon'],
    65: ['Gambia', 'Gambie'],
    66: ['Georgia', 'Géorgie'],
    67: ['Germany', 'Allemagne'],
    68: ['Ghana', 'Ghana'],
    69: ['Greece', 'Grèce'],
    70: ['Greenland', 'Groenland'],
    71: ['Grenada', 'Grenade'],
    72: ['Guatemala', 'Guatemala'],
    73: ['Guinea', 'Guinée'],
    74: ['Guinea-Bissau', 'Guinée-Bissau'],
    75: ['Guyana', 'Guyana'],
    76: ['Haiti', 'Haïti'],
    77: ['Honduras', 'Honduras'],
    78: ['Hong Kong', 'Hong Kong'],
    79: ['Hungary', 'Hongrie'],
    80: ['Iceland', 'Islande'],
    81: ['India', 'Inde'],
    82: ['Indonesia', 'Indonésie'],
    83: ['Iran', 'Iran'],
    84: ['Iraq', 'Irak'],
    85: ['Ireland', 'Irlande', 'Ireland {Republic}'],
    86: ['Israel', 'Israël'],
    87: ['Italy', 'Italie'],
    88: ['Jamaica', 'Jamaïque'],
    89: ['Japan', 'Japon'],
    90: ['Jordan', 'Jordanie'],
    91: ['Kazakhstan', 'Kazakhstan'],
    92: ['Kenya', 'Kenya'],
    93: ['Kiribati', 'Kiribati'],
    94: ['Korea, North', 'Corée du Nord', 'Korea North'],
    95: ['Korea, South', 'Corée du Sud', 'Korea South'],
    96: ['Kuwait', 'Koweït'],
    97: ['Kyrgyzstan', 'Kirghizistan'],
    98: ['Laos', 'Laos'],
    99: ['Latvia', 'Lettonie'],
    100: ['Lebanon', 'Liban'],
    101: ['Lesotho', 'Lesotho'],
    102: ['Liberia', 'Libéria'],
    103: ['Libya', 'Libye'],
    104: ['Liechtenstein', 'Liechtenstein'],
    105: ['Lithuania', 'Lituanie'],
    106: ['Luxembourg', 'Luxembourg'],
    107: ['Macedonia', 'Macédoine du Nord'],
    108: ['Madagascar', 'Madagascar'],
    109: ['Malawi', 'Malawi'],
    110: ['Malaysia', 'Malaisie'],
    111: ['Maldives', 'Maldives'],
    112: ['Mali', 'Mali'],
    113: ['Malta', 'Malte'],
    114: ['Marshall Islands', 'Îles Marshall'],
    115: ['Mauritania', 'Mauritanie'],
    116: ['Mauritius', 'Maurice'],
    117: ['Mexico', 'Mexique'],
    118: ['Micronesia', 'Micronésie'],
    119: ['Moldova', 'Moldavie'],
    120: ['Mongolia', 'Mongolie'],
    121: ['Morocco', 'Maroc'],
    122: ['Monaco', 'Monaco'],
    123: ['Mozambique', 'Mozambique'],
    124: ['Namibia', 'Namibie'],
    125: ['Nauru', 'Nauru'],
    126: ['Nepal', 'Népal'],
    127: ['Netherlands', 'Pays-Bas'],
    128: ['New Zealand', 'Nouvelle-Zélande'],
    129: ['Nicaragua', 'Nicaragua'],
    130: ['Niger', 'Niger'],
    131: ['Nigeria', 'Nigéria'],
    132: ['Norway', 'Norvège'],
    133: ['Oman', 'Oman'],
    134: ['Pakistan', 'Pakistan'],
    135: ['Panama', 'Panama'],
    136: ['Papua New Guinea', 'Papouasie-Nouvelle-Guinée'],
    137: ['Paraguay', 'Paraguay'],
    138: ['Peru', 'Pérou'],
    139: ['Philippines', 'Philippines'],
    140: ['Poland', 'Pologne'],
    141: ['Portugal', 'Portugal'],
    142: ['Qatar', 'Qatar'],
    143: ['Romania', 'Roumanie'],
    144: ['Russia', 'Russie', 'Russian Federation'],
    145: ['Rwanda', 'Rwanda'],
    146: ['Samoa', 'Samoa'],
    147: ['San Marino', 'Saint-Marin'],
    148: ['Sao Tome', 'Sao Tomé-et-Principe', 'Sao Tome & Principe'],
    149: ['Saudi Arabia', 'Arabie saoudite'],
    150: ['Senegal', 'Sénégal'],
    151: ['Serbia and Montenegro', 'Serbie-et-Monténégro'],
    152: ['Seychelles', 'Seychelles'],
    153: ['Sierra Leone', 'Sierra Leone'],
    154: ['Singapore', 'Singapour'],
    155: ['Slovakia', 'Slovaquie'],
    156: ['Slovenia', 'Slovénie'],
    157: ['Solomon Islands', 'Îles Salomon'],
    158: ['Somalia', 'Somalie'],
    159: ['South Africa', 'Afrique du Sud'],
    160: ['Spain', 'Espagne'],
    161: ['Sri Lanka', 'Sri Lanka'],
    162: ['Sudan', 'Soudan'],
    163: ['Suriname', 'Suriname'],
    164: ['Swaziland', 'Eswatini'],
    165: ['Sweden', 'Suède'],
    166: ['Switzerland', 'Suisse'],
    167: ['Syria', 'Syrie'],
    168: ['Taiwan', 'Taïwan'],
    169: ['Tajikistan', 'Tadjikistan'],
    170: ['Tanzania', 'Tanzanie'],
    171: ['Thailand', 'Thaïlande'],
    172: ['Togo', 'Togo'],
    173: ['Tonga', 'Tonga'],
    174: ['Trinidad and Tobago', 'Trinité-et-Tobago', 'Trinidad & Tobago'],
    175: ['Tunisia', 'Tunisie'],
    176: ['Turkey', 'Türkiye'],
    177: ['Turkmenistan', 'Turkménistan'],
    178: ['Uganda', 'Ouganda'],
    179: ['Ukraine', 'Ukraine'],
    180: ['United Arab Emirates', 'Émirats arabes unis'],
    181: ['United Kingdom', 'Royaume-Uni'],
    182: ['United States', 'États-Unis'],
    183: ['Uruguay', 'Uruguay'],
    184: ['Uzbekistan', 'Ouzbékistan'],
    185: ['Vanuatu', 'Vanuatu'],
    186: ['Venezuela', 'Venezuela'],
    187: ['Vietnam', 'Viêt Nam'],
    188: ['Yemen', 'Yémen'],
    189: ['Zambia', 'Zambie'],
    190: ['Zimbabwe', 'Zimbabwe'],
    191: ['Palestine', 'Palestine'],
    192: ['St. Lucia', 'Sainte-Lucie', 'St Lucia'],
    193: ['Kosovo', 'Kosovo'],
    194: ['Montenegro', 'Monténégro'],
    195: ['Palau', 'Palaos'],
    196: ['Serbia', 'Serbie'],
    197: ['South Sudan', 'Soudan du Sud'],
    198: ['St. Kitts and Nevis', 'Saint-Kitts-et-Nevis', 'St Kitts & Nevis'],
    199: ['St. Vincent and the Grenadines', 'Saint-Vincent-et-les Grenadines', 'Saint Vincent & the Grenadines'],
    200: ['Tuvalu', 'Tuvalu'],
    201: ['Vatican City', 'Cité du Vatican'],
    202: ['Not found', 'Non trouvé'],
}


# =========================================================
# APPLICATION FORM EXPORT
# Copied from the existing application form script so this script is
# self-contained (the Windmill cross-script import resolved to the wrong
# path). Logic is unchanged.
# =========================================================

import subprocess
import sys
import zipfile
from io import BytesIO
from urllib.parse import urljoin

from playwright.sync_api import sync_playwright

BASE = "https://nnas.smapply.io"
PROGRAM_ID = 295185

TENANT_ID = wmill.get_variable("u/epham2/AZURE_TENANT_ID")
CLIENT_ID = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_ID")
CLIENT_SECRET = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_SECRET")

SMAPPLY_EMAIL = "epham@nnas.ca"
SMAPPLY_PASSWORD = wmill.get_variable("u/epham2/SMAPPLY_EDDIE_PASSWORD")

CONSENT_BUTTON_TEXT = "Continue"
EMAIL_SELECTOR = 'input[name="email"], input[type="email"], #id_email'
EMAIL_NEXT_BUTTON_TEXT = "Next"
PASSWORD_SELECTOR = 'input[name="password"], input[type="password"], #id_password'
PASSWORD_SUBMIT_BUTTON_TEXT = "Log in"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)
PLAYWRIGHT_HEADLESS = True

_GRAPH_TOKEN = None
_SP_SITE_ID = None
_SP_DRIVE_ID = None
_SP_LIST_ID = None


def get_graph_token() -> str:
    global _GRAPH_TOKEN
    if _GRAPH_TOKEN:
        return _GRAPH_TOKEN
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
        raise RuntimeError(f"Microsoft Graph token request failed {token_response.status_code}: "
                           f"{token_response.text[:3000]}")
    _GRAPH_TOKEN = token_response.json()["access_token"]
    return _GRAPH_TOKEN


def graph_request(method: str, endpoint: str, expected_statuses=(200,), **kwargs):
    headers = kwargs.pop("headers", {})
    headers["Authorization"] = f"Bearer {get_graph_token()}"
    response = requests.request(method, f"{GRAPH_BASE}{endpoint}", headers=headers, timeout=120, **kwargs)
    if response.status_code not in expected_statuses:
        raise RuntimeError(f"Microsoft Graph request failed. {method} {endpoint} "
                           f"{response.status_code}: {response.text[:5000]}")
    if response.status_code == 204:
        return None
    return response.json() if response.content else None


def encode_graph_path(path: str) -> str:
    return quote(path.strip("/"), safe="/")


def init_sharepoint_graph() -> None:
    global _SP_SITE_ID, _SP_DRIVE_ID, _SP_LIST_ID
    if _SP_SITE_ID and _SP_DRIVE_ID and _SP_LIST_ID:
        return
    site = graph_request("GET", f"/sites/{SHAREPOINT_HOSTNAME}:{SITE_PATH}")
    _SP_SITE_ID = site["id"]
    drives = graph_request("GET", f"/sites/{_SP_SITE_ID}/drives")["value"]
    drive = next((d for d in drives if d.get("name") == DOCUMENT_LIBRARY_NAME), None)
    if not drive:
        raise RuntimeError(f"Could not find document library drive '{DOCUMENT_LIBRARY_NAME}'.")
    _SP_DRIVE_ID = drive["id"]
    lists = graph_request("GET", f"/sites/{_SP_SITE_ID}/lists?$select=id,displayName")["value"]
    sp_list = next((l for l in lists if l.get("displayName") == DOCUMENT_LIBRARY_NAME), None)
    if not sp_list:
        raise RuntimeError(f"Could not find SharePoint list '{DOCUMENT_LIBRARY_NAME}'.")
    _SP_LIST_ID = sp_list["id"]


def get_drive_item_by_path(drive_relative_path: str):
    init_sharepoint_graph()
    response = requests.get(
        f"{GRAPH_BASE}/drives/{_SP_DRIVE_ID}/root:/{encode_graph_path(drive_relative_path)}",
        headers={"Authorization": f"Bearer {get_graph_token()}"}, timeout=120)
    if response.status_code == 404:
        return None
    if response.status_code != 200:
        raise RuntimeError(f"Failed checking SharePoint path {drive_relative_path}: "
                           f"{response.status_code} {response.text[:2000]}")
    return response.json()


def patch_driveitem_fields(drive_item_id: str, fields: dict) -> None:
    init_sharepoint_graph()
    item = graph_request("GET", f"/drives/{_SP_DRIVE_ID}/items/{drive_item_id}?$expand=listItem")
    list_item = item.get("listItem")
    if not list_item or not list_item.get("id"):
        raise RuntimeError(f"Drive item {drive_item_id} has no listItem information.")
    graph_request("PATCH", f"/sites/{_SP_SITE_ID}/lists/{_SP_LIST_ID}/items/{list_item['id']}/fields",
                  headers={"Content-Type": "application/json"}, json=fields, expected_statuses=(200, 204))


def _extract_urls(obj) -> list:
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
        raise RuntimeError(f"{label} failed: status={r.status_code}, "
                           f"content-type={r.headers.get('content-type')}, snippet={r.text[:500]}")


def extract_single_pdf_from_zip(zip_bytes: bytes):
    if zip_bytes[:4] != b"PK\x03\x04":
        raise RuntimeError("Expected ZIP bytes but magic header does not match.")
    with zipfile.ZipFile(BytesIO(zip_bytes), "r") as z:
        pdf_names = [n for n in z.namelist() if n.lower().endswith(".pdf") and not n.endswith("/")]
        if len(pdf_names) != 1:
            raise RuntimeError(f"Expected exactly 1 PDF in ZIP, found {len(pdf_names)}: {pdf_names}")
        return z.read(pdf_names[0]), pdf_names[0]


def _ensure_chromium_installed() -> None:
    try:
        subprocess.run([sys.executable, "-m", "playwright", "install", "--with-deps", "chromium"],
                       check=True, capture_output=True, text=True, timeout=300)
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"Failed to install Chromium for Playwright. "
                           f"stdout={e.stdout[-2000:]}\nstderr={e.stderr[-2000:]}")


def smapply_login_playwright(email: str, password: str, headless: bool = True,
                             timeout_ms: int = 30000) -> requests.Session:
    _ensure_chromium_installed()
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless,
                                    args=["--disable-blink-features=AutomationControlled"])
        context = browser.new_context(user_agent=USER_AGENT, viewport={"width": 1280, "height": 800})
        page = context.new_page()
        try:
            page.goto(f"{BASE}/acc/l/", wait_until="domcontentloaded", timeout=timeout_ms)
            try:
                consent_btn = page.get_by_role("button", name=CONSENT_BUTTON_TEXT, exact=False)
                if consent_btn.count() > 0:
                    consent_btn.first.click(timeout=5000)
                    page.wait_for_timeout(500)
            except Exception:
                pass
            page.wait_for_selector(EMAIL_SELECTOR, timeout=timeout_ms)
            page.fill(EMAIL_SELECTOR, email, timeout=timeout_ms)
            page.get_by_role("button", name=EMAIL_NEXT_BUTTON_TEXT, exact=False).first.click(timeout=timeout_ms)
            page.wait_for_selector(PASSWORD_SELECTOR, timeout=timeout_ms)
            page.fill(PASSWORD_SELECTOR, password, timeout=timeout_ms)
            page.get_by_role("button", name=PASSWORD_SUBMIT_BUTTON_TEXT, exact=False).first.click(timeout=timeout_ms)
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
        raise RuntimeError(f"Playwright login did not leave the login page. Final URL: {final_url}. "
                           f"Cookies present: {[c['name'] for c in cookies]}")
    session = requests.Session()
    for c in cookies:
        session.cookies.set(c["name"], c["value"], domain=c.get("domain"), path=c.get("path", "/"))
    return session


def _get_csrf_from_cookies(session: requests.Session) -> str:
    for c in session.cookies:
        if c.name == "csrftoken":
            return c.value
    raise RuntimeError("No csrftoken cookie found (CSRF cookie missing).")


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


def _ajax_headers(session: requests.Session, submission_id: int) -> dict:
    return {
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "Content-Type": "application/json; charset=UTF-8",
        "Origin": BASE,
        "Referer": f"{BASE}/admin/awards/{PROGRAM_ID}/manage/applications/{submission_id}/view/",
        "X-Requested-With": "XMLHttpRequest",
        "X-CSRFToken": _get_csrf_from_cookies(session),
    }


def dl_mdl(session: requests.Session, submission_id: int, round_ids: list) -> dict:
    payload = {"submission_ids": [str(submission_id)], "round_ids": round_ids, "from_round_id": None}
    r = session.post(f"{BASE}/dl/mdl/", headers=_ajax_headers(session, submission_id), json=payload, timeout=60)
    _safe_raise_for_status(r, "dl/mdl")
    return r.json()


def dlr(session: requests.Session, submission_id: int, stage_ids: list) -> dict:
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
    r = session.post(f"{BASE}/dlr/", headers=_ajax_headers(session, submission_id), json=payload, timeout=60)
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
            raise RuntimeError(f"Expected JSON from key poll. ct={ct} snippet={r.text[:200]}")
        data = r.json()
        last = data
        if data.get("success") is True and data.get("url"):
            return urljoin(BASE, data["url"])
        time.sleep(2)
    raise RuntimeError(f"Timed out waiting for ZIP url from key poll. Last={last}")


def download_zip(session: requests.Session, zip_url: str, referer_url: str) -> bytes:
    r = session.get(zip_url, headers={"Accept": "*/*", "Referer": referer_url}, timeout=300, allow_redirects=True)
    _safe_raise_for_status(r, "download zip")
    blob = r.content
    if blob[:4] != b"PK\x03\x04":
        raise RuntimeError(f"Downloaded file is not a ZIP. bytes={len(blob)}")
    return blob


def upload_pdf_to_sharepoint(reference_id: str, pdf_bytes: bytes) -> str:
    """Uploads to TM/{reference}/{reference}_{timestamp}.pdf and sets FileType to Other,
    exactly as the existing script does."""
    init_sharepoint_graph()
    application_folder_path = f"{LIBRARY_BASE_FOLDER}/{reference_id}"
    if not get_drive_item_by_path(application_folder_path):
        raise RuntimeError(f"SharePoint application folder not found: {application_folder_path}")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{reference_id}_{timestamp}.pdf"
    uploaded_file = graph_request(
        "PUT",
        f"/drives/{_SP_DRIVE_ID}/root:/{encode_graph_path(f'{application_folder_path}/{filename}')}:/content",
        headers={"Content-Type": "application/pdf"},
        data=pdf_bytes,
        expected_statuses=(200, 201),
    )
    patch_driveitem_fields(uploaded_file["id"], {"FileType": "Other"})
    return filename


# =========================================================
# COUNTRY READING (same logic as check_submitted_copies)
# =========================================================

def _norm(s: str) -> str:
    s = unicodedata.normalize("NFKD", s or "")
    s = "".join(c for c in s if not unicodedata.combining(c)).lower()
    s = re.sub(r"[{}(),.'\u2019\-&]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


NAME_TO_INDEX: Dict[str, int] = {}
for _i, _ns in NAMES.items():
    for _n in _ns:
        NAME_TO_INDEX.setdefault(_norm(_n), _i)

HELP_TEXT = re.compile(r"choose a country|choisissez un pays|dropdown|liste deroulante", re.I)
SECTION_HEADING = re.compile(
    r"details for your (current|previous) nursing registration\s*(\d*)|"
    r"(actuelle|anterieure|precedente).{0,40}inscription\s*(\d*)", re.I)


def _section(line: str) -> Optional[int]:
    m = SECTION_HEADING.search(_norm(line))
    if not m:
        return None
    kind = (m.group(1) or m.group(3) or "").lower()
    num = m.group(2) or m.group(4) or ""
    if kind in ("current", "actuelle"):
        return 0
    return int(num) if num.isdigit() else None


def _is_country_question(lines: List[str], i: int) -> bool:
    if HELP_TEXT.search(lines[i]):
        return False
    if not re.search(r"\bcountry\b|\bpays\b", _norm(lines[i])):
        return False
    window = _norm(" ".join(x for x in lines[i:i + 3] if not HELP_TEXT.search(x)))
    return bool(re.search(r"regist|licen|inscri|autoris|permis", window))


def read_countries(pdf_bytes: bytes) -> Dict[int, Optional[int]]:
    """Field position -> country index read from the PDF (None if blank)."""
    text = "\n".join((p.extract_text() or "") for p in PdfReader(io.BytesIO(pdf_bytes)).pages)
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    out: Dict[int, Optional[int]] = {}
    section: Optional[int] = None
    for i in range(len(lines)):
        sec = _section(lines[i])
        if sec is not None:
            section = sec
            continue
        if section is None or not _is_country_question(lines, i) or section in out:
            continue
        found = None
        for j in range(i + 1, min(len(lines), i + 12)):
            if _is_country_question(lines, j) or _section(lines[j]) is not None:
                break
            if not HELP_TEXT.search(lines[j]) and _norm(lines[j]) in NAME_TO_INDEX:
                found = NAME_TO_INDEX[_norm(lines[j])]
                break
        out[section] = found
        section = None
    return out


def check_countries(pdf_bytes: bytes, expected: Dict[int, str]) -> Tuple[bool, str]:
    got = read_countries(pdf_bytes)
    problems = []
    for pos, country in sorted(expected.items()):
        want = NAME_TO_INDEX.get(_norm(country))
        have = got.get(pos)
        label = "Current" if pos == 0 else f"Previous {pos}"
        if have != want:
            shown = NAMES[have][0] if have is not None else "not read"
            problems.append(f"{label}: expected {country}, PDF shows {shown}")
    return (not problems), "; ".join(problems) or "all countries match"


# =========================================================
# SM APPLY API AND SHAREPOINT (read, list, move)
# =========================================================

def app_reference(aid: int) -> str:
    token = wmill.get_variable(SMAPPLY_TOKEN_VAR)
    r = requests.get(f"{SMAPPLY_BASE}/applications/{aid}/",
                     headers={"Authorization": f"Bearer {token}"}, timeout=60)
    return str(r.json().get("title") or "").strip() if r.ok else ""


class SharePoint:
    def __init__(self):
        tenant = wmill.get_variable("u/epham2/AZURE_TENANT_ID").strip()
        tok = requests.post(
            f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token",
            data={
                "client_id": wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_ID").strip(),
                "client_secret": wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_SECRET").strip(),
                "scope": "https://graph.microsoft.com/.default",
                "grant_type": "client_credentials",
            },
            timeout=60,
        )
        tok.raise_for_status()
        self.h = {"Authorization": f"Bearer {tok.json()['access_token']}"}
        site = self._get(f"/sites/{SHAREPOINT_HOSTNAME}:{SITE_PATH}")
        drives = self._get(f"/sites/{site['id']}/drives")["value"]
        self.drive = next(d for d in drives if d.get("name") == DOCUMENT_LIBRARY_NAME)["id"]

    def _get(self, endpoint: str) -> Any:
        r = requests.get(endpoint if endpoint.startswith("http") else f"{GRAPH_BASE}{endpoint}",
                         headers=self.h, timeout=180)
        if r.status_code == 404:
            return None
        r.raise_for_status()
        return r.json()

    def item(self, path: str) -> Optional[Dict[str, Any]]:
        return self._get(f"/drives/{self.drive}/root:/{quote(path.strip('/'), safe='/')}")

    def form_copies(self, folder_id: str) -> List[Dict[str, Any]]:
        data = self._get(f"/drives/{self.drive}/items/{folder_id}/children?$top=200&$select=id,name,file")
        return sorted([c for c in (data or {}).get("value", [])
                       if c.get("file") and FORM_PATTERN.match(c["name"])], key=lambda c: c["name"])

    def ensure_folder(self, parent_id: str, name: str) -> str:
        r = requests.post(f"{GRAPH_BASE}/drives/{self.drive}/items/{parent_id}/children",
                          headers={**self.h, "Content-Type": "application/json"},
                          json={"name": name, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"},
                          timeout=60)
        if r.status_code == 201:
            return r.json()["id"]
        if r.status_code == 409:
            data = self._get(f"/drives/{self.drive}/items/{parent_id}/children?$filter=name eq '{name}'")
            return data["value"][0]["id"]
        r.raise_for_status()
        return ""

    def move(self, item_id: str, folder_id: str) -> None:
        r = requests.patch(f"{GRAPH_BASE}/drives/{self.drive}/items/{item_id}",
                           headers={**self.h, "Content-Type": "application/json"},
                           json={"parentReference": {"id": folder_id}}, timeout=60)
        r.raise_for_status()

    def upload(self, name: str, content: bytes) -> str:
        path = quote(f"{RESULT_FOLDER}/{name}", safe="/")
        r = requests.put(f"{GRAPH_BASE}/drives/{self.drive}/root:/{path}:/content",
                         headers={**self.h, "Content-Type": "application/octet-stream"},
                         data=content, timeout=300)
        r.raise_for_status()
        return r.json().get("webUrl", "")


# =========================================================
# EXPORT (same steps as the existing script's main)
# =========================================================

def export_pdf(session: requests.Session, reference: str) -> bytes:
    submission_id = find_submission_id(session, reference)
    dl_mdl(session, submission_id, ROUND_IDS)
    dlr_resp = dlr(session, submission_id, STAGE_IDS)
    page_url = find_download_page_url(dlr_resp)
    key_url = resolve_dl_key_url_once(session, page_url)
    zip_url = wait_for_ready_and_get_zip_url(session, key_url, referer_url=page_url)
    pdf_bytes, _ = extract_single_pdf_from_zip(download_zip(session, zip_url, referer_url=page_url))
    return pdf_bytes


# =========================================================
# MAIN
# =========================================================



# =========================================================
# WINDOW CHECK
# =========================================================

SWITCH_UTC = datetime(2026, 10, 30, 21, 20, tzinfo=timezone.utc)
SWITCH_BACK_UTC = datetime(2026, 11, 5, 16, 20, tzinfo=timezone.utc)
REGISTRATION_TASK_ID = 2044263
FIELD_POSITIONS = {"GNviXZh7zt": 0, "78nblFoLys": 1, "0gqFv2wuRp": 2, "TqaiTTrVqd": 3,
                   "pkmuzT9P75": 4, "Je7rLX84r5": 5, "pr11MJ2OXs": 6}
EXPORT_RE = re.compile(r"^SFTP_Export_(\d{8})_(\d{6})$", re.I)
STAMP_RE = re.compile(r"_(\d{8})_(\d{6})\.pdf$", re.I)


def _api_headers() -> Dict[str, str]:
    return {"Authorization": f"Bearer {wmill.get_variable(SMAPPLY_TOKEN_VAR)}",
            "Accept": "application/json"}


def list_applications() -> List[Tuple[int, str]]:
    out, url = [], f"{SMAPPLY_BASE}/applications/?program={PROGRAM_ID}&per_page=100"
    h = _api_headers()
    while url:
        r = requests.get(url, headers=h, timeout=120)
        r.raise_for_status()
        data = r.json()
        for a in data.get("results", []):
            out.append((int(a["id"]), str(a.get("title") or "").strip()))
        url = data.get("next")
        time.sleep(0.2)
    return out


def stored_countries(aid: int) -> Dict[int, int]:
    """Field position -> stored index (answered fields only)."""
    r = requests.get(f"{SMAPPLY_BASE}/applications/{aid}/tasks/{REGISTRATION_TASK_ID}/",
                     headers=_api_headers(), timeout=120)
    r.raise_for_status()
    data = (r.json() or {}).get("data") or {}
    out: Dict[int, int] = {}
    for key, pos in FIELD_POSITIONS.items():
        v = data.get(key)
        if isinstance(v, dict):
            v = v.get("value", v.get("response"))
        try:
            v = int(v)
        except (TypeError, ValueError):
            continue
        if v >= 0:
            out[pos] = v
    return out


def _stamp(name: str, pattern=STAMP_RE) -> Optional[datetime]:
    m = pattern.search(name)
    if not m:
        return None
    return datetime.strptime(m.group(1) + m.group(2), "%Y%m%d%H%M%S").replace(tzinfo=timezone.utc)


def in_window(t: Optional[datetime]) -> bool:
    return t is not None and SWITCH_UTC <= t < SWITCH_BACK_UTC


def label(pos: int) -> str:
    return "Current" if pos == 0 else f"Previous {pos}"


def compare(pdf_bytes: bytes, stored: Dict[int, int]) -> List[str]:
    shown = read_countries(pdf_bytes)
    diffs = []
    for pos, idx in sorted(stored.items()):
        have = shown.get(pos)
        if have != idx:
            diffs.append(f"{label(pos)}: form shows "
                         f"{NAMES[have][0] if have is not None else 'nothing readable'}, "
                         f"SM Apply has {NAMES[idx][0]}")
    return diffs


# =========================================================
# MAIN
# =========================================================

def main(dry_run: bool = True, only: str = "", max_regenerate: int = 0):
    sp = SharePoint()
    wanted = {x.strip() for x in only.split(",") if x.strip()}
    apps = [(aid, ref) for aid, ref in list_applications()
            if not wanted or ref in wanted or str(aid) in wanted]

    rows: List[Dict[str, Any]] = []
    flagged: List[Dict[str, Any]] = []
    scanned = no_folder = 0
    for aid, ref in apps:
        if not ref:
            continue
        folder = sp.item(f"{LIBRARY_BASE_FOLDER}/{ref}")
        if not folder:
            no_folder += 1
            continue
        scanned += 1
        data = sp._get(f"/drives/{sp.drive}/items/{folder['id']}/children"
                       "?$top=999&$select=id,name,file,folder")
        children = (data or {}).get("value", [])
        forms = [c for c in children if c.get("file") and FORM_PATTERN.match(c["name"])]
        if not forms:
            continue
        latest = max(forms, key=lambda c: c["name"])
        latest_t = _stamp(latest["name"])
        if not in_window(latest_t):
            continue

        row: Dict[str, Any] = {"application_id": aid, "reference": ref,
                               "latest_form": latest["name"]}
        exports = [c for c in children if c.get("folder") and in_window(
            _stamp(c["name"] + ".pdf", re.compile(r"_(\d{8})_(\d{6})\.pdf$")))
            and EXPORT_RE.match(c["name"])]
        rb = []
        for ex in exports:
            sub = sp._get(f"/drives/{sp.drive}/items/{ex['id']}/children?$select=name,folder")
            rb += [f"{ex['name']}/{s['name']}" for s in (sub or {}).get("value", []) if s.get("folder")]
        row["sftp_exports_in_window"] = "; ".join(rb)
        try:
            stored = stored_countries(aid)
            row["stored"] = "; ".join(f"{label(p)}: {NAMES[i][0]}" for p, i in sorted(stored.items()))
            if not stored:
                row["result"] = "no registration countries answered"
            else:
                diffs = compare(sp_download(sp, latest["id"]), stored)
                row["differences"] = "; ".join(diffs)
                if diffs:
                    row["result"] = "needs regeneration"
                    flagged.append({"aid": aid, "ref": ref, "stored": stored, "row": row})
                else:
                    row["result"] = "form already correct"
        except Exception as exc:
            row["result"] = f"error: {str(exc)[:200]}"
        rows.append(row)

    regenerated = 0
    if not dry_run and flagged:
        session = None
        for n, f in enumerate(flagged, start=1):
            if max_regenerate and regenerated >= max_regenerate:
                f["row"]["result"] = "needs regeneration (not run: max_regenerate reached)"
                continue
            try:
                if session is None:
                    session = smapply_login_playwright(SMAPPLY_EMAIL, SMAPPLY_PASSWORD,
                                                       headless=PLAYWRIGHT_HEADLESS)
                try:
                    pdf = export_pdf(session, f["ref"])
                except Exception:
                    session = smapply_login_playwright(SMAPPLY_EMAIL, SMAPPLY_PASSWORD,
                                                       headless=PLAYWRIGHT_HEADLESS)
                    pdf = export_pdf(session, f["ref"])
                diffs = compare(pdf, f["stored"])
                if diffs:
                    f["row"]["result"] = "NOT uploaded: new form still differs: " + "; ".join(diffs)
                    continue
                f["row"]["uploaded_file"] = upload_pdf_to_sharepoint(f["ref"], pdf)
                f["row"]["result"] = "regenerated"
                regenerated += 1
            except Exception as exc:
                f["row"]["result"] = f"error: {str(exc)[:200]}"
                f["row"]["traceback"] = traceback.format_exc()[-1500:]
            print(f"[{n}/{len(flagged)}] {f['ref']}: {f['row']['result']}")
            time.sleep(PAUSE_SECONDS)

    fields = ["application_id", "reference", "latest_form", "stored", "differences",
              "result", "uploaded_file", "sftp_exports_in_window", "traceback"]
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=fields, extrasaction="ignore")
    w.writeheader()
    w.writerows(rows)
    ts = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    counts: Dict[str, int] = {}
    for r in rows:
        k = r.get("result", "").split(":")[0]
        counts[k] = counts.get(k, 0) + 1
    return {
        "dry_run": dry_run,
        "applications_listed": len(apps),
        "folders_scanned": scanned,
        "no_folder": no_folder,
        "forms_generated_in_window": len(rows),
        "results": counts,
        "needing_regeneration": [f["ref"] for f in flagged],
        "with_sftp_exports_in_window": [f["ref"] for f in flagged if f["row"].get("sftp_exports_in_window")],
        "sharepoint_file": sp.upload(f"window_forms_{ts}.csv", buf.getvalue().encode("utf-8")),
    }


def sp_download(sp: "SharePoint", item_id: str) -> bytes:
    r = requests.get(f"{GRAPH_BASE}/drives/{sp.drive}/items/{item_id}/content",
                     headers=sp.h, timeout=300)
    r.raise_for_status()
    return r.content