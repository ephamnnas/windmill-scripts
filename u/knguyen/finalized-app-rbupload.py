# py: 3.12
# requirements:
# wmill
# requests

"""
NNAS RB Portal Upload  (self-contained Windmill flow step)
==========================================================

Runs AFTER nnas_sftp_package.py in the same flow. For one application it pushes
each regulator's report to the RB Portal (WordPress, rbportal.nnas.ca):

  1. Resolves the SharePoint application folder from the same webhook fields the
     SFTP step uses (reference_id + provider).
  2. Re-reads "credential reports/<ABBREV>/*_latest.pdf" from SharePoint -- the
     previous step's /tmp files do not survive into this container, so the
     canonical SharePoint copies are the source. Only the most-final PDF per
     (regulator, designation) is used; the redacted "Applicant Copy" and the
     editable .docx are ignored.
  3. For each (regulator x designation) leg:
       - binds the EXISTING RB record created at submission stage by
         nnas_rb_record_upsert.py (application_id + regulatory_body), setting
         status + issued_date, and
       - uploads that leg's full report PDF to the record so the portal serves a
         unique PDF per college / designation.

RECORD IDENTITY (this is what keeps it to one row per leg)
----------------------------------------------------------
The submission step addresses rows by application_id + canonical
regulatory_body, and writes discipline as a long label ("Registered Nurse").
This step uses the SAME alias resolver (_SINGLE / _BY_DISC / resolve_rb) and the
SAME discipline_label(), so the finder binds the existing row instead of
creating a second one. Do not reintroduce a discipline-keyed match: colleges
that serve two designations (BCCNM, NSCN, NL, PEI) already resolve to distinct
canonical names per discipline, so regulatory_body alone is unique per leg.

Every field is pruned of empty strings before posting. A missing date_created
must not blank the order_date the submission stage already wrote.

Independent of SFTP success: the portal is a separate delivery channel, so a
failed SFTP leg does not block (or get blocked by) the portal push. Each leg is
best-effort; one failure does not abort the others.

Entrypoint binds the same flow_input.* as the SFTP step (reference_id,
additional_content, user_id, dry_run), so the step inputs are wired identically.

Requires:
  - RBP Upload API plugin v1.6.0+ (composite finder + post_id report route +
    report_file meta).
  - RB Portal plugin v3.5.0+ (download resolves report_file meta).
  - The Windmill worker egress IP on the plugin allowlist (already true; the
    scraper posts from here).
"""

import os
import re
import json
import time
import shutil
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional, Tuple

import wmill
import requests
from urllib.parse import quote as _url_quote


# =========================================================
# SHAREPOINT / MICROSOFT GRAPH CONFIG  (matches nnas_sftp_package.py)
# =========================================================
SITE_PREFIX   = "/sites/NNAS-Portal-Sharepoint"

TENANT_ID     = wmill.get_variable("u/epham2/AZURE_TENANT_ID")
CLIENT_ID     = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_ID")
CLIENT_SECRET = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_SECRET")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SHAREPOINT_HOSTNAME = "nnasca.sharepoint.com"
SITE_PATH = "/sites/NNAS-Portal-Sharepoint"
DOCUMENT_LIBRARY_NAME = "Applications"

_GRAPH_TOKEN = None
_SP_SITE_ID = None
_SP_DRIVE_ID = None
_SP_LIST_ID = None

CREDENTIAL_REPORTS_SUBFOLDER = "credential reports"

# Provider (application_info.provider) selects the SharePoint base path.
PROVIDER_BASE_TEMPLATE = {
    "tm":  "applications/tm/{app_id}",
    "ces": "applications/ces/{app_id}",
}
DEFAULT_PROVIDER = "ces"

REPORT_DESIGNATIONS = ("RN", "LPN", "RPN")


# =========================================================
# RB PORTAL (WordPress) DELIVERY CONFIG
# =========================================================
RBP_API_BASE     = "https://rbportal.nnas.ca/wp-json/rbp/v1"

# Point these at the SAME Windmill variables the scraper / record-upsert step
# already uses (WordPress user + Application Password).
RBP_WP_USER_VAR  = "u/knguyen/RBP_WP_USER"
RBP_APP_PASS_VAR = "u/knguyen/RBP_APP_PASS"

# Status written at report-delivery time. CONFIRMED against the portal: the
# download handler only serves when status == "Report Issued" (rb-portal.php
# ajax_download_report).
RBP_PUSH_STATUS  = "Report Issued"

# Recipient tags. The submission step writes the FULL tag set on every row.
# True  -> this step rewrites the leg's row with just that leg's tag. Visibility
#          is unchanged (the portal also requires the strict regulatory_body
#          match), and the term list is tighter.
# False -> all_recipients_raw is never sent, so whatever the submission step
#          wrote is preserved. Use this if anything downstream reads the full
#          recipient list off the record.
RBP_SEND_RECIPIENT_TAG = True

# Expiration business rules. Left None until confirmed: when set to a day
# count, order_expiration_date = order_date + N days and
# report_expiration_date = issued_date + N days are written with the record.
# Note the submission step already writes order_expiration_date as order_date
# + 2 years, so leaving RBP_ORDER_EXPIRY_DAYS as None is correct unless you
# intend this step to override that value.
# ported_date is never written by this step (set by the porting workflow).
RBP_ORDER_EXPIRY_DAYS: Optional[int] = None
RBP_REPORT_EXPIRY_DAYS: Optional[int] = None


# =========================================================
# REGULATORY BODY -> SFTP ABBREVIATION  (matches nnas_sftp_package.py)
# =========================================================
# Used ONLY to locate the report PDFs in SharePoint. Portal-facing fields come
# from resolve_rb() below, never from this map.
RB_ABBREV = {
    # --- RN ---
    "British Columbia College of Nurses and Midwives":                   "bccnm",
    "College Of Nurses Of Ontario":                                      "cno",
    "College of Registered Nurses Of Alberta":                           "crna",
    "College Of Registered Nurses Of Manitoba":                          "crnm",
    "College of Registered Nurses of Saskatchewan":                      "crns",
    "College of Nursing of New Brunswick":                               "cnnb",
    "College of Registered Nurses of Newfoundland and Labrador":         "nlcn",
    "College of Registered Nurses and Midwives of Prince Edward Island": "peicnm",
    "Nova Scotia College of Nursing - RN":                               "nscn",
    # --- LPN ---
    "Association Of New Brunswick Licensed Practical Nurses":            "anblpn",
    "College Of Licensed Practical Nurses Of Alberta":                   "clha",
    "College Of Licensed Practical Nurses Of Manitoba":                  "clpnm",
    "College of Licensed Practical Nurses of Saskatchewan":              "clpns",
    "College Of Licensed Practical Nurses Of Newfoundland And Labrador": "nlcn",
    "College of Licensed Practical Nurses of Prince Edward Island":      "peicnm",
    "Nova Scotia College of Nursing - LPN":                              "nscn",
    # --- Registered Psychiatric Nurses ---
    "College of Registered Psychiatric Nurses of Alberta":               "crpna",
    "College of Registered Psychiatric Nurses of Manitoba":              "crpnm",
    "College of Registered Psychiatric Nurses of Saskatchewan":          "crpns",
    # --- portal display-name aliases ---
    "NL College of Nurses":                                              "nlcn",
    "BC College of Nurses and Midwives":                                 "bccnm",
    "Nova Scotia College of Nursing":                                    "nscn",
    "PEI College of Nursing & Midwifery":                                "peicnm",
    "College of Licensed Practical Nurses and Health Care Aides of Alberta": "clha",
}
_KNOWN_ABBREVS = sorted(set(RB_ABBREV.values()), key=len, reverse=True)


# =========================================================
# SHARED RECORD VOCABULARY
# =========================================================
# Ported verbatim from nnas_rb_record_upsert.py. These two maps and the three
# functions under them are the contract between the submission stage and this
# delivery stage. If you edit one copy, edit both, or the finder stops binding
# and you get duplicate rows again.
#
# The portal surfaces a record to an RB user only when BOTH hold:
#   (a) the record's rb_recipient term intersects the user's tags, AND
#   (b) the record's regulatory_body matches one of the user's mapped names
#       (rb-portal.php tag_to_regulatory_bodies()).

_SINGLE = {  # normalized alias -> (TAG, canonical regulatory_body)
    "college of nurses of ontario":                              ("CNO",   "College Of Nurses Of Ontario"),
    "college of registered nurses of alberta":                  ("CRNA",  "College of Registered Nurses Of Alberta"),
    "college and association of registered nurses of alberta":  ("CRNA",  "College of Registered Nurses Of Alberta"),
    "college of registered nurses of manitoba":                 ("CRNM",  "College Of Registered Nurses Of Manitoba"),
    "college of registered nurses of saskatchewan":             ("CRNS",  "College of Registered Nurses of Saskatchewan"),
    "saskatchewan registered nurses association":               ("CRNS",  "College of Registered Nurses of Saskatchewan"),
    "college of nursing of new brunswick":                      ("CNNB",  "College of Nursing of New Brunswick"),
    "nurses association of new brunswick":                      ("CNNB",  "College of Nursing of New Brunswick"),
    "college of licensed practical nurses of manitoba":         ("CLPNM", "College Of Licensed Practical Nurses Of Manitoba"),
    "college of licensed practical nurses of saskatchewan":     ("CLPNS", "College of Licensed Practical Nurses of Saskatchewan"),
    "saskatchewan association of licensed practical nurses":    ("CLPNS", "College of Licensed Practical Nurses of Saskatchewan"),
    "college of licensed practical nurses and health care aides of alberta": ("CLHA", "College of Licensed Practical Nurses and Health Care Aides of Alberta"),
    "college of licensed practical nurses of alberta":          ("CLHA",  "College of Licensed Practical Nurses and Health Care Aides of Alberta"),
    "association of new brunswick licensed practical nurses":   ("ANBLPN","Association Of New Brunswick Licensed Practical Nurses"),
    "college of registered psychiatric nurses of alberta":      ("CRPNA", "College Of Registered Psychiatric Nurses Of Alberta"),
    "college of registered psychiatric nurses of manitoba":     ("CRPNM", "College Of Registered Psychiatric Nurses Of Manitoba"),
    "college of registered psychiatric nurses of saskatchewan": ("CRPNS", "College of Registered Psychiatric Nurses of Saskatchewan"),
    "registered psychiatric nurses association of saskatchewan":("CRPNS", "College of Registered Psychiatric Nurses of Saskatchewan"),
}

_BY_DISC = {  # normalized alias -> {"rn": (TAG, name), "lpn": (TAG, name)}
    "nl college of nurses": {
        "rn":  ("CRNNL",  "College Of Registered Nurses Of Newfoundland And Labrador"),
        "lpn": ("CLPNNL", "College Of Licensed Practical Nurses Of Newfoundland And Labrador"),
    },
    "newfoundland and labrador college of nurses": {
        "rn":  ("CRNNL",  "College Of Registered Nurses Of Newfoundland And Labrador"),
        "lpn": ("CLPNNL", "College Of Licensed Practical Nurses Of Newfoundland And Labrador"),
    },
    "college of registered nurses of newfoundland and labrador": {
        "rn":  ("CRNNL",  "College Of Registered Nurses Of Newfoundland And Labrador"),
        "lpn": ("CRNNL",  "College Of Registered Nurses Of Newfoundland And Labrador"),
    },
    "college of licensed practical nurses of newfoundland and labrador": {
        "rn":  ("CLPNNL", "College Of Licensed Practical Nurses Of Newfoundland And Labrador"),
        "lpn": ("CLPNNL", "College Of Licensed Practical Nurses Of Newfoundland And Labrador"),
    },
    "pei college of nursing & midwifery": {
        "rn":  ("CRNPEI",  "College Of Registered Nurses Of Prince Edward Island"),
        "lpn": ("CLPNPEI", "College of Licensed Practical Nurses of Prince Edward Island"),
    },
    "pei college of nursing and midwifery": {
        "rn":  ("CRNPEI",  "College Of Registered Nurses Of Prince Edward Island"),
        "lpn": ("CLPNPEI", "College of Licensed Practical Nurses of Prince Edward Island"),
    },
    "college of registered nurses and midwives of prince edward island": {
        "rn":  ("CRNPEI",  "College Of Registered Nurses Of Prince Edward Island"),
        "lpn": ("CRNPEI",  "College Of Registered Nurses Of Prince Edward Island"),
    },
    "college of licensed practical nurses of prince edward island": {
        "rn":  ("CLPNPEI", "College of Licensed Practical Nurses of Prince Edward Island"),
        "lpn": ("CLPNPEI", "College of Licensed Practical Nurses of Prince Edward Island"),
    },
    "nova scotia college of nursing": {
        "rn":  ("NSCN", "Nova Scotia College of Nursing-RN"),
        "lpn": ("NSCN", "Nova Scotia College of Nursing-LPN"),
    },
    "nova scotia college of nursing - rn": {
        "rn":  ("NSCN", "Nova Scotia College of Nursing-RN"),
        "lpn": ("NSCN", "Nova Scotia College of Nursing-RN"),
    },
    "nova scotia college of nursing - lpn": {
        "rn":  ("NSCN", "Nova Scotia College of Nursing-LPN"),
        "lpn": ("NSCN", "Nova Scotia College of Nursing-LPN"),
    },
    "bc college of nurses and midwives": {
        "rn":  ("BCCNM", "College Of Registered Nurses Of British Columbia"),
        "lpn": ("BCCNM", "College Of Licensed Practical Nurses Of British Columbia"),
    },
    "british columbia college of nurses and midwives": {
        "rn":  ("BCCNM", "College Of Registered Nurses Of British Columbia"),
        "lpn": ("BCCNM", "College Of Licensed Practical Nurses Of British Columbia"),
    },
}


def resolve_rb(name: str, discipline: str) -> Tuple[Optional[str], str]:
    """(alias, 'rn'|'lpn') -> (tag, canonical_name), or (None, name) if unmapped."""
    key = re.sub(r"\s+", " ", (name or "").strip().lower())
    if key in _SINGLE:
        return _SINGLE[key]
    if key in _BY_DISC:
        disc = discipline if discipline in ("rn", "lpn") else "rn"
        return _BY_DISC[key].get(disc) or next(iter(_BY_DISC[key].values()))
    return (None, name)


def discipline_label(discipline: str, regulatory_body: str) -> str:
    """Display label for the per-RB row (matches the portal's Discipline values).
    Psychiatric colleges arrive via the RN field but are a distinct discipline,
    so they are detected by name and override the rn/lpn split."""
    if "psychiatric" in (regulatory_body or "").lower():
        return "Registered Psychiatric Nurse"
    if discipline == "lpn":
        return "Licensed/Registered Practical Nurse"
    return "Registered Nurse"


def disc_key_for(designation: str, rb_name: str,
                 rn_list: List[str], lpn_list: List[str]) -> str:
    """Filename designation token (or RB-list membership) -> 'rn' | 'lpn'.
    RPN maps to 'rn' because the psychiatric colleges arrive in the RN field;
    discipline_label() then corrects the displayed label by college name."""
    d = (designation or "").strip().upper()
    if d == "LPN":
        return "lpn"
    if d in ("RN", "RPN"):
        return "rn"
    if rb_name in lpn_list and rb_name not in rn_list:
        return "lpn"
    return "rn"


def abbrev_for(rb_name: str) -> str:
    """SFTP folder abbreviation for the ORIGINAL SMApply alias. Not portal-facing."""
    if rb_name in RB_ABBREV:
        return RB_ABBREV[rb_name]
    low = {k.lower(): v for k, v in RB_ABBREV.items()}
    return low.get((rb_name or "").strip().lower(), "")


def regulator_of_report(filename: str) -> Optional[str]:
    """abbrev token embedded in a report filename, longest-first."""
    stem = os.path.splitext(filename or "")[0].lower()
    tokens = set(re.split(r"[-_.\s]+", stem))
    for ab in _KNOWN_ABBREVS:
        if ab in tokens:
            return ab
    return None


def designation_of_report(filename: str) -> str:
    """RN / LPN / RPN token in a report filename, or '' if none."""
    stem = os.path.splitext(filename or "")[0].upper()
    tokens = set(re.split(r"[-_.\s]+", stem))
    for d in REPORT_DESIGNATIONS:
        if d in tokens:
            return d
    return ""


# =========================================================
# SMALL HELPERS  (matches nnas_sftp_package.py semantics)
# =========================================================

def first_nonempty(*vals: str) -> str:
    for v in vals:
        if v and str(v).strip():
            return str(v).strip()
    return ""


_TAG_RE = re.compile(r"<[^>]+>")


def safe_str(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, dict):
        for k in ("label", "value", "name", "text"):
            if k in v and v[k] is not None:
                return _TAG_RE.sub("", str(v[k])).strip()
        return ""
    return _TAG_RE.sub("", str(v)).strip()


def _prune(fields: Dict[str, Any]) -> Dict[str, Any]:
    """Drop empty-string / None values before posting.

    Without this, a webhook that omits date_created sends order_date="" and
    wipes the value the submission stage wrote. Booleans and the match flags
    are preserved; only blank strings and None are removed.
    """
    out: Dict[str, Any] = {}
    for k, v in fields.items():
        if v is None:
            continue
        if isinstance(v, str) and not v.strip():
            continue
        out[k] = v
    return out


def normalize_portal_date(value: Any) -> str:
    """
    Normalize to the portal's canonical stored format Y-m-d (what the CSV
    importer writes; the filter SQL also tolerates d-m-Y, but new writes
    should match the imported rows). Handles: datetime, YYYY/MM/DD,
    YYYY-MM-DD, dd-mm-yyyy, and the webhook's '2026 Jun 7 03:43 AM' style.
    Unrecognized input returns '' rather than polluting a date column.
    """
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    v = (str(value or "")).strip()
    if not v:
        return ""
    m = re.match(r"^(\d{4})[/-](\d{1,2})[/-](\d{1,2})", v)              # YYYY/MM/DD, ISO
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    m = re.match(r"^(\d{1,2})-(\d{1,2})-(\d{4})$", v)                    # dd-mm-yyyy
    if m:
        return f"{m.group(3)}-{int(m.group(2)):02d}-{int(m.group(1)):02d}"
    m = re.match(r"^(\d{4})\s+([A-Za-z]{3})[a-z]*\s+(\d{1,2})\b", v)     # 2026 Jun 7 [03:43 AM]
    if m:
        months = {mn: i for i, mn in enumerate(
            ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
             "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], start=1)}
        mon = months.get(m.group(2)[:3].title())
        if mon:
            return f"{m.group(1)}-{mon:02d}-{int(m.group(3)):02d}"
    return ""


def _plus_days(ymd: str, days: Optional[int]) -> str:
    if not ymd or not days:
        return ""
    try:
        return (datetime.strptime(ymd, "%Y-%m-%d") + timedelta(days=days)).strftime("%Y-%m-%d")
    except ValueError:
        return ""


def base_template_for(provider: str) -> Tuple[str, str, Optional[str]]:
    p = (provider or "").strip().lower()
    if not p:
        return PROVIDER_BASE_TEMPLATE[DEFAULT_PROVIDER], DEFAULT_PROVIDER, None
    if p in PROVIDER_BASE_TEMPLATE:
        return PROVIDER_BASE_TEMPLATE[p], p, None
    return (PROVIDER_BASE_TEMPLATE[DEFAULT_PROVIDER], p,
            f"Unknown provider '{provider}'. Expected one of: "
            f"{', '.join(sorted(PROVIDER_BASE_TEMPLATE))} (or blank for {DEFAULT_PROVIDER}).")


def parse_additional_content(additional_content: Optional[str]) -> Dict[str, Any]:
    if not additional_content:
        return {}
    if isinstance(additional_content, dict):
        return additional_content
    s = str(additional_content).strip()
    try:
        raw = json.loads(s)
        if isinstance(raw, str):
            raw = json.loads(raw)
    except json.JSONDecodeError:
        if (s.startswith('"') and s.endswith('"')) or (s.startswith("'") and s.endswith("'")):
            try:
                raw = json.loads(s[1:-1])
                if isinstance(raw, str):
                    raw = json.loads(raw)
            except Exception:
                return {}
        else:
            return {}
    return raw if isinstance(raw, dict) else {}


# =========================================================
# SHAREPOINT / MICROSOFT GRAPH HELPERS
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
    return _url_quote(path.strip("/"), safe="/")


def _drive_relative_path(path: str) -> str:
    """
    Converts old SharePoint server-relative paths such as:
      /sites/NNAS-Portal-Sharepoint/Applications/CES/APPID
      /sites/NNAS-Portal-Sharepoint/applications/ces/APPID

    into paths relative to the Applications document-library drive:
      CES/APPID
    """
    if not path:
        return ""

    p = str(path).strip()

    if p.lower().startswith(SITE_PREFIX.lower()):
        p = p[len(SITE_PREFIX):]

    p = p.strip("/")
    parts = [part for part in p.split("/") if part]

    if parts and parts[0].lower() == DOCUMENT_LIBRARY_NAME.lower():
        parts = parts[1:]

    return "/".join(parts)


def _server_relative_url_from_drive_path(drive_path: str) -> str:
    return f"{SITE_PREFIX}/{DOCUMENT_LIBRARY_NAME}/{drive_path.strip('/')}"


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


def list_drive_children(drive_relative_path: str) -> List[Dict[str, Any]]:
    init_sharepoint_graph()

    clean_path = drive_relative_path.strip("/")

    if clean_path:
        endpoint = f"/drives/{_SP_DRIVE_ID}/root:/{encode_graph_path(clean_path)}:/children"
    else:
        endpoint = f"/drives/{_SP_DRIVE_ID}/root/children"

    items: List[Dict[str, Any]] = []
    data = graph_request("GET", endpoint, expected_statuses=(200,))
    items.extend(data.get("value", []))

    next_link = data.get("@odata.nextLink")
    while next_link:
        token = get_graph_token()
        response = requests.get(
            next_link,
            headers={"Authorization": f"Bearer {token}"},
            timeout=120,
        )
        if response.status_code != 200:
            raise RuntimeError(
                f"Microsoft Graph pagination failed.\n"
                f"Status: {response.status_code}\n"
                f"Response body:\n{response.text[:5000]}"
            )
        data = response.json()
        items.extend(data.get("value", []))
        next_link = data.get("@odata.nextLink")

    return items


def ctx():
    """
    Compatibility shim. The old Office365 client context is no longer needed;
    SharePoint calls are handled through Microsoft Graph helper functions.
    """
    init_sharepoint_graph()
    return None


def folder_exists(sp, folder_rel: str) -> bool:
    return get_drive_item_by_path(_drive_relative_path(folder_rel)) is not None


def list_folder_files(sp, folder_rel: str, recursive: bool = True) -> List[Dict[str, Any]]:
    results: List[Dict[str, Any]] = []
    start_drive_path = _drive_relative_path(folder_rel)

    def walk(drive_path: str):
        try:
            children = list_drive_children(drive_path)
        except Exception:
            return

        for item in children:
            name = item.get("name", "")
            if not name:
                continue

            child_drive_path = f"{drive_path.rstrip('/')}/{name}".strip("/")

            if item.get("file") is not None:
                results.append({
                    "name": name,
                    "url": _server_relative_url_from_drive_path(child_drive_path),
                    "drive_path": child_drive_path,
                    "id": item.get("id", ""),
                    "webUrl": item.get("webUrl", ""),
                })
                continue

            if recursive and item.get("folder") is not None and name not in ("Forms",):
                walk(child_drive_path)

    walk(start_drive_path)
    return results


def download_file(sp, server_rel: str, local_path: str) -> None:
    drive_path = _drive_relative_path(server_rel)
    item = get_drive_item_by_path(drive_path)

    if not item:
        raise RuntimeError(f"SharePoint file not found: {server_rel}")

    os.makedirs(os.path.dirname(local_path), exist_ok=True)

    response = requests.get(
        f"{GRAPH_BASE}/drives/{_SP_DRIVE_ID}/items/{item['id']}/content",
        headers={"Authorization": f"Bearer {get_graph_token()}"},
        timeout=300,
        allow_redirects=True,
    )

    if response.status_code != 200:
        raise RuntimeError(
            f"Failed downloading SharePoint file: {server_rel}\n"
            f"Status: {response.status_code}\n"
            f"Response body:\n{response.text[:2000]}"
        )

    with open(local_path, "wb") as fh:
        fh.write(response.content)


# =========================================================
# RB PORTAL CLIENT
# =========================================================

class RBPortalClient:
    """
    Thin client for the RBP Upload API (Application Password + IP allowlist).
    Retries on connection blips / 5xx / 429 with exponential backoff -- same
    failure mode as the SMApply PATCH ConnectionError (Docker DNS blip on the
    self-hosted Windmill). Multi-MB PDF uploads get their own longer timeout.
    """

    _RETRYABLE_STATUS = {429, 500, 502, 503, 504}
    _TAG_RE = re.compile(r"<[^>]+>")

    def __init__(self, base_url: str, wp_user: str, app_pass: str,
                 *, timeout: int = 30, upload_timeout: int = 120,
                 retries: int = 4, backoff_base: float = 1.5):
        if not wp_user or not app_pass:
            raise RuntimeError(
                "RB Portal credentials not resolved from Windmill variables "
                f"({RBP_WP_USER_VAR} / {RBP_APP_PASS_VAR})."
            )
        self.base = base_url.rstrip("/")
        self.timeout = timeout
        self.upload_timeout = upload_timeout
        self.retries = retries
        self.backoff_base = backoff_base
        self.session = requests.Session()
        self.session.auth = (wp_user, app_pass)

    def _err_text(self, resp: requests.Response) -> str:
        # WordPress error bodies can be full HTML pages; strip tags so
        # Windmill job logs stay readable.
        return self._TAG_RE.sub("", resp.text or "").strip()[:300]

    def _request(self, method: str, path: str, **kw) -> requests.Response:
        url = f"{self.base}/{path.lstrip('/')}"
        kw.setdefault("timeout", self.timeout)
        last_err = None
        for attempt in range(self.retries + 1):
            try:
                resp = self.session.request(method, url, **kw)
                if resp.status_code in self._RETRYABLE_STATUS:
                    last_err = RuntimeError(
                        f"HTTP {resp.status_code}: {self._err_text(resp)}")
                    raise last_err
                return resp
            except (requests.ConnectionError, requests.Timeout, RuntimeError) as e:
                last_err = e
                if attempt >= self.retries:
                    break
                time.sleep(self.backoff_base ** attempt)
        raise last_err

    def upsert_record(self, application_id: str, fields: dict) -> dict:
        path = f"/record/{_url_quote(str(application_id), safe='')}"
        resp = self._request("POST", path, json=_prune(fields))
        resp.raise_for_status()
        return resp.json()

    def upload_report(self, post_id: int, pdf_path: str, *,
                      regulatory_body: str = "", discipline: str = "",
                      overwrite: bool = True) -> dict:
        path = f"/report/{int(post_id)}"
        data = {"overwrite": "true" if overwrite else "false", "require_record": "true",
                # Explicit binding: the plugin (v1.6.2+) gives this field top
                # precedence, so a numeric legacy application_id can never be
                # confused with this post_id (or vice versa).
                "post_id": str(int(post_id))}
        if regulatory_body:
            data["regulatory_body"] = regulatory_body
        if discipline:
            data["discipline"] = discipline
        with open(pdf_path, "rb") as fh:
            files = {"file": (os.path.basename(pdf_path), fh, "application/pdf")}
            resp = self._request("POST", path, data=data, files=files,
                                 timeout=self.upload_timeout)
        resp.raise_for_status()
        return resp.json()


def _match_flags() -> dict:
    """Record addressing.

    application_id + regulatory_body only, which is exactly how
    nnas_rb_record_upsert.py addressed the row at submission. Discipline is
    deliberately NOT part of the key: the canonical names already differ per
    designation for every dual-designation college (BCCNM, NSCN, NL, PEI), so
    adding discipline buys nothing and breaks the bind whenever the two steps
    disagree on the label.
    """
    return {"_match_by_regulatory_body": True}


# =========================================================
# ENTRYPOINT
# =========================================================

def main(reference_id: Optional[str] = None,
         additional_content: Optional[str] = None,
         user_id: Any = None,
         dry_run: bool = False,
         **kwargs: Any):
    """
    Same input shape as nnas_sftp_package.py:
        reference_id (or title) -> ApplicationID + SharePoint folder name
        user_id                 -> NNAS <ID> (stored as nnas_id on the record)
        additional_content      -> the en-masse JSON blob (regulators, provider)

    dry_run: resolve and report what WOULD be uploaded, but make no portal calls.
    """
    title        = kwargs.get("title")
    nnas_id      = kwargs.get("nnas_id")
    date_created = kwargs.get("date_created")   # webhook top-level -> Order Date

    payload = parse_additional_content(additional_content)

    # ---- applicant-level portal fields (shared by every leg) ----
    pi = payload.get("personal_info") or {}
    ien_name = " ".join(p for p in (safe_str(pi.get("sa_first_name")),
                                    safe_str(pi.get("sa_middle_name")),
                                    safe_str(pi.get("sa_last_name"))) if p)
    dob        = normalize_portal_date(safe_str(pi.get("sa_dob")))
    country    = safe_str(pi.get("sa_initial_education_country"))
    order_date = normalize_portal_date(date_created)
    order_expiration_date = _plus_days(order_date, RBP_ORDER_EXPIRY_DAYS)

    app_id = first_nonempty(reference_id, title, safe_str(payload.get("reference_id")))
    if not app_id:
        return {"status": "error",
                "message": "No reference_id/title to locate the application."}

    doc_id = first_nonempty(
        (str(user_id) if user_id is not None else ""),
        (str(nnas_id) if nnas_id is not None else ""),
        safe_str(payload.get("nnas_id")),
    )

    provider = safe_str((payload.get("application_info") or {}).get("provider"))
    base_template, provider_norm, provider_err = base_template_for(provider)
    if provider_err:
        return {"status": "error", "message": provider_err,
                "debug": {"provider": provider}}

    sp = ctx()
    app_base_rel = f"{SITE_PREFIX}/{base_template.format(app_id=app_id)}"
    reports_rel  = f"{app_base_rel}/{CREDENTIAL_REPORTS_SUBFOLDER}"

    if not folder_exists(sp, app_base_rel):
        return {"status": "error",
                "message": (f"Application folder not found for provider "
                            f"'{provider_norm.upper()}': {app_base_rel}."),
                "debug": {"provider": provider_norm, "base_path": app_base_rel}}

    # ---- regulator list (union of RN + LPN, de-duplicated, order preserved) ----
    ai = payload.get("application_info") or {}
    rn_list  = [x.strip() for x in safe_str(ai.get("regulatory_body_rn")).split(",") if x.strip()]
    lpn_list = [x.strip() for x in safe_str(ai.get("regulatory_body_lpn")).split(",") if x.strip()]
    rb_names: List[str] = []
    seen = set()
    for rb in rn_list + lpn_list:
        if rb.lower() not in seen:
            seen.add(rb.lower())
            rb_names.append(rb)

    # ---- map credential-report PDFs to (abbrev -> designation -> file) ----
    # Only *_latest*.pdf is considered; docx sources and "Applicant Copy"
    # redactions (named *_applicant.pdf) are excluded.
    report_files = list_folder_files(sp, reports_rel, recursive=True)
    reports_by_abbrev: Dict[str, Dict[str, Dict[str, Any]]] = {}
    skipped: List[str] = []
    for rf in report_files:
        nm = rf["name"]
        low = nm.lower()
        if not low.endswith(".pdf"):
            continue
        if "latest" not in low:
            skipped.append(nm + " (not _latest)")
            continue
        ab = regulator_of_report(nm)
        if not ab:
            skipped.append(nm + " (no regulator abbrev matched)")
            continue
        desig = designation_of_report(nm)
        reports_by_abbrev.setdefault(ab, {})[desig] = rf

    debug: Dict[str, Any] = {
        "application_id": app_id, "id": doc_id,
        "provider": provider_norm.upper(),
        "regulators": rb_names,
        "reports_by_abbrev": {ab: [e["name"] for e in d.values()]
                              for ab, d in reports_by_abbrev.items()},
        "match_mode": "application_id+regulatory_body (matches submission stage)",
        "send_recipient_tag": RBP_SEND_RECIPIENT_TAG,
    }

    # ---- open the RB Portal client (skipped on dry_run) ----
    rbp = None
    rbp_connect_error: Optional[str] = None
    if dry_run:
        rbp_connect_error = "dry run: RB portal push skipped"
    else:
        try:
            rbp = RBPortalClient(RBP_API_BASE,
                                 wmill.get_variable(RBP_WP_USER_VAR),
                                 wmill.get_variable(RBP_APP_PASS_VAR))
        except Exception as e:
            rbp_connect_error = f"RB portal client init failed: {e}"

    workdir = os.path.join("tmp", f"rbp_{app_id}_{datetime.utcnow():%Y%m%d_%H%M%S}")
    os.makedirs(workdir, exist_ok=True)
    today = normalize_portal_date(datetime.utcnow())

    results: List[Dict[str, Any]] = []
    unmapped: List[str] = []
    seen_rows: set = set()        # canonical regulatory_body already handled
    uploaded = 0

    for rb_name in rb_names:
        abbrev = abbrev_for(rb_name)
        if not abbrev:
            results.append({"regulator": rb_name, "skipped": "no abbreviation mapping"})
            continue

        by_des = reports_by_abbrev.get(abbrev, {})
        if not by_des:
            results.append({"regulator": rb_name, "abbrev": abbrev,
                            "skipped": "no _latest report pdf in SharePoint"})
            continue

        # Untagged-leg guard: with mixed filename conventions a folder can hold
        # both an old untagged "..._latest.pdf" and a new "..._RN_..._latest.pdf".
        # The untagged leg's discipline is inferred from the RB list, so if a
        # tagged leg for that same discipline exists, the two would upsert the
        # same row and the second PDF would overwrite the first. Drop the
        # untagged one; the designation-tagged file is the newer convention.
        inferred = ("RN" if rb_name in rn_list else
                    "LPN" if rb_name in lpn_list else "")
        legs = dict(by_des)
        if "" in legs and inferred and inferred in legs:
            skipped.append(f"{legs['']['name']} (untagged; tagged {inferred} "
                           f"leg already present for {abbrev})")
            legs.pop("")

        for desig, rf in legs.items():
            # Resolve to the SAME identity the submission stage used.
            disc_key = disc_key_for(desig or inferred, rb_name, rn_list, lpn_list)
            recipient_tag, canonical_rb = resolve_rb(rb_name, disc_key)
            disc_label = discipline_label(disc_key, canonical_rb)

            leg: Dict[str, Any] = {
                "regulator": rb_name,
                "abbrev": abbrev,
                "regulatory_body": canonical_rb,
                "discipline": disc_label,
                "source": rf["name"],
            }

            if recipient_tag is None:
                # Unmapped alias: the record is still written (admin-visible)
                # but it will not bind to a submission-stage row, because that
                # stage wrote the same unmapped name. Surface it so the alias
                # can be added to BOTH copies of the map.
                unmapped.append(rb_name)
                leg["warning_tag"] = (
                    f"'{rb_name}' is not in _SINGLE / _BY_DISC. No recipient tag "
                    f"was sent and regulatory_body falls back to the raw alias. "
                    f"Add the alias to both this script and "
                    f"nnas_rb_record_upsert.py.")

            row_key = (canonical_rb or "").strip().lower()
            if row_key and row_key in seen_rows:
                leg["skipped"] = ("another alias in this application already "
                                  "resolved to this canonical regulatory_body")
                results.append(leg)
                continue
            if row_key:
                seen_rows.add(row_key)

            if rbp is None:                       # dry run / no client
                leg["would_upload"] = True
                leg["recipient_tag"] = recipient_tag or ""
                results.append(leg)
                continue

            try:
                pdf_local = os.path.join(workdir, f"{abbrev}_{disc_key}_{rf['name']}")
                download_file(sp, rf["url"], pdf_local)

                # Three-phase delivery so status only flips once the PDF is
                # actually stored. The portal renders the download link purely
                # off status == "Report Issued"; setting it before the upload
                # would, on an upload failure, leave a live link that 404s for
                # the college. Phase 1 binds the existing submission-stage row
                # (or creates it if submission never ran), phase 2 stores the
                # PDF against its post_id, phase 3 flips status + issued_date.
                #
                # Everything here is pruned of empty strings by the client, so
                # a missing dob / country / order_date leaves the submission
                # stage's value in place rather than blanking the column.
                bind_fields: Dict[str, Any] = {
                    "nnas_id": doc_id,
                    "regulatory_body": canonical_rb,
                    "discipline": disc_label,
                    "ien_name": ien_name,
                    "dob": dob,
                    "country": country,
                    "order_date": order_date,
                    "order_expiration_date": order_expiration_date,
                    **_match_flags(),
                }
                if RBP_SEND_RECIPIENT_TAG and recipient_tag:
                    bind_fields["all_recipients_raw"] = recipient_tag

                up = rbp.upsert_record(app_id, bind_fields)
                post_id = up.get("post_id")

                rep = rbp.upload_report(post_id, pdf_local,
                                        regulatory_body=canonical_rb,
                                        discipline=disc_label)

                fin_fields: Dict[str, Any] = {
                    "regulatory_body": canonical_rb,
                    "discipline": disc_label,
                    "status": RBP_PUSH_STATUS,
                    "issued_date": today,
                    **_match_flags(),
                }
                report_expiration = _plus_days(today, RBP_REPORT_EXPIRY_DAYS)
                if report_expiration:
                    fin_fields["report_expiration_date"] = report_expiration
                fin = rbp.upsert_record(app_id, fin_fields)

                leg.update({
                    "post_id": post_id,
                    "matched_by": up.get("matched_by"),
                    "created": up.get("created"),
                    "saved_as": rep.get("saved_as"),
                    "recipient_tag": recipient_tag or "",
                    "status_set": fin.get("ok", False) and fin.get("post_id") == post_id,
                })
                if up.get("created"):
                    # Expected only when the submission stage never ran for this
                    # application. If it fires on a normal flow, the two scripts
                    # have drifted apart again: compare canonical names.
                    leg["warning_created"] = (
                        "no existing record matched application_id + "
                        f"'{canonical_rb}'; a NEW row was created. Verify the "
                        "submission stage wrote the same canonical name.")
                if fin.get("post_id") != post_id:
                    # Should be impossible with deterministic finders; surface it.
                    leg["warning"] = (f"status upsert bound post_id "
                                      f"{fin.get('post_id')} != {post_id}")
                uploaded += 1
            except Exception as e:
                leg["error"] = str(e)
                skipped.append(f"RB portal delivery failed for {abbrev}/{disc_key}: {e}")
            results.append(leg)

    # tidy up the temp dir
    try:
        shutil.rmtree(workdir, ignore_errors=True)
    except Exception:
        pass

    debug["skipped"] = skipped
    debug["unmapped_regulators"] = sorted(set(unmapped))
    debug["rbp_connect_error"] = rbp_connect_error
    return {
        "status": "ok",
        "application_id": app_id,
        "id": doc_id,
        "dry_run": dry_run,
        "rbp_delivered": (not dry_run) and rbp_connect_error is None,
        "uploaded_count": uploaded,
        "leg_count": len(results),
        "results": results,
        "debug": debug,
    }