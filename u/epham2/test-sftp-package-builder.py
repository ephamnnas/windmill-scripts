# py: 3.12
# requirements:
# wmill
# requests
# pdfplumber
# paramiko
# pymupdf

"""
NNAS SFTP Package Builder  (self-contained Windmill script)
===========================================================

Independent flow from the credential-report generator. For one application
that already has its documents in SharePoint, this script:

  1. Parses the en-masse `additional_content` webhook blob (application_info,
     personal_info, registration, ...).
  2. Downloads and parses the captured application-summary PDF from the base
     folder (fixed label/value text layout -- no OCR) for the applicant-level
     fields not present in additional_content (Gender, languages, conditions
     Y/N block, ID document types, etc.).
  3. Collects the files to package from SharePoint:
       - base-folder ID docs:  <ApplicationID>_ID1 / _ID2 ...
       - the "International Credentials" subfolder (recursive)
       - translation documents in "International Credentials" (any filename
         containing "translation") -> always images/other/, doctype TRANSLATION
       - the credential Report -- ONLY THE MOST FINAL VERSION (*_latest.pdf).
       - the applicant application-summary form from the application root
         (<ReferenceID>_<YYYYMMDD>_<HHMMSS>), most recent timestamp only.
  4. Renames each file to the NNAS convention and lays them out as:
       images/identity/   images/registration/   images/report/   images/other/
       {ID}-{ApplicationID}-{category}-{DOCTYPE}{N}.{ext}
  5. Pulls each document's metadata block (DocumentNumber, IssuedFrom,
     ReceivedDate, DateIssued, ExpirationDate, Size, LastModifiedDate) from the
     SharePoint file metadata.
  6. Emits the interchange CSV and XML  ({ID}_{ApplicationID}.csv / .xml).
  7. Checks the CSV output (verify_csv_output) and the XML output
     (verify_xml_output) before delivery.
  8. Zips everything to {ApplicationID}.zip, creates a NEW export folder under
     the application folder, and uploads the zip there. The CSV and XML are
     also uploaded as loose files beside the zip (EXPORT_LOOSE_INTERCHANGE).

Entrypoint signature is identical to the report generator so the same webhook
payload maps cleanly.

CHANGES IN THIS REVISION
------------------------
* XML output check (verify_xml_output). After each regulator's XML is built,
  it is parsed and checked per <Application>: well-formed, one Application per
  designation, DateOfBirth in dd-mm-yyyy, at least one identity document, and
  no form label text leaking into values (e.g. "First language: ..." in
  DateOfBirth). Key fields that are blank are listed by tag name, and fields
  with a value but an empty iso-code are listed for information. Results are
  returned as packages[].xml_check. Like the CSV check, only tag names and
  counts are reported, never applicant values.
* When XML_CHECK_BLOCKS_DELIVERY is True (default), a regulator whose XML
  fails the check is held back from the SFTP, the same way a failed CSV check
  is. The list of held-back regulators is debug.check_blocked (previously
  debug.csv_check_blocked).
* Loose interchange files. When EXPORT_LOOSE_INTERCHANGE is True (default),
  the CSV and XML are uploaded to the SharePoint export folder next to the
  zip, and their links are returned as packages[].csv_file / xml_file. The
  file contents are NOT returned in the run output, since they carry
  applicant PII; open them in SharePoint instead.

Previous revision:

* CSV output check (verify_csv_output). After each regulator's CSV is built,
  it is re-read and checked row by row: one row per designation,
  ApplicationID and NursingGroup match the record, and DateOfBirth is a real
  dd-mm-yyyy date, not in the future, gives an age between MIN_APPLICANT_AGE
  and MAX_APPLICANT_AGE, and matches both the applicant model and the NNAS
  index row. Problems are added to `warnings` (prefixed with the regulator)
  and the full result is returned as packages[].csv_check. Like check_dob,
  no DOB or other applicant values are written to the run output.
* When CSV_CHECK_BLOCKS_DELIVERY is True (default), a regulator whose CSV
  fails the check is NOT pushed to the SFTP: no zip upload and no index rows.
  The SharePoint export still happens so the package can be inspected. Set
  it to False to warn only.

Features ported from the TM script:

* Applicant application-summary form is now PACKAGED. The newest
  '<ReferenceID>_<YYYYMMDD>_<HHMMSS>' file in the application root is copied
  into images/other/ of every regulator package as
  '<ReferenceID>-Applicant Form - <YYYYMMDD_HHMMSS>.<ext>'. Previously the
  form was parsed but never included in the zip.
* The application-summary PDF that gets PARSED is the same file that gets
  PACKAGED (pick_latest_applicant_form). The old loose filter (first PDF in
  the listing that is not an ID doc or report) is kept only as a fallback
  when no stamped summary exists, and that fallback now raises a warning.
  debug.form_pdf_source records which path was taken.
* Date of birth cross-check (check_dob). The PDF value and additional_content
  personal_info.sa_dob are both normalized to dd-mm-yyyy and compared. A
  warning is raised when they disagree, when both are blank, or when the
  written value cannot be normalized. Precedence is unchanged (PDF wins).
  Warnings and debug never echo the DOB values, to keep applicant PII out of
  run output.
* Report selection requires the '_latest.pdf' suffix (REPORT_LATEST_SUFFIX)
  instead of "latest" anywhere in the filename, so '_latest.docx' and names
  such as '_latest_old.pdf' are no longer packaged. Skip messages name the
  suffix being tested.
* A failure to parse the application summary is surfaced in `warnings`
  instead of only in debug.

Retained from the previous CES revision:

* Translation documents in "International Credentials" (filename contains
  TRANSLATION_KEYWORD) are routed to images/other/ with doctype TRANSLATION,
  with ReceivedDate falling back to the SharePoint upload date.
* `additional_content` is repaired before parsing; parse failure and repair
  are surfaced in `warnings`.
* The PDF fallback for the regulatory-body prompt reads ALL selected values,
  stops at the next form prompt, and splits on commas.
* Reports for regulators not in the resolved regulator list are reported in
  `warnings`.
* Path handling: provider base templates carry no leading slash and
  `app_base_rel` is built through `_display_server_relative_path`; the
  "Additional documents" prefix comparison is explicit and case-insensitive.
"""

import os
import re
import csv
import json
import zipfile
import shutil
import xml.etree.ElementTree as ET
from io import StringIO
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple
from xml.sax.saxutils import escape as _xml_escape
from urllib.parse import quote

import wmill
import requests
import pdfplumber
import paramiko
import fitz  # PyMuPDF, for the applicant-copy redaction


# =========================================================
# SHAREPOINT CONFIG
# =========================================================

SITE_URL      = "https://nnasca.sharepoint.com/sites/NNAS-Portal-Sharepoint"
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

INTL_CREDENTIALS_SUBFOLDER = "International Credentials"
ADDITIONAL_DOCS_SUBFOLDER  = "Additional documents"
EXPORT_FOLDER_PREFIX       = "SFTP_Export"

# When True, the interchange CSV and XML are uploaded to the SharePoint export
# folder as loose files next to each regulator's zip, so testers can open them
# without unzipping. Only their links are returned in the run output.
EXPORT_LOOSE_INTERCHANGE = True

# Translation documents inside "International Credentials". Matched by
# filename (case-insensitive substring), the same rule the credential report
# generator uses for Part C4. Matches are always packaged under images/other/
# with this doctype, regardless of any SharePoint tags on the file, e.g.
#   {ID}-{ApplicationID}-other-TRANSLATION1.pdf
TRANSLATION_KEYWORD = "translation"
TRANSLATION_DOCTYPE = "TRANSLATION"


def is_translation_file(file_name: str) -> bool:
    return TRANSLATION_KEYWORD.lower() in (file_name or "").lower()


# Provider (application_info.provider) selects the SharePoint base path,
# relative to the Applications document library root.
# TM  -> tm/<app_id>
# CES -> ces/<app_id>   (also the default when provider is absent)
# NOTE: no leading slash. app_base_rel is built as f"{SITE_PREFIX}/{template}"
# and must have single slashes so it can be sliced against file URLs.
PROVIDER_BASE_TEMPLATE = {
    "tm":  "tm/{app_id}",
    "ces": "ces/{app_id}",
}
DEFAULT_PROVIDER = "ces"


def base_template_for(provider: str) -> Tuple[str, str, Optional[str]]:
    """
    Resolve (template, normalized_provider, error).
    - blank/absent  -> DEFAULT_PROVIDER, no error
    - tm / ces      -> matching template, no error
    - anything else -> error string (caller should fail; do NOT silently default)
    """
    p = (provider or "").strip().lower()
    if not p:
        return PROVIDER_BASE_TEMPLATE[DEFAULT_PROVIDER], DEFAULT_PROVIDER, None
    if p in PROVIDER_BASE_TEMPLATE:
        return PROVIDER_BASE_TEMPLATE[p], p, None
    return (PROVIDER_BASE_TEMPLATE[DEFAULT_PROVIDER], p,
            f"Unknown provider '{provider}'. Expected one of: "
            f"{', '.join(sorted(PROVIDER_BASE_TEMPLATE))} (or blank for {DEFAULT_PROVIDER}).")

# Credential reports are stored by regulator inside this subfolder of the
# application folder, e.g. "credential reports/BCCNM/". Only the most-final
# version is packaged -- identified by the "_latest.pdf" filename suffix.
# A regulator that assesses an applicant under more than one designation (e.g.
# BCCNM does both RN and LPN) will have one "_latest.pdf" PER designation, each
# carrying an _RN_ / _LPN_ / _RPN_ token in the filename.
CREDENTIAL_REPORTS_SUBFOLDER = "credential reports"
REPORT_LATEST_SUFFIX         = "_latest.pdf"

# Recognized designation tokens inside a report filename.
REPORT_DESIGNATIONS = ("RN", "LPN", "RPN")

# Applicant copy: a redacted report (Assessment Notes block removed) is written
# here, in a new folder under each regulator's report folder.
APPLICANT_COPY_SUBFOLDER = "Applicant Copy"
APPLICANT_COPY_SUFFIX    = "_applicant"   # replaces the "_latest" suffix

# =========================================================
# SFTP (regulator delivery) CONFIG
# =========================================================
SFTP_HOST = "rb-ftp-connect.nnas.ca"
SFTP_PORT = 2022
SFTP_USERNAME_VAR = "u/epham2/NNAS_SFTP_USERNAME"
SFTP_PASSWORD_VAR = "u/epham2/NNAS_SFTP_PASSWORD"

# Per-regulator destination: /REGULATORY/<ABBREV-UPPER>/ on the SFTP root.
SFTP_REGULATORY_ROOT = "/REGULATORY"

# The running NNAS index appended per regulator (headerless, 9 columns):
#   LASTNAME,FIRSTNAME,id,applicationID,DOB,RN|LPN,x,N,x
# Named per regulator + day: <RB>_NNAS_index_<DDMMYYYY>.csv  (e.g.
# CNO_NNAS_index_01062026.csv). A new file is started each calendar day; the
# same day's file is read and appended to.
SFTP_INDEX_TEMPLATE = "{rb}_NNAS_index_{ddmmyyyy}.csv"


def index_filename(abbrev: str, when: Optional[datetime] = None) -> str:
    when = when or datetime.utcnow()
    return SFTP_INDEX_TEMPLATE.format(rb=abbrev.upper(),
                                      ddmmyyyy=when.strftime("%d%m%Y"))

# Folder layout inside the zip (lowercase, matching the reference Path values).
IMAGES_ROOT = "images"
CATEGORY_FOLDERS = {
    "identity":     f"{IMAGES_ROOT}/identity",
    "registration": f"{IMAGES_ROOT}/registration",
    "report":       f"{IMAGES_ROOT}/report",
    "other":        f"{IMAGES_ROOT}/other",
}

# FileType tag (lower) -> category. Anything else -> "other".
FILETYPE_TO_CATEGORY = {
    "identity":     "identity",
    "id":           "identity",
    "registration": "registration",
    "report":       "report",
    "education":    "other",
    "other":        "other",
}

# Base-folder ID document convention: <ApplicationID>_ID<n>.<ext>
ID_DOC_RE = re.compile(r"_ID(\d+)\b", re.IGNORECASE)


# =========================================================
# APPLICANT APPLICATION-SUMMARY FORM  (application root)
# =========================================================
# The captured application summary lives in the application root and is named
#   <ReferenceID>_<YYYYMMDD>_<HHMMSS>.<ext>   e.g. NNAS-T-0000000123_20260618_040729.pdf
# There may be several; only the most recent timestamp is packaged. It is
# copied into images/other/ as
#   <ReferenceID>-Applicant Form - <YYYYMMDD_HHMMSS>.<ext>
# The strict <ReferenceID>_ prefix keeps generated credential reports
# (<ref>_CredentialForm_YYYYMMDD_HHMMSS.pdf) out of this match.
# The same newest file is also the one parsed for applicant-level fields.
APPLICANT_FORM_LABEL = "Applicant Form"


def applicant_form_stamp(file_name: str, app_id: str) -> Optional[str]:
    """'YYYYMMDD_HHMMSS' if file_name is exactly '<app_id>_<date>_<time>.<ext>'."""
    stem = os.path.splitext(file_name or "")[0]
    m = re.match(rf"^{re.escape(app_id)}_(\d{{8}})_(\d{{6}})$", stem, re.IGNORECASE)
    return f"{m.group(1)}_{m.group(2)}" if m else None


def pick_latest_applicant_form(base_files: List[Dict[str, Any]],
                               app_id: str) -> Tuple[Optional[Dict[str, Any]], Optional[str], List[str]]:
    """
    Most recent '<app_id>_<YYYYMMDD>_<HHMMSS>' file in the application root.
    Returns (entry, stamp, all_candidate_names). Stamps sort lexicographically
    in chronological order, so a plain reverse sort picks the newest.
    """
    cands: List[Tuple[str, Dict[str, Any]]] = []
    for f in base_files:
        stamp = applicant_form_stamp(f.get("name", ""), app_id)
        if stamp:
            cands.append((stamp, f))
    if not cands:
        return None, None, []
    cands.sort(key=lambda t: t[0], reverse=True)
    return cands[0][1], cands[0][0], [f["name"] for _, f in cands]


def applicant_form_name(app_id: str, stamp: str, ext: str) -> str:
    """Packaged name: '<ReferenceID>-Applicant Form - <YYYYMMDD_HHMMSS><ext>'."""
    return f"{app_id}-{APPLICANT_FORM_LABEL} - {stamp}{ext}"


# =========================================================
# PER-DOCUMENT METADATA MAPPING  (SharePoint column DISPLAY names)
# =========================================================
# Single source of truth. None -> no column in the library -> emitted blank.

COLUMN_MAP: Dict[str, Optional[str]] = {
    "DocumentNumber": "TrackingReferenceNumber",
    "IssuedFrom":     "Education Countries",
    "DateIssued":     "DateIssued",       # absent in the Item-properties panel -> blank
    "ReceivedDate":   "PackageReceivedDate",
    "ExpirationDate": "ExpirationDate",   # absent in the Item-properties panel -> blank
}
DOCTYPE_FIELDS = ["Title", "Image Tags", "FileType"]

# NOTE: only these countries get an iso-code. Any other country is written
# with an empty iso-code attribute; verify_xml_output lists those fields as
# informational (empty_iso_code_fields).
ISO_NUMERIC = {
    "UNITED STATES": "840", "NIGERIA": "566", "CANADA": "124",
    "PHILIPPINES": "608", "INDIA": "356", "UNITED KINGDOM": "826",
}


# =========================================================
# NURSING GROUP / REGULATORY BODY -> PROVINCE
# =========================================================

RB_PROVINCE_MAP = {
    "British Columbia College of Nurses and Midwives":                   "British Columbia",
    "College Of Nurses Of Ontario":                                      "Ontario",
    "College of Registered Nurses Of Alberta":                           "Alberta",
    "College Of Registered Nurses Of Manitoba":                          "Manitoba",
    "College of Registered Nurses and Midwives of Prince Edward Island": "Prince Edward Island",
    "College of Registered Nurses of Newfoundland and Labrador":         "Newfoundland and Labrador",
    "College of Registered Nurses of Saskatchewan":                      "Saskatchewan",
    "College of Nursing of New Brunswick":                               "New Brunswick",
    "Nova Scotia College of Nursing - RN":                               "Nova Scotia",
    "Association Of New Brunswick Licensed Practical Nurses":            "New Brunswick",
    "College Of Licensed Practical Nurses Of Alberta":                   "Alberta",
    "College Of Licensed Practical Nurses Of Manitoba":                  "Manitoba",
    "College Of Licensed Practical Nurses Of Newfoundland And Labrador": "Newfoundland and Labrador",
    "College of Licensed Practical Nurses of Prince Edward Island":      "Prince Edward Island",
    "College of Licensed Practical Nurses of Saskatchewan":              "Saskatchewan",
    "Nova Scotia College of Nursing - LPN":                              "Nova Scotia",
    # --- portal display-name aliases (names as they arrive in additional_content) ---
    "NL College of Nurses":                                              "Newfoundland and Labrador",
    "BC College of Nurses and Midwives":                                 "British Columbia",
    "Nova Scotia College of Nursing":                                    "Nova Scotia",
    "PEI College of Nursing & Midwifery":                                "Prince Edward Island",
    "College of Licensed Practical Nurses and Health Care Aides of Alberta": "Alberta",
}


# Regulatory body -> SFTP abbreviation (from the sftp-* folder screenshot).
# This is the RBNAME used to match each report file in the "credential
# reports" folder to its regulator and to pick the /REGULATORY/<RB> folder.
# Confident mappings cover all RN bodies; a few LPN/RPN rows are best-effort
# (flagged) -- adjust here if an abbreviation differs.
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
    "College Of Licensed Practical Nurses Of Alberta":                   "clha",    # best-effort
    "College Of Licensed Practical Nurses Of Manitoba":                  "clpnm",
    "College of Licensed Practical Nurses of Saskatchewan":             "clpns",
    "College Of Licensed Practical Nurses Of Newfoundland And Labrador": "nlcn",    # NL merged college
    "College of Licensed Practical Nurses of Prince Edward Island":      "peicnm",  # best-effort
    "Nova Scotia College of Nursing - LPN":                              "nscn",
    # --- Registered Psychiatric Nurses (not in province map) ---
    "College of Registered Psychiatric Nurses of Alberta":               "crpna",
    "College of Registered Psychiatric Nurses of Manitoba":              "crpnm",
    "College of Registered Psychiatric Nurses of Saskatchewan":          "crpns",
    # --- portal display-name aliases (names as they arrive in additional_content) ---
    # These are the short/variant names SMApply sends in regulatory_body_rn/lpn;
    # they point at the same abbreviations as their canonical entries above so
    # report matching (which keys off the abbrev token in the filename) is
    # unaffected. Add new variants here as other applications surface them.
    "NL College of Nurses":                                              "nlcn",
    "BC College of Nurses and Midwives":                                 "bccnm",
    "Nova Scotia College of Nursing":                                    "nscn",
    "PEI College of Nursing & Midwifery":                                "peicnm",
    "College of Licensed Practical Nurses and Health Care Aides of Alberta": "clha",
}

# All known abbreviations (longest-first) for robust filename matching.
_KNOWN_ABBREVS = sorted(set(RB_ABBREV.values()), key=len, reverse=True)


def abbrev_for(rb_name: str) -> str:
    """Abbreviation for a regulatory body name (exact, then case-insensitive)."""
    if rb_name in RB_ABBREV:
        return RB_ABBREV[rb_name]
    low = {k.lower(): v for k, v in RB_ABBREV.items()}
    return low.get((rb_name or "").strip().lower(), "")


def province_for(rb_name: str) -> str:
    """Province for a regulatory body name (exact, then case-insensitive)."""
    if rb_name in RB_PROVINCE_MAP:
        return RB_PROVINCE_MAP[rb_name]
    low = {k.lower(): v for k, v in RB_PROVINCE_MAP.items()}
    return low.get((rb_name or "").strip().lower(), "")


def regulator_of_report(filename: str) -> Optional[str]:
    """
    Identify which regulator a report file belongs to by the abbreviation token
    embedded in its name, e.g. 'NNAS-123_bccnm_latest.pdf' -> 'bccnm'.
    Tokens are matched on [-_. ] boundaries; longest abbreviation wins.
    """
    stem = os.path.splitext(filename or "")[0].lower()
    tokens = set(re.split(r"[-_.\s]+", stem))
    for ab in _KNOWN_ABBREVS:          # longest-first
        if ab in tokens:
            return ab
    return None


def designation_of_report(filename: str) -> str:
    """
    Pull the designation token (RN / LPN / RPN) from a report filename, e.g.
    'NNAS-T-..._BCCNM_LPN_..._latest.pdf' -> 'LPN'. Returns '' if none present.
    """
    stem = os.path.splitext(filename or "")[0].upper()
    tokens = set(re.split(r"[-_.\s]+", stem))
    for d in REPORT_DESIGNATIONS:
        if d in tokens:
            return d
    return ""


# =========================================================
# VALUE / HTML HELPERS
# =========================================================

_TAG_RE      = re.compile(r"<[^>]+>")
_LI_CLOSE_RE = re.compile(r"</li\s*>", re.IGNORECASE)


def strip_html(s: Any) -> str:
    if not s:
        return ""
    s = str(s).replace("<br/>", "\n").replace("<br />", "\n").replace("<br>", "\n")
    s = _LI_CLOSE_RE.sub("\n", s)
    s = _TAG_RE.sub("", s)
    return "\n".join(ln.strip() for ln in s.splitlines() if ln.strip()).strip()


def safe_str(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, dict):
        for k in ("label", "value", "name", "text"):
            if k in v and v[k] is not None:
                return strip_html(str(v[k]).strip())
        return ""
    return strip_html(str(v).strip())


def yn(value: str) -> str:
    v = (value or "").strip().lower()
    if v in ("yes", "y", "true"):
        return "Y"
    if v in ("no", "n", "false"):
        return "N"
    return (value or "").strip()


def gender_code(value: str) -> str:
    v = (value or "").strip().lower()
    if v.startswith("female") or v == "f":
        return "F"
    if v.startswith("male") or v == "m":
        return "M"
    return (value or "").strip()


_MONTHS = {m: f"{i:02d}" for i, m in enumerate(
    ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
     "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], start=1)}


def to_ddmmyyyy(value: Any) -> str:
    """Normalize many date forms to dd-mm-yyyy."""
    if isinstance(value, datetime):
        return value.strftime("%d-%m-%Y")
    v = (value or "").strip()
    if not v:
        return ""
    m = re.match(r"^(\d{4})\s+([A-Za-z]{3})[a-z]*\s+(\d{1,2})$", v)        # 1977 Jul 1
    if m and m.group(2)[:3].title() in _MONTHS:
        return f"{int(m.group(3)):02d}-{_MONTHS[m.group(2)[:3].title()]}-{m.group(1)}"
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", v)                           # ISO 8601
    if m:
        return f"{m.group(3)}-{m.group(2)}-{m.group(1)}"
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", v)                      # M/D/YYYY (US)
    if m:
        return f"{int(m.group(2)):02d}-{int(m.group(1)):02d}-{m.group(3)}"
    m = re.match(r"^(\d{4})[/](\d{1,2})[/](\d{1,2})$", v)                  # YYYY/MM/DD
    if m:
        return f"{int(m.group(3)):02d}-{int(m.group(2)):02d}-{m.group(1)}"
    m = re.match(r"^(\d{1,2})-(\d{1,2})-(\d{4})$", v)                      # dd-mm-yyyy
    if m:
        return f"{int(m.group(1)):02d}-{int(m.group(2)):02d}-{m.group(3)}"
    return v


def iso_code(country: str) -> str:
    return ISO_NUMERIC.get((country or "").strip().upper(), "")


# =========================================================
# ADDITIONAL_CONTENT PARSING (tolerant, with JSON repair)
# =========================================================

# Characters that may legally follow a closing quote in JSON (after whitespace).
_JSON_AFTER_STRING = set(',:}]')


def repair_json_quotes(s: str) -> str:
    """
    Repair the two malformations SMApply emits in `additional_content`:

      1. Unescaped double quotes INSIDE string values. Rich-text assessment
         notes routinely contain them, e.g.
             "...uses the term "endorsement" when a license is granted..."
         which terminates the string early and blows up json.loads.
      2. Raw control characters (newline / tab / carriage return) inside string
         values, which are also invalid JSON.

    The scanner tracks whether it is inside a string. When it meets a quote
    while inside one, it looks ahead past whitespace: if the next meaningful
    character is a JSON structural delimiter (, : } ]) or end-of-input, the
    quote genuinely closes the string; otherwise it is an inner quote and gets
    escaped. Backslash escapes are passed through untouched, so already-valid
    input round-trips unchanged.
    """
    out: List[str] = []
    in_str = False
    i, n = 0, len(s)

    while i < n:
        ch = s[i]

        if not in_str:
            out.append(ch)
            if ch == '"':
                in_str = True
            i += 1
            continue

        # --- inside a string ---
        if ch == "\\":
            out.append(ch)
            if i + 1 < n:
                out.append(s[i + 1])
                i += 2
            else:
                i += 1
            continue

        if ch == '"':
            j = i + 1
            while j < n and s[j] in " \t\r\n":
                j += 1
            if j >= n or s[j] in _JSON_AFTER_STRING:
                out.append(ch)          # genuine string terminator
                in_str = False
            else:
                out.append('\\"')       # stray inner quote -> escape it
            i += 1
            continue

        if ch == "\n":
            out.append("\\n")
        elif ch == "\r":
            out.append("\\r")
        elif ch == "\t":
            out.append("\\t")
        elif ord(ch) < 0x20:
            out.append(f"\\u{ord(ch):04x}")
        else:
            out.append(ch)
        i += 1

    return "".join(out)


def _try_loads(s: str) -> Tuple[Optional[Any], Optional[str]]:
    """json.loads with one level of double-encoding unwrapped."""
    try:
        raw = json.loads(s)
        if isinstance(raw, str):
            raw = json.loads(raw)
        return raw, None
    except Exception as e:
        return None, f"{type(e).__name__}: {e}"


def parse_additional_content(additional_content: Optional[str]) -> Tuple[Dict[str, Any], Dict[str, Any]]:
    """
    Parse the en-masse webhook blob, repairing malformed JSON if needed.

    Attempts, in order:
        raw  ->  unwrapped (surrounding quote pair stripped)
             ->  repaired  ->  unwrapped + repaired
    The first attempt yielding a dict wins. `debug["method"]` records which one
    succeeded so a repair is visible in the run output rather than silent.
    """
    dbg: Dict[str, Any] = {"present": bool(additional_content),
                           "type": str(type(additional_content))}

    if not additional_content:
        return {}, {**dbg, "parsed": False}

    if isinstance(additional_content, dict):
        return additional_content, {**dbg, "parsed": True, "method": "already_dict"}

    if not isinstance(additional_content, str):
        return {}, {**dbg, "parsed": False, "error": "unexpected type"}

    s = additional_content.strip()
    unwrapped = None
    if len(s) >= 2 and ((s[0] == '"' and s[-1] == '"') or (s[0] == "'" and s[-1] == "'")):
        unwrapped = s[1:-1]

    attempts: List[Tuple[str, str]] = [("raw", s)]
    if unwrapped is not None:
        attempts.append(("unwrapped", unwrapped))
    attempts.append(("repaired", repair_json_quotes(s)))
    if unwrapped is not None:
        attempts.append(("unwrapped+repaired", repair_json_quotes(unwrapped)))

    errors: Dict[str, str] = {}
    for method, candidate in attempts:
        raw, err = _try_loads(candidate)
        if err:
            errors[method] = err[:300]
            continue
        if not isinstance(raw, dict):
            errors[method] = f"not a dict: {type(raw)}"
            continue
        return raw, {**dbg,
                     "parsed": True,
                     "method": method,
                     "repaired": method.endswith("repaired"),
                     "keys": sorted(raw.keys())[:12],
                     "prior_errors": errors or None}

    return {}, {**dbg, "parsed": False, "errors": errors}


# =========================================================
# APPLICATION-SUMMARY PDF PARSER  (fixed label/value layout)
# =========================================================

PROMPTS: List[Tuple[str, str]] = [
    ("Select the regulatory body or bodies", "rb_selected"),
    ("Salutation, if any",                "salutation"),
    ("First, or Given Name",              "first_name"),
    ("Middle, or Second name",            "middle_name"),
    ("Last, or Family name",              "last_name"),
    ("Other, or Previous name(s)",        "other_names"),
    ("Gender Identity",                   "gender"),
    ("Date of birth",                     "dob"),
    ("Email address",                     "email"),
    ("Phone Number Type",                 "phone_type"),
    ("Phone Number",                      "phone"),
    ("Alternative Phone Number Type",     "alt_phone_type"),
    ("Alternate Phone Number",            "alt_phone"),
    ("Is your mailing address",           "mailing_same"),
    ("Which country did you receive your initial nursing education", "initial_edu_country"),
    ("What language was your Theory instruction delivered in",   "theory_language"),
    ("What language was your Clinical instruction delivered in", "clinical_language"),
    ("What is the primary language of the school",               "school_language"),
    ("Did you obtain a nursing license?", "obtained_license"),
    ("Does the license or registration make you eligible to practice", "eligible_initial_jur"),
    ("In which country were you registered or licensed", "reg_country"),
    ("What is the name of the regulatory body that issued the license", "reg_body_name"),
    ("What name appears on your registration or license", "reg_name_used"),
    ("What is the title of your registration or license", "reg_title"),
    ("What is your nursing registration or license number", "reg_number"),
    ("What is the current status of your registration or license", "reg_status"),
    ("On what date was this registration or license issued", "reg_issue_date"),
    ("What is the method by which you were registered or licensed", "reg_method"),
    ("Is this a lifetime practice registration or license", "reg_lifetime"),
    ("Are you eligible to renew this registration or license", "reg_renew_eligible"),
    ("Do you have any current limitations or restrictions", "reg_limitations"),
    ("Has your nursing registration or license ever been encumbered", "reg_ever_suspended"),
    ("Have you ever been registered to practice in any other health disciplines", "other_health_prof"),
    ("Have you ever been registered to practice in any other regulated professions", "other_regulated_prof"),
    ("Do you currently hold, or have you held a nursing license in any jurisdiction in Canada", "cur_canadian_reg"),
    ("In how many other Canadian Provinces or Territories have you applied", "num_other_provinces"),
    ("Have you ever been referred for a competency assessment", "ref_competency"),
    ("Have you completed a competency assessment", "complete_competency"),
    ("Have you ever been convicted of a criminal offense", "criminal_offense"),
    ("Do you suffer from or live with any physical or mental conditions", "addiction_impairing"),
    ("Have you ever been the subject of a finding of professional misconduct", "prof_misconduct_ever"),
    ("Are you the subject of a current investigation, inquiry, or proceeding", "misconduct_invest_now"),
    ("Are you the subject of a current proceeding in respect to any offence", "any_cur_proceeding"),
]
ID_DOC_TYPE_PROMPT = "- Document Type"
_NOISE_PREFIXES = ("i ", "*")
_PAGENUM_RE = re.compile(r"^\d+\s*/\s*\d+$")
_NO_RESPONSE = "(No response)"
_INLINE = {"Address line 1": "cur_addr1", "Address line 2": "cur_addr2",
           "City": "cur_city", "Province/State": "cur_province",
           "Postal/Zip code": "cur_postal", "Country": "cur_country"}


def _pdf_lines(pdf_path: str, max_pages: int = 22) -> List[str]:
    out: List[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        for pg in pdf.pages[:max_pages]:
            for raw in (pg.extract_text() or "").split("\n"):
                s = raw.strip()
                if s and not _PAGENUM_RE.match(s):
                    out.append(s)
    return out


def _clean(v: Optional[str]) -> str:
    v = (v or "").strip()
    return "" if v == _NO_RESPONSE else v


def _is_value_line(s: str) -> bool:
    return bool(s) and not any(s.startswith(p) for p in _NOISE_PREFIXES)


def _looks_help(s: str) -> bool:
    if not s or s.startswith("i ") or s.startswith("*"):
        return True
    if s.endswith("?") or s.endswith(","):
        return True
    if "@" in s or re.match(r"^[+(]?\d[\d\s\-()]{4,}$", s):
        return False
    toks = s.split()
    first = toks[0] if toks else ""
    if first[:1].islower():
        return True
    if first in ("If", "Enter", "Choose", "Select", "For", "Note", "Please",
                 "When", "Spell", "documents.", "other", "which", "nursing"):
        return True
    markers = (" you ", "your ", "registration/license", "jurisdiction",
               "exactly as it appears", "dropdown", "text box", "nursing education")
    if s.endswith(".") and any(m in s for m in markers):
        return True
    return False


def _next_value(lines: List[str], i: int) -> Optional[str]:
    in_help = False
    for j in range(i + 1, min(i + 12, len(lines))):
        s = lines[j]
        if not _is_value_line(s) or _looks_help(s):
            in_help = True
            continue
        if in_help and s.endswith(".") and len(s.split()) > 2:
            continue
        return s
    return None


def _starts_prompt(s: str) -> bool:
    """True if `s` is the start of another form prompt or inline label."""
    return any(s.startswith(p.split("\n")[0]) for p, _ in PROMPTS) or \
           any(s.startswith(lab) for lab in _INLINE)


def _values_after(lines: List[str], i: int, marker: str,
                  max_scan: int = 6, max_values: int = 12) -> List[str]:
    """
    ALL value lines following `marker` -- multi-select answers put each chosen
    option on its own line, so returning only the first (the old behaviour)
    silently dropped every regulator after the first one.
    Collection stops at the first non-value / help line, or at the start of
    the next form prompt (otherwise the following question and its answer
    would be swallowed into the list).
    """
    out: List[str] = []
    for j in range(i + 1, min(i + max_scan, len(lines))):
        if not lines[j].startswith(marker):
            continue
        tail = lines[j][len(marker):].strip()
        if tail:
            out.append(_clean(tail))
        for k in range(j + 1, min(j + 1 + max_values, len(lines))):
            s = lines[k]
            if not _is_value_line(s) or _looks_help(s) or _starts_prompt(s):
                break
            out.append(_clean(s))
        break
    return [v for v in out if v]


def parse_application_pdf(pdf_path: str) -> Dict[str, str]:
    lines = _pdf_lines(pdf_path)
    out: Dict[str, str] = {}
    for i, line in enumerate(lines):
        for lab, key in _INLINE.items():
            if line.startswith(lab) and len(line) > len(lab) and key not in out:
                out[key] = _clean(line[len(lab):].strip())
        for prompt, key in PROMPTS:
            if key in out:
                continue
            if not line.startswith(prompt.split("\n")[0]):
                continue
            if key == "phone" and line.startswith("Phone Number Type"):
                continue
            if key == "alt_phone" and line.startswith("Alternative Phone Number Type"):
                continue
            if key == "phone_type" and not line.startswith("Phone Number Type"):
                continue
            if key == "rb_selected":
                out[key] = ", ".join(_values_after(lines, i, "Responses Selected:"))
                continue
            val = _next_value(lines, i)
            if val is not None:
                out[key] = _clean(val)
    id_types = [_next_value(lines, idx) for idx, ln in enumerate(lines)
                if ln.endswith(ID_DOC_TYPE_PROMPT)]
    if id_types:
        out["id_doc1_type"] = _clean(id_types[0] or "")
    if len(id_types) >= 2:
        out["id_doc2_type"] = _clean(id_types[1] or "")
    return out


# =========================================================
# SHAREPOINT / MICROSOFT GRAPH HELPERS
# =========================================================

def ctx():
    """
    Compatibility shim. The rest of the script still passes a `sp` argument
    into SharePoint helper functions, but Graph auth is handled globally.
    """
    return None


class GraphFileRef:
    """Small compatibility object for code that expects file_obj.properties."""
    def __init__(self, properties: Optional[Dict[str, Any]] = None):
        self.properties = properties or {}


def _library_relative_path(path: str) -> str:
    """
    Convert old SharePoint server-relative paths into paths relative to the
    Applications document library root, which is how Microsoft Graph addresses
    the drive.

    Accepted examples:
      /sites/NNAS-Portal-Sharepoint/Applications/CES/NNAS-...
      /sites/NNAS-Portal-Sharepoint/CES/NNAS-...
      Applications/CES/NNAS-...
      CES/NNAS-...
    """
    p = (path or "").strip().replace("\\", "/").strip()

    if p.startswith(SITE_URL):
        p = p[len(SITE_URL):]

    p = p.strip("/")

    site_prefix = SITE_PREFIX.strip("/")
    if p.lower().startswith(site_prefix.lower() + "/"):
        p = p[len(site_prefix):].strip("/")

    if p.lower().startswith("applications/"):
        p = p[len("Applications/"):]

    return p.strip("/")


def _display_server_relative_path(library_relative_path: str) -> str:
    """
    Keep return values compatible with the old office365 client by returning a
    server-relative-looking path. The helper functions convert it back to a
    library-relative Graph path when downloading.
    """
    p = _library_relative_path(library_relative_path)
    return f"{SITE_PREFIX}/{p}".replace("//", "/")


def _encode_graph_path(path: str) -> str:
    return quote(path.strip("/"), safe="/")


def _get_graph_token() -> str:
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


def _graph_request(method: str, endpoint: str, expected_statuses=(200,), **kwargs):
    token = _get_graph_token()

    headers = kwargs.pop("headers", {})
    headers["Authorization"] = f"Bearer {token}"

    response = requests.request(
        method,
        f"{GRAPH_BASE}{endpoint}",
        headers=headers,
        timeout=180,
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


def _graph_get_or_none(endpoint: str):
    response = requests.get(
        f"{GRAPH_BASE}{endpoint}",
        headers={"Authorization": f"Bearer {_get_graph_token()}"},
        timeout=180,
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


def _init_sharepoint_graph() -> None:
    global _SP_SITE_ID, _SP_DRIVE_ID, _SP_LIST_ID

    if _SP_SITE_ID and _SP_DRIVE_ID and _SP_LIST_ID:
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

    drive = next((d for d in drives if d.get("name") == DOCUMENT_LIBRARY_NAME), None)
    if not drive:
        available = [d.get("name") for d in drives]
        raise RuntimeError(
            f"Could not find document library drive '{DOCUMENT_LIBRARY_NAME}'. "
            f"Available drives: {available}"
        )
    _SP_DRIVE_ID = drive["id"]

    lists = _graph_request(
        "GET",
        f"/sites/{_SP_SITE_ID}/lists?$select=id,displayName",
        expected_statuses=(200,),
    )["value"]

    sp_list = next((l for l in lists if l.get("displayName") == DOCUMENT_LIBRARY_NAME), None)
    if not sp_list:
        available = [l.get("displayName") for l in lists]
        raise RuntimeError(
            f"Could not find SharePoint list '{DOCUMENT_LIBRARY_NAME}'. "
            f"Available lists: {available}"
        )
    _SP_LIST_ID = sp_list["id"]


def _get_drive_item_by_path(drive_relative_path: str):
    _init_sharepoint_graph()
    clean_path = _library_relative_path(drive_relative_path)

    if not clean_path:
        return _graph_request(
            "GET",
            f"/drives/{_SP_DRIVE_ID}/root",
            expected_statuses=(200,),
        )

    return _graph_get_or_none(
        f"/drives/{_SP_DRIVE_ID}/root:/{_encode_graph_path(clean_path)}"
    )


def _get_drive_item_content(item_id: str) -> bytes:
    _init_sharepoint_graph()
    response = requests.get(
        f"{GRAPH_BASE}/drives/{_SP_DRIVE_ID}/items/{item_id}/content",
        headers={"Authorization": f"Bearer {_get_graph_token()}"},
        timeout=300,
        allow_redirects=True,
    )
    if response.status_code != 200:
        raise RuntimeError(
            f"Failed downloading SharePoint item content.\n"
            f"Item ID: {item_id}\n"
            f"Status: {response.status_code}\n"
            f"Response body:\n{response.text[:2000]}"
        )
    return response.content


def _get_drive_item_with_fields(item_id: str) -> Dict[str, Any]:
    _init_sharepoint_graph()

    item = _graph_request(
        "GET",
        f"/drives/{_SP_DRIVE_ID}/items/{item_id}?$expand=listItem($expand=fields)",
        expected_statuses=(200,),
    )

    props: Dict[str, Any] = {}
    list_item = item.get("listItem") or {}
    fields = list_item.get("fields") or {}
    if isinstance(fields, dict):
        props.update(fields)

    item_name = item.get("name", "")
    if item_name:
        props.setdefault("FileLeafRef", item_name)
        props.setdefault("Name", item_name)

    if item.get("lastModifiedDateTime"):
        props.setdefault("Modified", item.get("lastModifiedDateTime"))
        props.setdefault("TimeLastModified", item.get("lastModifiedDateTime"))

    if item.get("createdDateTime"):
        props.setdefault("Created", item.get("createdDateTime"))

    return {
        "item": item,
        "props": props,
    }


def _create_folder(parent_path: str, folder_name: str):
    _init_sharepoint_graph()

    parent_path = _library_relative_path(parent_path)
    full_path = f"{parent_path}/{folder_name}".strip("/")

    existing = _get_drive_item_by_path(full_path)
    if existing:
        return existing

    if parent_path:
        endpoint = f"/drives/{_SP_DRIVE_ID}/root:/{_encode_graph_path(parent_path)}:/children"
    else:
        endpoint = f"/drives/{_SP_DRIVE_ID}/root/children"

    response = requests.post(
        f"{GRAPH_BASE}{endpoint}",
        headers={
            "Authorization": f"Bearer {_get_graph_token()}",
            "Content-Type": "application/json",
        },
        json={
            "name": folder_name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "fail",
        },
        timeout=180,
    )

    if response.status_code == 201:
        return response.json()

    if response.status_code == 409:
        existing_after_conflict = _get_drive_item_by_path(full_path)
        if existing_after_conflict:
            return existing_after_conflict

    raise RuntimeError(
        f"Microsoft Graph folder create failed.\n"
        f"Parent path: {parent_path}\n"
        f"Folder name: {folder_name}\n"
        f"Full path: {full_path}\n"
        f"Status: {response.status_code}\n"
        f"Response body:\n{response.text[:5000]}"
    )


def ensure_folder(sp, folder_rel: str) -> None:
    parts = [p for p in _library_relative_path(folder_rel).split("/") if p]
    current_path = ""

    for part in parts:
        _create_folder(current_path, part)
        current_path = f"{current_path}/{part}".strip("/")


def folder_exists(sp, folder_rel: str) -> bool:
    """True if the SharePoint folder exists (does not create it)."""
    item = _get_drive_item_by_path(folder_rel)
    return bool(item and item.get("folder") is not None)


def list_folder_files(sp, folder_rel: str, recursive: bool = True) -> List[Dict[str, Any]]:
    results: List[Dict[str, Any]] = []

    start_library_path = _library_relative_path(folder_rel)
    folder = _get_drive_item_by_path(start_library_path)
    if not folder or folder.get("folder") is None:
        return results

    def walk(folder_item_id: str, current_library_path: str):
        endpoint = f"/drives/{_SP_DRIVE_ID}/items/{folder_item_id}/children?$top=999"

        while endpoint:
            data = _graph_request("GET", endpoint, expected_statuses=(200,))
            for child in data.get("value", []):
                name = child.get("name", "")
                if not name:
                    continue

                child_library_path = f"{current_library_path.rstrip('/')}/{name}".strip("/")

                if child.get("file") is not None:
                    full = _get_drive_item_with_fields(child["id"])
                    item = full["item"]
                    props = full["props"]
                    modified = item.get("lastModifiedDateTime") or props.get("Modified") or ""
                    created = item.get("createdDateTime") or props.get("Created") or ""

                    results.append({
                        "file": GraphFileRef({"TimeLastModified": modified}),
                        "props": props,
                        "name": name,
                        "url": _display_server_relative_path(child_library_path),
                        "modified": modified,
                        "created": created,
                        "id": child.get("id", ""),
                    })

                elif recursive and child.get("folder") is not None and name != "Forms":
                    walk(child["id"], child_library_path)

            next_link = data.get("@odata.nextLink")
            if next_link:
                if next_link.startswith(GRAPH_BASE):
                    endpoint = next_link[len(GRAPH_BASE):]
                else:
                    raise RuntimeError(f"Unexpected Microsoft Graph nextLink: {next_link}")
            else:
                endpoint = None

    walk(folder["id"], start_library_path)
    return results


def download_file(sp, server_rel: str, local_path: str) -> None:
    item = _get_drive_item_by_path(server_rel)
    if not item:
        raise RuntimeError(f"SharePoint file not found: {server_rel}")

    content = _get_drive_item_content(item["id"])
    os.makedirs(os.path.dirname(local_path), exist_ok=True)
    with open(local_path, "wb") as fh:
        fh.write(content)


def upload_file(sp, local_path: str, folder_rel: str, filename: str) -> str:
    ensure_folder(sp, folder_rel)

    folder_path = _library_relative_path(folder_rel)
    upload_path = f"{folder_path}/{filename}".strip("/")

    with open(local_path, "rb") as fh:
        content = fh.read()

    uploaded = _graph_request(
        "PUT",
        f"/drives/{_SP_DRIVE_ID}/root:/{_encode_graph_path(upload_path)}:/content",
        headers={"Content-Type": "application/octet-stream"},
        data=content,
        expected_statuses=(200, 201),
    )

    return (uploaded or {}).get("webUrl") or _display_server_relative_path(upload_path)


def _scalar(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return "Yes" if v else "No"
    if isinstance(v, (list, tuple)):
        return ", ".join(_scalar(x) for x in v if x is not None)
    if isinstance(v, dict):
        for k in ("Label", "label", "value", "Value", "name", "Name", "text"):
            if v.get(k):
                return str(v[k]).strip()
        return ""
    return str(v).strip()


def meta_get(props: Dict[str, Any], display_name: Optional[str]) -> str:
    if not display_name:
        return ""

    cands = [
        display_name,
        display_name.replace(" ", "_x0020_"),
        display_name.replace(" ", "_x0020_").replace("-", "_x002d_"),
        display_name.replace(" ", ""),
    ]

    for k in cands:
        if k in props and props[k] not in (None, ""):
            return _scalar(props[k])

    low = {k.lower(): k for k in props}
    for k in cands:
        rk = low.get(k.lower())
        if rk and props[rk] not in (None, ""):
            return _scalar(props[rk])

    return ""


# =========================================================
# CATEGORIZATION + NAMING
# =========================================================

def categorize(props: Dict[str, Any], file_name: str, app_id: str,
               form: Dict[str, str]) -> Tuple[str, str, Optional[int]]:
    """Return (category, doctype, forced_number)."""
    base = os.path.splitext(file_name)[0]
    # base-folder ID docs -> identity, doctype from the parsed form
    if base.lower().startswith(app_id.lower()):
        m = ID_DOC_RE.search(base)
        if m:
            n = int(m.group(1))
            dt = form.get(f"id_doc{n}_type", "") or "ID DOCUMENT"
            return ("identity", dt.upper(), n)
    file_type = meta_get(props, "FileType").strip()
    category = FILETYPE_TO_CATEGORY.get(file_type.lower(), "other")
    doctype = ""
    for fld in DOCTYPE_FIELDS:
        doctype = meta_get(props, fld).strip()
        if doctype:
            break
    return (category, (doctype or file_type or category).upper(), None)


_ILLEGAL = re.compile(r'[\\/:*?"<>|]+')


def sftp_name(doc_id: str, app_id: str, category: str, doctype: str, number: int, ext: str) -> str:
    seg = _ILLEGAL.sub(" ", doctype).strip()
    return f"{doc_id}-{app_id}-{category}-{seg}{number}{ext}"


def doc_metadata_block(props: Dict[str, Any], local_path: str, file_obj: Any,
                       doctype: str) -> Dict[str, str]:
    country = meta_get(props, COLUMN_MAP["IssuedFrom"])
    try:
        size = str(os.path.getsize(local_path)) if local_path and os.path.exists(local_path) else ""
    except OSError:
        size = ""
    modified = ""
    if file_obj is not None and hasattr(file_obj, "properties"):
        modified = file_obj.properties.get("TimeLastModified", "")
    if not modified:
        modified = meta_get(props, "Modified")
    return {
        "DocumentType":     doctype,
        "DocumentNumber":   meta_get(props, COLUMN_MAP["DocumentNumber"]),
        "IssuedFrom":       country,
        "IssuedFromIso":    iso_code(country),
        "DateIssued":       to_ddmmyyyy(meta_get(props, COLUMN_MAP["DateIssued"])),
        "ReceivedDate":     to_ddmmyyyy(meta_get(props, COLUMN_MAP["ReceivedDate"])),
        "ExpirationDate":   to_ddmmyyyy(meta_get(props, COLUMN_MAP["ExpirationDate"])),
        "Size":             size,
        "Extension":        os.path.splitext(local_path or "")[1] or ".pdf",
        "LastModifiedDate": to_ddmmyyyy(modified),
    }


# =========================================================
# APPLICANT DATA MODEL  (form > additional_content)
# =========================================================

def first_nonempty(*vals: str) -> str:
    for v in vals:
        if v and str(v).strip():
            return str(v).strip()
    return ""


_DDMMYYYY_RE = re.compile(r"^\d{2}-\d{2}-\d{4}$")


def check_dob(form: Dict[str, str], payload: Dict[str, Any]) -> Dict[str, Any]:
    """
    Cross-check the date of birth from its two sources, using the same
    precedence as build_applicant_model (PDF first, then personal_info.sa_dob).

    Returns a dict with:
      source      -> which source supplies the written DOB (or None)
      pdf_present / ac_present -> whether each source had a value
      match       -> True / False when both present, else None
      normalized  -> whether the written value is dd-mm-yyyy
      warnings    -> list of human-readable issues

    The DOB values themselves are deliberately NOT returned, so the run output
    (warnings + debug) carries no applicant date of birth.
    """
    pi = payload.get("personal_info") or {}
    pdf_norm = to_ddmmyyyy((form.get("dob") or "").strip())
    ac_norm  = to_ddmmyyyy(safe_str(pi.get("sa_dob")))
    written  = pdf_norm or ac_norm

    if pdf_norm:
        source = "application_summary_pdf"
    elif ac_norm:
        source = "additional_content.personal_info.sa_dob"
    else:
        source = None

    warns: List[str] = []
    match: Optional[bool] = None
    if pdf_norm and ac_norm:
        match = (pdf_norm == ac_norm)
        if not match:
            warns.append(
                "Date of birth mismatch: the application summary PDF and "
                "additional_content personal_info.sa_dob disagree after "
                "normalization. The PDF value was written to the CSV, XML and "
                "NNAS index. Verify against the applicant record."
            )

    normalized = bool(written) and bool(_DDMMYYYY_RE.match(written))
    if not written:
        warns.append(
            "Date of birth is blank in both the application summary PDF and "
            "additional_content. DateOfBirth and the NNAS index DOB column will be empty."
        )
    elif not normalized:
        warns.append(
            f"Date of birth from {source} is not in a recognized format and was "
            "written unnormalized (expected dd-mm-yyyy)."
        )

    return {
        "source": source,
        "pdf_present": bool(pdf_norm),
        "ac_present": bool(ac_norm),
        "match": match,
        "normalized": normalized,
        "warnings": warns,
    }


def build_applicant_model(form: Dict[str, str], payload: Dict[str, Any],
                          doc_id: str, app_id: str,
                          rb_name: Optional[str] = None) -> Dict[str, Any]:
    ai = payload.get("application_info") or {}
    pi = payload.get("personal_info") or {}
    reg = payload.get("registration") or {}
    e1 = reg.get("entry_1") or {}

    rn_list  = [x.strip() for x in safe_str(ai.get("regulatory_body_rn")).split(",") if x.strip()]
    lpn_list = [x.strip() for x in safe_str(ai.get("regulatory_body_lpn")).split(",") if x.strip()]

    # If a specific regulator is supplied (per-regulator package), use it;
    # otherwise fall back to the first listed body (RN list first, then LPN).
    rb_selected = (rb_name or "").strip() or \
        first_nonempty(*(rn_list + lpn_list), form.get("rb_selected", ""))

    # NursingGroup reflects THIS regulator's list membership (an application may
    # carry both RN and LPN regulators), falling back to the app-wide selection.
    low = rb_selected.lower()
    if any(low == r.lower() for r in rn_list):
        nursing_group = "RN"
    elif any(low == r.lower() for r in lpn_list):
        nursing_group = "LPN"
    else:
        nursing_group = safe_str(ai.get("nursing_group_selection"))

    province = province_for(rb_selected)

    same_mailing = first_nonempty(form.get("mailing_same"),
                                  safe_str(pi.get("sa_mailing_same_as_current_flag"))).lower().startswith("y")

    home = {
        "StreetAddress1": first_nonempty(form.get("cur_addr1"), safe_str(pi.get("sa_current_address_line1"))),
        "StreetAddress2": first_nonempty(form.get("cur_addr2"), safe_str(pi.get("sa_current_address_line2"))),
        "AptNo": "",
        "City": first_nonempty(form.get("cur_city"), safe_str(pi.get("sa_current_city"))),
        "State": first_nonempty(form.get("cur_province"), safe_str(pi.get("sa_current_province"))),
        "Country": first_nonempty(form.get("cur_country"), safe_str(pi.get("sa_current_country"))),
        "PostalCode": first_nonempty(form.get("cur_postal"), safe_str(pi.get("sa_current_postal_code"))),
    }
    mailing = dict(home) if same_mailing else {
        "StreetAddress1": safe_str(pi.get("sa_mailing_address_line1")),
        "StreetAddress2": safe_str(pi.get("sa_mailing_address_line2")),
        "AptNo": "",
        "City": safe_str(pi.get("sa_mailing_city")),
        "State": safe_str(pi.get("sa_mailing_province")),
        "Country": safe_str(pi.get("sa_mailing_country")),
        "PostalCode": safe_str(pi.get("sa_mailing_postal_code")),
    }

    lang = first_nonempty(form.get("theory_language"), form.get("school_language"),
                          safe_str(pi.get("sa_first_language")), "English")

    return {
        "ID": doc_id,
        "ApplicationID": app_id,
        "NursingGroup": nursing_group,
        "LanguageOfApplication": lang.upper(),
        "RBName": rb_selected,
        "RBAbbrev": abbrev_for(rb_selected),
        "RBProvince": province,
        "Personal": {
            "FirstName": first_nonempty(form.get("first_name"), safe_str(pi.get("sa_first_name"))),
            "MiddleName": first_nonempty(form.get("middle_name"), safe_str(pi.get("sa_middle_name"))),
            "LastName": first_nonempty(form.get("last_name"), safe_str(pi.get("sa_last_name"))),
            "OtherName": first_nonempty(form.get("other_names"), safe_str(pi.get("sa_other_names"))),
            "DateOfBirth": to_ddmmyyyy(first_nonempty(form.get("dob"), safe_str(pi.get("sa_dob")))),
            "Gender": gender_code(form.get("gender", "")),
            "SpokenLanguage": "",
            "WrittenLanguage": "",
            "FirstLanguage": lang,
            "InitialNursingEducationCountry": first_nonempty(
                form.get("initial_edu_country"), safe_str(pi.get("sa_initial_education_country"))),
            "HomeAddress": home,
            "MailingAddress": mailing,
            "PrimaryPhone": first_nonempty(form.get("phone"), safe_str(pi.get("sa_phone"))),
            "AlternatePhone": first_nonempty(form.get("alt_phone"), safe_str(pi.get("sa_alt_phone"))),
            "Email": first_nonempty(form.get("email"), safe_str(pi.get("sa_email"))),
        },
        "Registration": {
            "RegisteredInJurisdictionOfInitialProgram": yn(first_nonempty(
                form.get("eligible_initial_jur"),
                safe_str(reg.get("sa_eligible_to_practice_in_initial_edu_jurisdiction_flag")))),
            "DiplomaRightToPractice": yn(first_nonempty(
                form.get("diploma_right"),
                safe_str(reg.get("sa_diploma_grants_right_to_practice_flag")))),
            "DiplomaNumber": "",
            "LicenseTitle": first_nonempty(form.get("reg_title"), safe_str(e1.get("sa_registration_title_1"))),
            "LicenseNumber": first_nonempty(form.get("reg_number"), safe_str(e1.get("sa_registration_number_1"))),
            "LicenseStatus": first_nonempty(form.get("reg_status"), safe_str(e1.get("sa_registration_status_1"))),
            "LicenseIssueDate": to_ddmmyyyy(first_nonempty(
                form.get("reg_issue_date"), safe_str(e1.get("sa_registration_issued_date_1")))),
            "LicenseExpiredDate": to_ddmmyyyy(first_nonempty(
                safe_str(e1.get("sa_registration_expiry_date_1")), form.get("reg_expiry"))),
            "NameUsedWhenRegIssued": first_nonempty(
                form.get("reg_name_used"), safe_str(e1.get("sa_registrant_name_issued_1"))),
            "LicenseAuthorityName": first_nonempty(
                form.get("reg_body_name"), safe_str(e1.get("sa_jurisdiction_name_1"))),
            "AuthorityAddress": {
                "StreetAddress1": safe_str(e1.get("sa_jurisdiction_address_1")),
                "StreetAddress2": "", "POBox": "", "City": "", "State": "",
                "Country": first_nonempty(form.get("reg_country"), safe_str(e1.get("sa_jurisdiction_country_1"))),
                "PostalCode": "",
            },
            "RegMethod": first_nonempty(form.get("reg_method"), safe_str(e1.get("sa_registration_method_1"))),
            "EligibleInJurisdiction": yn(form.get("reg_renew_eligible", "")),
            "LifetimeRegistration": yn(first_nonempty(
                form.get("reg_lifetime"), safe_str(e1.get("sa_registration_lifetime_flag_1")))),
            "RegOtherHealthProf": yn(form.get("other_health_prof", "")),
            "OtherRegulatedProf": yn(form.get("other_regulated_prof", "")),
            "LicenseEverSuspended": yn(first_nonempty(
                form.get("reg_ever_suspended"), safe_str(e1.get("sa_registration_discipline_flag_1")))),
            "LicenseSuspendDetails": "",
            "DateOfficialSigned": "",
            "ReceivedDate": "",
        },
        "Conditions": {
            "CurRegInAnotherCanadianProv": yn(form.get("cur_canadian_reg", "")),
            "CurCanadianRegIssueDate": "",
            "CurCanadianRegProvName": "",
            "NumberAppliedOtherProvinces": "0",
            "ProvinceApplied": "",
            "DateApplied": "",
            "RefCompetencyAssessment": yn(form.get("ref_competency", "")),
            "CompleteCompetencyAssess": yn(form.get("complete_competency", "")),
            "CurLimitation": yn(form.get("reg_limitations", "")),
            "LicenseRevoked": yn(form.get("reg_ever_suspended", "")),
            "CriminalOffense": yn(form.get("criminal_offense", "")),
            "AddictionImpairing": yn(form.get("addiction_impairing", "")),
            "ProfMisconductEver": yn(form.get("prof_misconduct_ever", "")),
            "MisconductInvestNow": yn(form.get("misconduct_invest_now", "")),
            "AnyCurProceeding": yn(form.get("any_cur_proceeding", "")),
        },
    }


# =========================================================
# XML SERIALIZER  (matches the NNAS <Applicant> structure)
# =========================================================

def _x(v: Any) -> str:
    return _xml_escape("" if v is None else str(v))


def _file_xml(doc: Dict[str, str], indent: str) -> str:
    return (
        f"{indent}<Files>\n"
        f"{indent}  <File>\n"
        f"{indent}    <Name>{_x(doc['Name'])}</Name>\n"
        f"{indent}    <Path>{_x(doc['Path'])}</Path>\n"
        f"{indent}    <Extension>{_x(doc['Extension'])}</Extension>\n"
        f"{indent}    <Size>{_x(doc['Size'])}</Size>\n"
        f"{indent}    <LastModifiedDate>{_x(doc['LastModifiedDate'])}</LastModifiedDate>\n"
        f"{indent}  </File>\n"
        f"{indent}</Files>\n"
    )


def _doc_xml(doc: Dict[str, str], indent: str, with_expiry: bool = True) -> str:
    s = (
        f"{indent}<Document>\n"
        f"{indent}  <DocumentType>{_x(doc['DocumentType'])}</DocumentType>\n"
        f"{indent}  <DocumentNumber>{_x(doc['DocumentNumber'])}</DocumentNumber>\n"
        f"{indent}  <IssuedFrom iso-code=\"{_x(doc['IssuedFromIso'])}\">{_x(doc['IssuedFrom'])}</IssuedFrom>\n"
        f"{indent}  <DateIssued>{_x(doc['DateIssued'])}</DateIssued>\n"
        f"{indent}  <ReceivedDate>{_x(doc['ReceivedDate'])}</ReceivedDate>\n"
    )
    if with_expiry:
        s += f"{indent}  <ExpirationDate>{_x(doc['ExpirationDate'])}</ExpirationDate>\n"
    s += _file_xml(doc, indent + "  ")
    s += f"{indent}</Document>\n"
    return s


def build_xml(records: List[Tuple[Dict[str, Any], Dict[str, List[Dict[str, str]]]]]) -> str:
    """
    Serialize one <Applicant> with one <Application> per record. A record is a
    (model, docs) pair; multiple records occur when a regulator assessed the
    applicant under more than one designation (e.g. RN and LPN), each carrying
    its own NursingGroup and its own report document.
    """
    out = StringIO()
    w = out.write
    w('<?xml version="1.0" encoding="UTF-8"?>\n')
    w("<Applicant>\n")
    w(f"  <ID>{_x(records[0][0]['ID'])}</ID>\n")
    w("  <Applications>\n")
    for model, docs in records:
        _application_xml(w, model, docs)
    w("  </Applications>\n")
    w("</Applicant>\n")
    return out.getvalue()


def _application_xml(w, model: Dict[str, Any], docs: Dict[str, List[Dict[str, str]]]) -> None:
    p = model["Personal"]
    r = model["Registration"]
    c = model["Conditions"]
    h, ma = p["HomeAddress"], p["MailingAddress"]
    aa = r["AuthorityAddress"]

    w("    <Application>\n")
    w(f"      <ApplicationID>{_x(model['ApplicationID'])}</ApplicationID>\n")
    w(f"      <NursingGroup>{_x(model['NursingGroup'])}</NursingGroup>\n")
    w(f"      <LanguageOfApplication>{_x(model['LanguageOfApplication'])}</LanguageOfApplication>\n")
    w(f"      <RBName>{_x(model['RBName'])}</RBName>\n")
    w(f"      <RBAbbrev>{_x(model['RBAbbrev'])}</RBAbbrev>\n")
    w(f"      <RBProvince>{_x(model['RBProvince'])}</RBProvince>\n")
    # personal
    w("      <PersonalInfo>\n")
    w(f"        <FirstName>{_x(p['FirstName'])}</FirstName>\n")
    w(f"        <MiddleName>{_x(p['MiddleName'])}</MiddleName>\n")
    w(f"        <LastName>{_x(p['LastName'])}</LastName>\n")
    w(f"        <OtherName>{_x(p['OtherName'])}</OtherName>\n")
    w(f"        <DateOfBirth>{_x(p['DateOfBirth'])}</DateOfBirth>\n")
    w(f"        <Gender>{_x(p['Gender'])}</Gender>\n")
    w("        <Languages>\n")
    w(f"          <SpokenLanguage>{_x(p['SpokenLanguage'])}</SpokenLanguage>\n")
    w(f"          <WrittenLanguage>{_x(p['WrittenLanguage'])}</WrittenLanguage>\n")
    w(f"          <FirstLanguage>{_x(p['FirstLanguage'])}</FirstLanguage>\n")
    w("        </Languages>\n")
    w(f"        <InitialNursingEducationCountry iso-code=\"{_x(iso_code(p['InitialNursingEducationCountry']))}\">"
      f"{_x(p['InitialNursingEducationCountry'])}</InitialNursingEducationCountry>\n")
    w("        <Addresses>\n")
    for tag, a in (("HomeAddress", h), ("MailingAddress", ma)):
        w(f"          <{tag}>\n")
        w(f"            <StreetAddress1>{_x(a['StreetAddress1'])}</StreetAddress1>\n")
        w(f"            <StreetAddress2>{_x(a['StreetAddress2'])}</StreetAddress2>\n")
        w(f"            <AptNo>{_x(a['AptNo'])}</AptNo>\n")
        w(f"            <City>{_x(a['City'])}</City>\n")
        w(f"            <State>{_x(a['State'])}</State>\n")
        w(f"            <Country iso-code=\"{_x(iso_code(a['Country']))}\">{_x(a['Country'])}</Country>\n")
        w(f"            <PostalCode>{_x(a['PostalCode'])}</PostalCode>\n")
        w(f"          </{tag}>\n")
    w("        </Addresses>\n")
    w("        <ContactInfo>\n")
    w(f"          <PrimaryPhone>{_x(p['PrimaryPhone'])}</PrimaryPhone>\n")
    w(f"          <AlternatePhone>{_x(p['AlternatePhone'])}</AlternatePhone>\n")
    w(f"          <Email>{_x(p['Email'])}</Email>\n")
    w("        </ContactInfo>\n")
    w("      </PersonalInfo>\n")
    w("      <NursingEducation></NursingEducation>\n")
    # registrations
    w("      <NursingRegistrations>\n")
    w(f"        <RegisteredInJurisdictionOfInitialProgram>{_x(r['RegisteredInJurisdictionOfInitialProgram'])}</RegisteredInJurisdictionOfInitialProgram>\n")
    w(f"        <DiplomaRightToPractice>{_x(r['DiplomaRightToPractice'])}</DiplomaRightToPractice>\n")
    w(f"        <DiplomaNumber>{_x(r['DiplomaNumber'])}</DiplomaNumber>\n")
    w("        <NursingRegistration>\n")
    w(f"          <LicenseTitle>{_x(r['LicenseTitle'])}</LicenseTitle>\n")
    w(f"          <LicenseNumber>{_x(r['LicenseNumber'])}</LicenseNumber>\n")
    w(f"          <LicenseStatus>{_x(r['LicenseStatus'])}</LicenseStatus>\n")
    w(f"          <LicenseIssueDate>{_x(r['LicenseIssueDate'])}</LicenseIssueDate>\n")
    w(f"          <LicenseExpiredDate>{_x(r['LicenseExpiredDate'])}</LicenseExpiredDate>\n")
    w(f"          <NameUsedWhenRegIssued>{_x(r['NameUsedWhenRegIssued'])}</NameUsedWhenRegIssued>\n")
    w(f"          <LicenseAuthorityName>{_x(r['LicenseAuthorityName'])}</LicenseAuthorityName>\n")
    w("          <Address>\n")
    w(f"            <StreetAddress1>{_x(aa['StreetAddress1'])}</StreetAddress1>\n")
    w(f"            <StreetAddress2>{_x(aa['StreetAddress2'])}</StreetAddress2>\n")
    w(f"            <POBox>{_x(aa['POBox'])}</POBox>\n")
    w(f"            <City>{_x(aa['City'])}</City>\n")
    w(f"            <State>{_x(aa['State'])}</State>\n")
    w(f"            <Country iso-code=\"{_x(iso_code(aa['Country']))}\">{_x(aa['Country'])}</Country>\n")
    w(f"            <PostalCode>{_x(aa['PostalCode'])}</PostalCode>\n")
    w("          </Address>\n")
    w(f"          <RegMethod>{_x(r['RegMethod'])}</RegMethod>\n")
    w(f"          <EligibleInJurisdiction>{_x(r['EligibleInJurisdiction'])}</EligibleInJurisdiction>\n")
    w(f"          <LifetimeRegistration>{_x(r['LifetimeRegistration'])}</LifetimeRegistration>\n")
    w(f"          <RegOtherHealthProf>{_x(r['RegOtherHealthProf'])}</RegOtherHealthProf>\n")
    w(f"          <OtherRegulatedProf>{_x(r['OtherRegulatedProf'])}</OtherRegulatedProf>\n")
    w(f"          <LicenseEverSuspended>{_x(r['LicenseEverSuspended'])}</LicenseEverSuspended>\n")
    w(f"          <LicenseSuspendDetails>{_x(r['LicenseSuspendDetails'])}</LicenseSuspendDetails>\n")
    w(f"          <DateOfficialSigned>{_x(r['DateOfficialSigned'])}</DateOfficialSigned>\n")
    w(f"          <ReceivedDate>{_x(r['ReceivedDate'])}</ReceivedDate>\n")
    w("          <Documents>\n")
    for d in docs.get("registration", []):
        w(_doc_xml(d, "            ", with_expiry=True))
    w("          </Documents>\n")
    w("        </NursingRegistration>\n")
    w("      </NursingRegistrations>\n")
    w("      <NursingExaminations></NursingExaminations>\n")
    # conditions
    w("      <NursingConditions>\n")
    w("        <NursingCondition>\n")
    w(f"          <CurRegInAnotherCanadianProv>{_x(c['CurRegInAnotherCanadianProv'])}</CurRegInAnotherCanadianProv>\n")
    w(f"          <CurCanadianRegIssueDate>{_x(c['CurCanadianRegIssueDate'])}</CurCanadianRegIssueDate>\n")
    w(f"          <CurCanadianRegProvName>{_x(c['CurCanadianRegProvName'])}</CurCanadianRegProvName>\n")
    w(f"          <NumberAppliedOtherProvinces>{_x(c['NumberAppliedOtherProvinces'])}</NumberAppliedOtherProvinces>\n")
    w("          <AppOtherProvRegs>\n")
    w("            <AppOtherProvReg>\n")
    w(f"              <ProvinceApplied>{_x(c['ProvinceApplied'])}</ProvinceApplied>\n")
    w(f"              <DateApplied>{_x(c['DateApplied'])}</DateApplied>\n")
    w("            </AppOtherProvReg>\n")
    w("          </AppOtherProvRegs>\n")
    w(f"          <RefCompetencyAssessment>{_x(c['RefCompetencyAssessment'])}</RefCompetencyAssessment>\n")
    w(f"          <CompleteCompetencyAssess>{_x(c['CompleteCompetencyAssess'])}</CompleteCompetencyAssess>\n")
    w("          <CurLimitations>\n")
    w(f"            <CurLimitation>{_x(c['CurLimitation'])}</CurLimitation>\n")
    w("            <CurLimitationDetails></CurLimitationDetails>\n")
    w("            <CurLimitationDate></CurLimitationDate>\n")
    w("          </CurLimitations>\n")
    w("          <LicenseRevokes>\n")
    w(f"            <LicenseRevoked>{_x(c['LicenseRevoked'])}</LicenseRevoked>\n")
    w("            <LicenseRevokedDetails></LicenseRevokedDetails>\n")
    w("            <LicenseRevokedDate></LicenseRevokedDate>\n")
    w("          </LicenseRevokes>\n")
    w("          <CriminalOffenses>\n")
    w(f"            <CriminalOffense>{_x(c['CriminalOffense'])}</CriminalOffense>\n")
    w("            <CriminalOffenseDetails></CriminalOffenseDetails>\n")
    w("            <CriminalOffenseDate></CriminalOffenseDate>\n")
    w("          </CriminalOffenses>\n")
    w("          <AddictionImpairings>\n")
    w(f"            <AddictionImpairing>{_x(c['AddictionImpairing'])}</AddictionImpairing>\n")
    w("            <AddictionImpairingDetails></AddictionImpairingDetails>\n")
    w("            <AddictionImpairingDate></AddictionImpairingDate>\n")
    w("          </AddictionImpairings>\n")
    w("          <ProfMisconductsEver>\n")
    w(f"            <ProfMisconductEver>{_x(c['ProfMisconductEver'])}</ProfMisconductEver>\n")
    w("            <ProfMisconductEverDetails></ProfMisconductEverDetails>\n")
    w("            <ProfMisconductEverDate></ProfMisconductEverDate>\n")
    w("          </ProfMisconductsEver>\n")
    w("          <MisconductInvestsNow>\n")
    w(f"            <MisconductInvestNow>{_x(c['MisconductInvestNow'])}</MisconductInvestNow>\n")
    w("            <MisconductInvestNowDetails></MisconductInvestNowDetails>\n")
    w("            <MisconductInvestNowDate></MisconductInvestNowDate>\n")
    w("          </MisconductInvestsNow>\n")
    w("          <AnyCurrentProceedings>\n")
    w(f"            <AnyCurProceeding>{_x(c['AnyCurProceeding'])}</AnyCurProceeding>\n")
    w("            <AnyCurProceedingDetails></AnyCurProceedingDetails>\n")
    w("            <AnyCurProceedingDate></AnyCurProceedingDate>\n")
    w("          </AnyCurrentProceedings>\n")
    w("        </NursingCondition>\n")
    w("      </NursingConditions>\n")
    w("      <NursingEmployers></NursingEmployers>\n")
    # language testing (empty template block)
    w("      <LanguageTesting>\n")
    w("        <LanguageTest>\n")
    for t in ("Provider", "Listening", "Reading", "Writing", "Speaking",
              "Vocabulary", "Overall", "LanguageTestDate", "TestReportNumber", "ReceivedDate"):
        w(f"          <{t}></{t}>\n")
    w("          <Documents>\n")
    w("            <Document>\n")
    w("              <DocumentType></DocumentType>\n")
    w("              <DocumentNumber></DocumentNumber>\n")
    w("              <IssuedFrom iso-code=\"\"></IssuedFrom>\n")
    w("              <DateIssued></DateIssued>\n")
    w("              <ReceivedDate></ReceivedDate>\n")
    w("              <ExpirationDate></ExpirationDate>\n")
    w("              <Files>\n")
    w("                <File>\n")
    for t in ("Name", "Path", "Extension", "Size", "LastModifiedDate"):
        w(f"                  <{t}></{t}>\n")
    w("                </File>\n")
    w("              </Files>\n")
    w("            </Document>\n")
    w("          </Documents>\n")
    w("        </LanguageTest>\n")
    w("      </LanguageTesting>\n")
    # identity documents (each Document wrapped in its own <Documents>)
    w("      <IdentityDocuments>\n")
    for d in docs.get("identity", []):
        w("        <Documents>\n")
        w(_doc_xml(d, "          ", with_expiry=True))
        w("        </Documents>\n")
    w("      </IdentityDocuments>\n")
    # other documents (report + other)
    w("      <OtherDocuments>\n")
    for d in docs.get("report", []) + docs.get("other", []):
        w("        <Documents>\n")
        w(_doc_xml(d, "          ", with_expiry=bool(d.get("ExpirationDate") or d.get("_with_expiry", True))))
        w("        </Documents>\n")
    w("      </OtherDocuments>\n")
    w("    </Application>\n")



# =========================================================
# CSV SERIALIZER  (flattened, quoted + trailing space, blanks as '-')
# =========================================================

def _csv_cell(v: Any) -> str:
    s = "" if v is None else str(v).strip()
    return f"{s if s else '-'} "   # value + trailing space; '-' for blanks (matches reference)


def _csv_header_values(model: Dict[str, Any],
                       docs: Dict[str, List[Dict[str, str]]]) -> Tuple[List[str], List[str]]:
    p, r, c = model["Personal"], model["Registration"], model["Conditions"]
    h, ma, aa = p["HomeAddress"], p["MailingAddress"], r["AuthorityAddress"]

    headers: List[str] = []
    values: List[str] = []

    def add(tag: str, val: Any):
        headers.append(f"<{tag}>")
        values.append(_csv_cell(val))

    add("ID", model["ID"]); add("ApplicationID", model["ApplicationID"])
    add("NursingGroup", model["NursingGroup"]); add("LanguageOfApplication", model["LanguageOfApplication"])
    add("RBName", model["RBName"]); add("RBAbbrev", model["RBAbbrev"])
    add("RBProvince", model["RBProvince"])
    add("FirstName", p["FirstName"]); add("MiddleName", p["MiddleName"]); add("LastName", p["LastName"])
    add("OtherName", p["OtherName"]); add("DateOfBirth", p["DateOfBirth"]); add("Gender", p["Gender"])
    add("SpokenLanguage", p["SpokenLanguage"]); add("WrittenLanguage", p["WrittenLanguage"])
    add("FirstLanguage", p["FirstLanguage"]); add("InitialNursingEducationCountry", p["InitialNursingEducationCountry"])
    for t in ("StreetAddress1", "StreetAddress2", "AptNo", "City", "State", "Country", "PostalCode"):
        add(t, h[t])
    add("MailingAddress", "")  # the reference repeats address tags under a MailingAddress marker
    for t in ("StreetAddress1", "StreetAddress2", "AptNo", "City", "State", "Country", "PostalCode"):
        add(t, ma[t])
    add("PrimaryPhone", p["PrimaryPhone"]); add("AlternatePhone", p["AlternatePhone"]); add("Email", p["Email"])
    add("RegisteredInJurisdictionOfInitialProgram", r["RegisteredInJurisdictionOfInitialProgram"])
    add("DiplomaRightToPractice", r["DiplomaRightToPractice"]); add("DiplomaNumber", r["DiplomaNumber"])
    add("LicenseTitle", r["LicenseTitle"]); add("LicenseNumber", r["LicenseNumber"])
    add("LicenseStatus", r["LicenseStatus"]); add("LicenseIssueDate", r["LicenseIssueDate"])
    add("LicenseExpiredDate", r["LicenseExpiredDate"]); add("NameUsedWhenRegIssued", r["NameUsedWhenRegIssued"])
    add("LicenseAuthorityName", r["LicenseAuthorityName"])
    for t in ("StreetAddress1", "StreetAddress2", "POBox", "City", "State", "Country", "PostalCode"):
        add(t, aa[t])
    add("RegMethod", r["RegMethod"]); add("EligibleInJurisdiction", r["EligibleInJurisdiction"])
    add("LifetimeRegistration", r["LifetimeRegistration"]); add("RegOtherHealthProf", r["RegOtherHealthProf"])
    add("OtherRegulatedProf", r["OtherRegulatedProf"]); add("LicenseEverSuspended", r["LicenseEverSuspended"])
    add("LicenseSuspendDetails", r["LicenseSuspendDetails"]); add("DateOfficialSigned", r["DateOfficialSigned"])
    add("ReceivedDate", r["ReceivedDate"])

    def add_doc(d: Dict[str, str], with_expiry: bool = True):
        add("DocumentType", d["DocumentType"]); add("DocumentNumber", d["DocumentNumber"])
        add("IssuedFrom", d["IssuedFrom"]); add("DateIssued", d["DateIssued"])
        add("ReceivedDate", d["ReceivedDate"])
        if with_expiry:
            add("ExpirationDate", d["ExpirationDate"])
        add("Name", d["Name"]); add("Path", d["Path"]); add("Extension", d["Extension"])
        add("Size", d["Size"]); add("LastModifiedDate", d["LastModifiedDate"])

    for d in docs.get("registration", []):
        add_doc(d, with_expiry=True)

    add("CurRegInAnotherCanadianProv", c["CurRegInAnotherCanadianProv"])
    add("CurCanadianRegIssueDate", c["CurCanadianRegIssueDate"]); add("CurCanadianRegProvName", c["CurCanadianRegProvName"])
    add("NumberAppliedOtherProvinces", c["NumberAppliedOtherProvinces"])
    add("ProvinceApplied", c["ProvinceApplied"]); add("DateApplied", c["DateApplied"])
    add("RefCompetencyAssessment", c["RefCompetencyAssessment"]); add("CompleteCompetencyAssess", c["CompleteCompetencyAssess"])
    add("CurLimitations", c["CurLimitation"]); add("CurLimitationDetails", ""); add("CurLimitationDate", "")
    add("LicenseRevoked", c["LicenseRevoked"]); add("LicenseRevokedDetails", ""); add("LicenseRevokedDate", "")
    add("CriminalOffense", c["CriminalOffense"]); add("CriminalOffenseDetails", ""); add("CriminalOffenseDate", "")
    add("AddictionImpairing", c["AddictionImpairing"]); add("AddictionImpairingDetails", ""); add("AddictionImpairingDate", "")
    add("ProfMisconductEver", c["ProfMisconductEver"]); add("ProfMisconductEverDetails", ""); add("ProfMisconductEverDate", "")
    add("MisconductInvestNow", c["MisconductInvestNow"]); add("MisconductInvestNowDetails", ""); add("MisconductInvestNowDate", "")
    add("AnyCurProceeding", c["AnyCurProceeding"]); add("AnyCurProceedingDetails", ""); add("AnyCurProceedingDate", "")

    # language testing (empty)
    for t in ("Provider", "Listening", "Reading", "Writing", "Speaking",
              "Vocabulary", "Overall", "LanguageTestDate", "TestReportNumber", "ReceivedDate"):
        add(t, "")
    # identity + report + other document blocks
    for d in docs.get("identity", []):
        add_doc(d, with_expiry=True)
    for d in docs.get("report", []):
        add_doc(d, with_expiry=True)
    for d in docs.get("other", []):
        add_doc(d, with_expiry=False)

    return headers, values


def build_csv(records: List[Tuple[Dict[str, Any], Dict[str, List[Dict[str, str]]]]]) -> str:
    """
    Build the interchange CSV with one data row per record (a record is a
    (model, docs) pair). Multiple records occur when a regulator assessed the
    applicant under more than one designation -- e.g. an RN row and an LPN row.
    The header is taken from the first record; all records share the same
    column structure (identical shared docs + one report each).
    """
    headers: Optional[List[str]] = None
    rows: List[List[str]] = []
    for model, docs in records:
        h, v = _csv_header_values(model, docs)
        if headers is None:
            headers = h
        rows.append(v)

    buf = StringIO()
    writer = csv.writer(buf, quoting=csv.QUOTE_ALL, lineterminator="\n")
    writer.writerow(headers or [])
    for v in rows:
        writer.writerow(v)
    return buf.getvalue()


# =========================================================
# CSV OUTPUT CHECK
# =========================================================
# Re-reads the CSV text exactly as it will be written into the zip and checks
# each row against the record it was built from. Like check_dob, it never puts
# the DOB or any other applicant value into warnings or debug output.

# Applicants younger than this (or older than the maximum) are flagged. A
# wrong birth year usually lands far outside this range, e.g. a DOB that
# makes the applicant a child.
MIN_APPLICANT_AGE = 18
MAX_APPLICANT_AGE = 90

# When True, a regulator whose CSV fails any check is NOT pushed to the SFTP
# (no zip upload, no index rows). The SharePoint export still happens so the
# package can be inspected. Set to False to warn only.
CSV_CHECK_BLOCKS_DELIVERY = True


def _csv_value(cell: str) -> str:
    """Undo _csv_cell: strip the trailing space and treat '-' as blank."""
    v = (cell or "").strip()
    return "" if v == "-" else v


def _age_on(dob: datetime, today: datetime) -> int:
    return today.year - dob.year - ((today.month, today.day) < (dob.month, dob.day))


def verify_csv_output(csv_text: str,
                      records: List[Tuple[Dict[str, Any], Dict[str, List[Dict[str, str]]]]],
                      app_id: str) -> Dict[str, Any]:
    """
    Check the built CSV before it is zipped and delivered:
      - one data row per record (designation)
      - ApplicationID and NursingGroup match the record
      - DateOfBirth is present, a real dd-mm-yyyy date, not in the future,
        and gives a plausible applicant age
      - DateOfBirth matches the model and the NNAS index row for that record

    Returns {"ok", "rows_expected", "rows_found", "row_checks", "warnings"}.
    No applicant values are included in the result.
    """
    warns: List[str] = []
    row_checks: List[Dict[str, Any]] = []

    try:
        parsed = list(csv.reader(StringIO(csv_text)))
    except Exception as e:
        return {"ok": False, "rows_expected": len(records), "rows_found": 0,
                "row_checks": [], "warnings": [f"CSV could not be re-read: {type(e).__name__}"]}

    if not parsed:
        return {"ok": False, "rows_expected": len(records), "rows_found": 0,
                "row_checks": [], "warnings": ["CSV is empty."]}

    header = [h.strip() for h in parsed[0]]
    rows = [r for r in parsed[1:] if any(c.strip() for c in r)]

    if len(rows) != len(records):
        warns.append(f"CSV has {len(rows)} data row(s) but {len(records)} were expected "
                     "(one per designation).")

    def col(tag: str) -> Optional[int]:
        # First occurrence; these tags appear once in the reference layout.
        try:
            return header.index(f"<{tag}>")
        except ValueError:
            return None

    i_app, i_grp, i_dob = col("ApplicationID"), col("NursingGroup"), col("DateOfBirth")
    for tag, idx in (("ApplicationID", i_app), ("NursingGroup", i_grp), ("DateOfBirth", i_dob)):
        if idx is None:
            warns.append(f"CSV header has no <{tag}> column.")

    today = datetime.utcnow()

    for n, (row, (model, _docs)) in enumerate(zip(rows, records), start=1):
        label = f"row {n} ({model.get('NursingGroup') or 'no designation'})"
        chk: Dict[str, Any] = {"row": n, "designation": model.get("NursingGroup", "")}
        problems: List[str] = []

        if len(row) != len(header):
            problems.append(f"has {len(row)} columns, header has {len(header)}")

        if i_app is not None and i_app < len(row):
            if _csv_value(row[i_app]) != str(app_id):
                problems.append("ApplicationID does not match this application")

        if i_grp is not None and i_grp < len(row):
            if _csv_value(row[i_grp]) != (model.get("NursingGroup") or ""):
                problems.append("NursingGroup does not match the report designation")

        dob_ok = False
        if i_dob is not None and i_dob < len(row):
            dob_txt = _csv_value(row[i_dob])
            if not dob_txt:
                problems.append("DateOfBirth is blank")
            else:
                try:
                    dob = datetime.strptime(dob_txt, "%d-%m-%Y")
                except ValueError:
                    dob = None
                    problems.append("DateOfBirth is not a valid dd-mm-yyyy date")
                if dob is not None:
                    if dob > today:
                        problems.append("DateOfBirth is in the future")
                    else:
                        age = _age_on(dob, today)
                        if age < MIN_APPLICANT_AGE:
                            problems.append(f"DateOfBirth gives an age under {MIN_APPLICANT_AGE}; "
                                            "check the birth year")
                        elif age > MAX_APPLICANT_AGE:
                            problems.append(f"DateOfBirth gives an age over {MAX_APPLICANT_AGE}; "
                                            "check the birth year")
                        else:
                            dob_ok = True

                if dob_txt != model["Personal"].get("DateOfBirth", ""):
                    problems.append("DateOfBirth in the CSV differs from the applicant model")
                    dob_ok = False
                if dob_txt != build_index_row(model)[4]:
                    problems.append("DateOfBirth in the CSV differs from the NNAS index row")
                    dob_ok = False

        chk["dob_ok"] = dob_ok
        chk["problems"] = problems
        row_checks.append(chk)
        warns.extend(f"CSV {label}: {p}." for p in problems)

    return {
        "ok": not warns,
        "rows_expected": len(records),
        "rows_found": len(rows),
        "row_checks": row_checks,
        "warnings": warns,
    }


# =========================================================
# XML OUTPUT CHECK
# =========================================================
# Parses the XML text exactly as it will be written into the zip and checks
# each <Application>. Only tag names and counts are reported; no applicant
# values are placed in warnings or debug output.

# Fields a tester usually expects to be filled. Reported by tag name only.
XML_PRESENCE_TAGS = (
    "FirstName", "LastName", "DateOfBirth", "Gender", "Email", "PrimaryPhone",
    "InitialNursingEducationCountry", "LicenseNumber", "LicenseStatus",
    "LicenseIssueDate", "LicenseAuthorityName",
)

# Form prompt text showing up inside a value means a field mapping slipped,
# e.g. "First language: ..." written into DateOfBirth.
_LABEL_LEAK_PATTERNS = [p.lower() for p, _ in PROMPTS] + [
    "first language:", "where applicant received",
]

# When True, a regulator whose XML fails the check is held back from the SFTP,
# the same way a failed CSV check is. Set to False to warn only.
XML_CHECK_BLOCKS_DELIVERY = True


def verify_xml_output(xml_text: str,
                      records: List[Tuple[Dict[str, Any], Dict[str, List[Dict[str, str]]]]]
                      ) -> Dict[str, Any]:
    """
    Check the built XML before it is zipped and delivered:
      - well-formed
      - one <Application> per record (designation)
      - DateOfBirth present and dd-mm-yyyy
      - at least one identity document listed
      - no form label text inside any value

    Also reports, for information only, key fields that are blank and fields
    that have a value but an empty iso-code attribute.

    Returns {"ok", "applications_expected", "applications_found",
    "applications", "warnings"}. No applicant values are included.
    """
    warns: List[str] = []
    apps_out: List[Dict[str, Any]] = []

    try:
        root = ET.fromstring(xml_text.encode("utf-8"))
    except ET.ParseError as e:
        return {"ok": False, "applications_expected": len(records),
                "applications_found": 0, "applications": [],
                "warnings": [f"XML is not well formed (line {e.position[0]})."]}

    apps = root.findall("./Applications/Application")
    if len(apps) != len(records):
        warns.append(f"XML has {len(apps)} <Application> element(s) but "
                     f"{len(records)} were expected (one per designation).")

    for n, app in enumerate(apps, start=1):
        grp = (app.findtext("NursingGroup") or "").strip()
        label = f"application {n} ({grp or 'no designation'})"

        blank = [t for t in XML_PRESENCE_TAGS
                 if not (app.findtext(f".//{t}") or "").strip()]

        leaks = sorted({el.tag for el in app.iter()
                        if (el.text or "").strip()
                        and any(p in el.text.lower() for p in _LABEL_LEAK_PATTERNS)})

        empty_iso = sorted({el.tag for el in app.iter()
                            if "iso-code" in el.attrib
                            and (el.text or "").strip()
                            and not el.attrib["iso-code"]})

        dob = (app.findtext("./PersonalInfo/DateOfBirth") or "").strip()
        dob_format_ok = bool(_DDMMYYYY_RE.match(dob))

        doc_counts = {
            "registration": len(app.findall("./NursingRegistrations//Documents/Document")),
            "identity": len(app.findall("./IdentityDocuments/Documents/Document")),
            "other_and_report": len(app.findall("./OtherDocuments/Documents/Document")),
        }

        if leaks:
            warns.append(f"XML {label}: form label text found in "
                         f"{', '.join(leaks)}; check the field mapping.")
        if not dob_format_ok:
            warns.append(f"XML {label}: DateOfBirth is blank or not dd-mm-yyyy.")
        if doc_counts["identity"] == 0:
            warns.append(f"XML {label}: no identity documents listed.")

        apps_out.append({
            "application": n,
            "designation": grp,
            "blank_fields": blank,                 # informational
            "label_leak_fields": leaks,
            "empty_iso_code_fields": empty_iso,    # informational
            "dob_format_ok": dob_format_ok,
            "doc_counts": doc_counts,
        })

    return {
        "ok": not warns,
        "applications_expected": len(records),
        "applications_found": len(apps),
        "applications": apps_out,
        "warnings": warns,
    }


# =========================================================
# SFTP DELIVERY  +  PER-RB NNAS INDEX
# =========================================================

def build_index_row(model: Dict[str, Any]) -> List[str]:
    """
    One headerless NNAS index row, matching SRNA_NNAS_index sample:
      LASTNAME, FIRSTNAME, id, applicationID, DOB, RN|LPN, x, N, x
    (columns 7 and 9 are placeholders 'x'; column 8 is 'N'.)
    """
    p = model["Personal"]
    return [
        p["LastName"],
        p["FirstName"],
        model["ID"],
        model["ApplicationID"],
        p["DateOfBirth"],            # already dd-mm-yyyy
        model["NursingGroup"],       # RN | LPN
        "x",
        "N",
        "x",
    ]


def _index_line(row: List[str]) -> str:
    # plain comma join (sample has no quoting); strip stray commas in fields
    safe = [str(c).replace(",", " ").strip() for c in row]
    return ",".join(safe) + "\n"


def sftp_connect() -> Tuple[paramiko.SFTPClient, paramiko.Transport]:
    username = wmill.get_variable(SFTP_USERNAME_VAR)
    password = wmill.get_variable(SFTP_PASSWORD_VAR)

    # Surface credential-resolution problems clearly (a blank/None secret is the
    # most common cause of "Authentication failed").
    if not username or not password:
        missing = []
        if not username:
            missing.append(SFTP_USERNAME_VAR)
        if not password:
            missing.append(SFTP_PASSWORD_VAR)
        raise RuntimeError(
            "SFTP credentials not resolved from Windmill variables: "
            + ", ".join(missing)
            + ". Check the variable paths exist and the script has access."
        )

    transport = paramiko.Transport((SFTP_HOST, SFTP_PORT))
    transport.connect(hostkey=None)  # open transport without auth first

    # 1) plain password auth
    try:
        transport.auth_password(username=username, password=password)
    except paramiko.AuthenticationException:
        # 2) some SFTP servers (incl. many SFTP-only gateways) require
        #    keyboard-interactive even when the prompt is just the password.
        try:
            def _handler(title, instructions, prompt_list):
                return [password for _ in prompt_list]
            transport.auth_interactive(username=username, handler=_handler)
        except Exception:
            transport.close()
            raise RuntimeError(
                "SFTP authentication failed for user "
                f"'{username}' (password and keyboard-interactive both rejected). "
                "Verify NNAS_SFTP_USERNAME / NNAS_SFTP_PASSWORD values."
            )

    if not transport.is_authenticated():
        transport.close()
        raise RuntimeError(f"SFTP authentication failed for user '{username}'.")

    sftp = paramiko.SFTPClient.from_transport(transport)
    return sftp, transport


def _sftp_isdir(sftp: paramiko.SFTPClient, path: str) -> bool:
    try:
        from stat import S_ISDIR
        return S_ISDIR(sftp.stat(path).st_mode)
    except IOError:
        return False


def sftp_regulator_dir(sftp: paramiko.SFTPClient, abbrev: str) -> str:
    """
    Resolve the regulator root folder under /REGULATORY/. Folders are named by
    the UPPER-CASE abbreviation (per the screenshot: BCCNM, CRNA, ...). Falls
    back to a case-insensitive match against existing folders; does NOT create
    new regulator folders (they are provisioned on the SFTP side). The daily
    NNAS index lives here, at the regulator root.
    """
    target = f"{SFTP_REGULATORY_ROOT}/{abbrev.upper()}"
    if _sftp_isdir(sftp, target):
        return target
    try:
        existing = sftp.listdir(SFTP_REGULATORY_ROOT)
    except IOError:
        return target  # let the upload surface the error
    for name in existing:
        if name.lower() == abbrev.lower():
            return f"{SFTP_REGULATORY_ROOT}/{name}"
    return target


def _sftp_mkdir(sftp: paramiko.SFTPClient, path: str) -> None:
    """Create a remote directory if it doesn't exist (single level)."""
    if not _sftp_isdir(sftp, path):
        try:
            sftp.mkdir(path)
        except IOError:
            # may already exist (race) or parent missing; re-check
            if not _sftp_isdir(sftp, path):
                raise


def sftp_app_dir(sftp: paramiko.SFTPClient, regulator_dir: str, app_id: str) -> str:
    """
    Per-package subfolder under the regulator root, created if missing:
        /REGULATORY/<RB>/<ApplicationID>/
    The zip is delivered here; the index stays at <regulator_dir>.
    """
    app_dir = f"{regulator_dir}/{app_id}"
    _sftp_mkdir(sftp, app_dir)
    return app_dir


def sftp_upload_zip(sftp: paramiko.SFTPClient, local_zip: str,
                    remote_dir: str, filename: str) -> str:
    remote_path = f"{remote_dir}/{filename}"
    sftp.put(local_zip, remote_path)
    return remote_path


def sftp_append_index(sftp: paramiko.SFTPClient, remote_dir: str,
                      filename: str, row: List[str]) -> str:
    """
    Append one row to the regulator's daily NNAS index file at <remote_dir>,
    preserving any existing rows for that day. Read-modify-write (SFTP has no
    true cross-server append). A new day -> a new filename -> a fresh file.
    """
    remote_path = f"{remote_dir}/{filename}"
    existing = ""
    try:
        with sftp.open(remote_path, "r") as fh:
            existing = fh.read().decode("utf-8", errors="replace")
    except IOError:
        existing = ""
    if existing and not existing.endswith("\n"):
        existing += "\n"
    new_content = existing + _index_line(row)
    with sftp.open(remote_path, "w") as fh:
        fh.write(new_content)
    return remote_path


# =========================================================
# APPLICANT COPY  (report with the Assessment Notes block removed)
# =========================================================

def strip_assessment_notes(in_pdf: str, out_pdf: str) -> bool:
    """
    Produce an applicant copy of a report PDF with the standalone "Assessment
    Notes" block (its grey header bar + body box, incl. any RB-specific notes)
    fully removed from page 1. The "Public Assessment Notes" block is kept.

    Removal is a true redaction (content under the region is deleted, not just
    covered). Returns True if the block was found and removed.
    """
    doc = fitz.open(in_pdf)
    try:
        pg = doc[0]
        pub = pg.search_for("Public Assessment Notes")
        all_an = pg.search_for("Assessment Notes")
        if not all_an:
            doc.save(out_pdf, garbage=4, deflate=True)
            return False
        # The standalone header is an "Assessment Notes" match that is NOT the
        # tail of a "Public Assessment Notes" match (same y0). Pick the lowest.
        pub_ys = {round(r.y0) for r in pub}
        header = None
        for r in sorted(all_an, key=lambda r: r.y0, reverse=True):
            if round(r.y0) not in pub_ys:
                header = r
                break
        if header is None:
            doc.save(out_pdf, garbage=4, deflate=True)
            return False
        # Union of drawing boxes that start at/above the header and extend below
        # it (the grey header bar + the body box).
        boxes = [dr["rect"] for dr in pg.get_drawings()
                 if dr["rect"].width > 200 and dr["rect"].height > 20]
        region = None
        for b in boxes:
            if b.y0 <= header.y0 + 2 and b.y1 >= header.y1:
                region = b if region is None else (region | b)
        if region is None:
            region = fitz.Rect(30, header.y0 - 4,
                               pg.rect.width - 30, pg.rect.height - 60)
        region = fitz.Rect(region.x0 - 2, region.y0 - 2,
                           region.x1 + 2, region.y1 + 2)
        pg.add_redact_annot(region, fill=(1, 1, 1))
        pg.apply_redactions()
        doc.save(out_pdf, garbage=4, deflate=True)
        return True
    finally:
        doc.close()


def applicant_copy_name(report_name: str) -> str:
    """report '..._latest.pdf' -> applicant copy '..._applicant.pdf'."""
    base, ext = os.path.splitext(report_name)
    if base.lower().endswith("_latest"):
        base = base[:-len("_latest")]
    return f"{base}{APPLICANT_COPY_SUFFIX}{ext}"



# =========================================================
# ENTRYPOINT
# =========================================================

def main(reference_id: Optional[str] = None,
         additional_content: Optional[str] = None,
         user_id: Any = None,
         dry_run: bool = False,
         **kwargs: Any):
    """
    Accepts the SMApply trigger JSON directly. Only three fields drive the run;
    every other top-level webhook key (title, id, award, status, user,
    date_created, last_edited, round, token, etc.) is absorbed by **kwargs so
    Windmill binds the payload cleanly WITHOUT cluttering the run form.

    Field mapping (per NNAS):
        reference_id (or title) -> ApplicationID (e.g. "NNAS-1563663808")
                                   and the SharePoint folder <provider>/<id>
        user_id                 -> NNAS <ID>     (e.g. 131243235)
        additional_content      -> the en-masse JSON blob

    dry_run: when True, everything runs as normal (zips are built, the
    applicant copies are produced, and all artifacts are uploaded to the
    SharePoint export folder), but the final push to the regulator SFTP is
    SKIPPED. No SFTP connection is opened and no NNAS index rows are written.
    The CSV and XML output checks still run, and the loose CSV/XML files are
    still uploaded to SharePoint, so a dry run is a full test of the package.

    Backward-compat: appid / application_title / nnas_id are still honored if
    present in kwargs (older flows), but no longer shown as separate fields.
    """
    # pull recognized fallbacks out of the catch-all
    title             = kwargs.get("title")
    application_title = kwargs.get("application_title")
    appid             = kwargs.get("appid")
    nnas_id           = kwargs.get("nnas_id")

    sp = ctx()
    payload, ac_dbg = parse_additional_content(additional_content)

    warnings: List[str] = []
    if additional_content and not ac_dbg.get("parsed"):
        warnings.append(
            "additional_content could not be parsed even after JSON repair -- "
            "the regulator list falls back to the application-summary PDF and "
            "may be incomplete. See debug.additional_content.errors."
        )
    elif ac_dbg.get("repaired"):
        warnings.append(
            f"additional_content required JSON repair (method: {ac_dbg.get('method')}). "
            "Most likely cause: unescaped double quotes inside an assessment note."
        )

    # ---- ApplicationID + SharePoint folder: reference_id, then title, then appid ----
    app_identifier = first_nonempty(
        reference_id, title, application_title,
        safe_str(payload.get("reference_id")),
        (str(appid) if appid is not None else ""),
    )
    if not app_identifier:
        return {"status": "error",
                "message": "No reference_id/title/ApplicationID to locate the SharePoint folder.",
                "warnings": warnings,
                "debug": {"additional_content": ac_dbg}}

    # ApplicationID and folder name are the same value (e.g. NNAS-1563663808).
    folder_app_id = app_identifier
    app_id_num    = app_identifier

    # ---- NNAS <ID>: user_id, then nnas_id alias, then additional_content ----
    doc_id = first_nonempty(
        (str(user_id) if user_id is not None else ""),
        (str(nnas_id) if nnas_id is not None else ""),
        safe_str(payload.get("nnas_id")),
    )

    # ---- provider selects the SharePoint base path ----
    provider = safe_str((payload.get("application_info") or {}).get("provider"))
    base_template, provider_norm, provider_err = base_template_for(provider)
    if provider_err:
        return {"status": "error", "message": provider_err,
                "warnings": warnings,
                "debug": {"provider": provider, "additional_content": ac_dbg}}

    # Built through _display_server_relative_path so the result has exactly the
    # same shape as the "url" values list_folder_files returns. The
    # "Additional documents" mirroring below slices file URLs by the length of
    # add_docs_rel, so a stray double slash here would corrupt every mirrored
    # file name.
    base_path    = base_template.format(app_id=folder_app_id)
    app_base_rel = _display_server_relative_path(base_path)
    intl_rel     = f"{app_base_rel}/{INTL_CREDENTIALS_SUBFOLDER}"
    reports_rel  = f"{app_base_rel}/{CREDENTIAL_REPORTS_SUBFOLDER}"
    add_docs_rel = f"{app_base_rel}/{ADDITIONAL_DOCS_SUBFOLDER}"

    # Fail loudly if the application folder doesn't exist for this provider,
    # rather than silently producing an empty package (e.g. a CES application
    # whose folder isn't under ces/).
    if not folder_exists(sp, app_base_rel):
        return {"status": "error",
                "message": (f"Application folder not found for provider "
                            f"'{provider_norm.upper()}': {app_base_rel}. "
                            f"Verify the application exists under "
                            f"{DOCUMENT_LIBRARY_NAME}/{provider_norm}/."),
                "warnings": warnings,
                "debug": {"provider": provider_norm, "base_path": app_base_rel,
                          "additional_content": ac_dbg}}

    debug: Dict[str, Any] = {"folder_app_id": folder_app_id, "app_id_num": app_id_num,
                             "doc_id": doc_id,
                             "provider": provider_norm.upper(),
                             "base_path": base_path,
                             "additional_content": ac_dbg}

    workdir = os.path.join("tmp", f"pkg_{folder_app_id}_{datetime.utcnow():%Y%m%d_%H%M%S}")
    dl_dir  = os.path.join(workdir, "_dl")
    os.makedirs(dl_dir, exist_ok=True)

    base_files     = list_folder_files(sp, app_base_rel, recursive=False)
    intl_files     = list_folder_files(sp, intl_rel, recursive=True)
    report_files   = list_folder_files(sp, reports_rel, recursive=True)
    add_docs_files = list_folder_files(sp, add_docs_rel, recursive=True)
    debug["base_files"]            = [f["name"] for f in base_files]
    debug["intl_files"]            = [f["name"] for f in intl_files]
    debug["report_files"]          = [f["name"] for f in report_files]
    debug["additional_docs_files"] = [f["name"] for f in add_docs_files]

    # ---- newest applicant application-summary form in the application root ----
    # '<ReferenceID>_<YYYYMMDD>_<HHMMSS>', newest timestamp only. This single
    # file is BOTH parsed for applicant-level fields (below) and packaged into
    # images/other/ of every regulator package.
    applicant_form_entry, applicant_form_ts, applicant_form_candidates = \
        pick_latest_applicant_form(base_files, folder_app_id)

    # ---- choose + parse the application-summary PDF ----
    # Primary: the newest stamped applicant form (same file that is packaged).
    # Fallback: the old loose filter (first PDF that is not an ID doc / report),
    # only when no stamped summary PDF exists. The fallback is warned because
    # listing order is not chronological.
    form: Dict[str, str] = {}
    form_pdf: Optional[Dict[str, Any]] = None
    form_pdf_source: Optional[str] = None
    if applicant_form_entry and applicant_form_entry["name"].lower().endswith(".pdf"):
        form_pdf = applicant_form_entry
        form_pdf_source = "latest_applicant_form"
    else:
        form_pdf = next((f for f in base_files
                         if f["name"].lower().endswith(".pdf")
                         and not ID_DOC_RE.search(os.path.splitext(f["name"])[0])
                         and meta_get(f["props"], "FileType").lower() != "report"
                         and "_credentialform_" not in f["name"].lower()), None)
        if form_pdf:
            form_pdf_source = "fallback_first_pdf"
            warnings.append(
                f"No '{folder_app_id}_<YYYYMMDD>_<HHMMSS>.pdf' application summary "
                f"found in the application root; parsed '{form_pdf['name']}' instead. "
                "Applicant fields (including DOB) may come from an unexpected file."
            )
        else:
            warnings.append(
                "No application-summary PDF found in the application root. "
                "Applicant fields fall back to additional_content only."
            )
    debug["form_pdf_source"] = form_pdf_source

    if form_pdf:
        try:
            local = os.path.join(dl_dir, form_pdf["name"])
            download_file(sp, form_pdf["url"], local)
            form = parse_application_pdf(local)
            debug["form_pdf"] = form_pdf["name"]
            debug["form_fields_parsed"] = len(form)
            debug["form_rb_selected"] = form.get("rb_selected", "")
        except Exception as e:
            debug["form_parse_error"] = str(e)
            warnings.append(
                f"Application summary '{form_pdf['name']}' could not be parsed; "
                "applicant fields fall back to additional_content only."
            )
    if not doc_id:
        doc_id = safe_str(payload.get("nnas_id")) or folder_app_id

    # ---- date of birth cross-check (PDF vs additional_content) ----
    dob_check = check_dob(form, payload)
    warnings.extend(dob_check.pop("warnings"))
    debug["dob"] = dob_check

    # ---- collect the SHARED documents (identity / registration / other) ----
    # These are identical across every regulator's package; the per-regulator
    # report is added separately below.
    shared: List[Dict[str, Any]] = []
    skipped: List[str] = []

    # ---- applicant application-summary form: prepare the packaged copy ----
    # Copied verbatim into images/other/ of every regulator package, renamed
    # '<ReferenceID>-Applicant Form - <YYYYMMDD_HHMMSS>'. Reuses the local file
    # if the parse step above already downloaded it.
    applicant_form_local: Optional[str] = None
    applicant_form_packaged: Optional[str] = None
    if applicant_form_entry:
        try:
            local_af = os.path.join(dl_dir, applicant_form_entry["name"])
            if not os.path.exists(local_af):
                download_file(sp, applicant_form_entry["url"], local_af)
            applicant_form_local = local_af
            applicant_form_packaged = applicant_form_name(
                folder_app_id, applicant_form_ts,
                os.path.splitext(applicant_form_entry["name"])[1] or ".pdf")
        except Exception as e:
            applicant_form_local = None
            skipped.append(f"{applicant_form_entry['name']} (applicant form download failed: {e})")
    debug["applicant_form"] = {
        "candidates": applicant_form_candidates,
        "selected": applicant_form_entry["name"] if applicant_form_entry else None,
        "packaged_as": applicant_form_packaged,
    }

    translation_files: List[str] = []

    def consider(entry: Dict[str, Any], from_base: bool):
        name, props = entry["name"], entry.get("props", {})
        if not name or name.lower().endswith((".zip", ".csv", ".xml")):
            skipped.append(name); return
        if form_pdf and from_base and name == form_pdf["name"]:
            skipped.append(name + " (application form)"); return

        # Translation documents in International Credentials: always
        # images/other/ with doctype TRANSLATION, regardless of SharePoint tags.
        if not from_base and is_translation_file(name):
            translation_files.append(name)
            shared.append({**entry, "_category": "other",
                           "_doctype": TRANSLATION_DOCTYPE, "_forced": None,
                           "_translation": True})
            return

        category, doctype, forced = categorize(props, name, folder_app_id, form)
        ftype = meta_get(props, "FileType").lower()
        if from_base:
            is_id = bool(ID_DOC_RE.search(os.path.splitext(name)[0])) and \
                    os.path.splitext(name)[0].lower().startswith(folder_app_id.lower())
            if ftype == "report" or category == "report":
                skipped.append(name + " (report handled from credential reports folder)")
                return
            if not is_id:
                skipped.append(name); return
        if category == "report":     # never let intl files masquerade as reports
            category = "other"
        shared.append({**entry, "_category": category, "_doctype": doctype, "_forced": forced})

    for e in base_files:
        consider(e, True)
    for e in intl_files:
        consider(e, False)
    debug["translation_files"] = translation_files

    # ---- map the credential-report files to regulators (latest only) ----
    # Only "*_latest.pdf" files are considered; each is matched to its regulator
    # by the abbreviation token in the filename. A regulator may have MORE THAN
    # ONE latest report (one per designation, e.g. RN and LPN), so collect a
    # dict per abbrev keyed by designation; a later duplicate for the same
    # (abbrev, designation) wins.
    reports_by_abbrev: Dict[str, Dict[str, Dict[str, Any]]] = {}
    for rf in report_files:
        nm = rf["name"]
        if not nm.lower().endswith(REPORT_LATEST_SUFFIX):
            if nm.lower().endswith((".pdf", ".docx")):
                # name the actual suffix under test: a .docx or an
                # _applicant.pdf is correctly excluded here even though its
                # stem may still contain "_latest".
                skipped.append(f"{nm} (not {REPORT_LATEST_SUFFIX})")
            continue
        ab = regulator_of_report(nm)
        if not ab:
            skipped.append(nm + " (no regulator abbrev matched)")
            continue
        desig = designation_of_report(nm)        # RN / LPN / RPN / ""
        entry = {**rf, "_category": "report", "_designation": desig,
                 "_doctype": ((f"{desig} " if desig else "")
                              + (meta_get(rf["props"], "Title") or "NNAS REPORT")).upper()}
        reports_by_abbrev.setdefault(ab, {})[desig] = entry  # keyed by designation
    debug["reports_by_abbrev"] = {
        ab: [e["name"] for e in by_des.values()] for ab, by_des in reports_by_abbrev.items()
    }

    # ---- which regulators did the applicant apply to? ----
    # An application may carry RN and LPN regulators at the same time, so take
    # the UNION of both lists (de-duplicated, order preserved).
    ai = payload.get("application_info") or {}
    rb_names: List[str] = []
    seen = set()
    for field in ("regulatory_body_rn", "regulatory_body_lpn"):
        for rb in safe_str(ai.get(field)).split(","):
            rb = rb.strip()
            if rb and rb.lower() not in seen:
                seen.add(rb.lower())
                rb_names.append(rb)
    rb_source = "additional_content" if rb_names else None

    # PDF fallback: the regulatory-body prompt is a MULTI-SELECT, so split the
    # collected answer on commas rather than treating it as one name.
    if not rb_names and form.get("rb_selected"):
        for rb in form["rb_selected"].split(","):
            rb = rb.strip()
            if rb and rb.lower() not in seen:
                seen.add(rb.lower())
                rb_names.append(rb)
        rb_source = "application_summary_pdf"
    debug["regulators"] = rb_names
    debug["regulators_source"] = rb_source

    # A report exists for a regulator the applicant is not recorded against:
    # surface it rather than silently dropping the package.
    resolved_abbrevs = {abbrev_for(rb) for rb in rb_names if abbrev_for(rb)}
    orphan_report_abbrevs = sorted(set(reports_by_abbrev) - resolved_abbrevs)
    if orphan_report_abbrevs:
        warnings.append(
            "Reports found for regulators not present in the resolved regulator "
            f"list, so no package was built for them: {', '.join(orphan_report_abbrevs)}."
        )
    debug["orphan_report_abbrevs"] = orphan_report_abbrevs

    if not shared and not reports_by_abbrev and not applicant_form_local:
        return {"status": "no_files", "message": "No files found to package.",
                "warnings": warnings, "debug": debug}

    # ---- download + rename the SHARED docs ONCE into a shared images tree ----
    shared_root = os.path.join(workdir, "_shared")
    shared_dirs = {cat: os.path.join(shared_root, path) for cat, path in CATEGORY_FOLDERS.items()}
    for d in shared_dirs.values():
        os.makedirs(d, exist_ok=True)

    order = {"identity": 0, "registration": 1, "report": 2, "other": 3}
    shared.sort(key=lambda c: (order.get(c["_category"], 9), c["name"]))

    counters: Dict[Tuple[str, str], int] = {}
    shared_docs: Dict[str, List[Dict[str, str]]] = {"identity": [], "registration": [], "other": []}
    translations_packaged: List[Dict[str, str]] = []

    for c in shared:
        category, doctype, forced = c["_category"], c["_doctype"], c.get("_forced")
        if category not in shared_docs:
            category = "other"
        local = os.path.join(dl_dir, c["name"])
        try:
            download_file(sp, c["url"], local)
        except Exception as e:
            skipped.append(f"{c['name']} (download failed: {e})"); continue
        ext = os.path.splitext(c["name"])[1] or ".pdf"
        if forced is not None:
            number = forced
        else:
            key = (category, doctype)
            counters[key] = counters.get(key, 0) + 1
            number = counters[key]
        new_name = sftp_name(doc_id, app_id_num, category, doctype, number, ext)
        shutil.copyfile(local, os.path.join(shared_dirs[category], new_name))
        block = doc_metadata_block(c.get("props", {}), local, c.get("file"), doctype)
        # Translation files are usually untagged: fall back to the SharePoint
        # upload date for ReceivedDate, matching the date on the report's C4.
        if c.get("_translation") and not block["ReceivedDate"]:
            block["ReceivedDate"] = to_ddmmyyyy(
                c.get("created") or meta_get(c.get("props", {}), "Created"))
        block["Name"] = new_name
        block["Path"] = CATEGORY_FOLDERS[category] + "/"
        shared_docs[category].append(block)
        if c.get("_translation"):
            translations_packaged.append({"source": c["name"],
                                          "packaged_as": f"{CATEGORY_FOLDERS[category]}/{new_name}"})
    debug["translations_packaged"] = translations_packaged

    # ---- mirror the "Additional documents" folder verbatim (files + subfolders) ----
    # Packaged as-is: original filenames and folder structure preserved, no
    # categorization or renaming. Included in every regulator's zip under an
    # "Additional documents/" tree, alongside the existing images/ tree.
    add_docs_local_root = os.path.join(workdir, "_additional_documents")
    add_docs_arcnames: List[Tuple[str, str]] = []  # (local_path, arcname_in_zip)
    add_docs_prefix_lower = add_docs_rel.lower()
    for entry in add_docs_files:
        url = entry.get("url") or ""
        if not url:
            continue
        # Case-insensitive prefix match (SharePoint folder names are not
        # case-sensitive, and Graph may return the canonical casing).
        if url.lower().startswith(add_docs_prefix_lower):
            rel = url[len(add_docs_rel):].lstrip("/")
        else:
            skipped.append(f"{url} (outside {ADDITIONAL_DOCS_SUBFOLDER} prefix)")
            continue
        if not rel:
            continue
        local = os.path.join(add_docs_local_root, *rel.split("/"))
        try:
            download_file(sp, url, local)  # creates parent dirs
        except Exception as e:
            skipped.append(f"{ADDITIONAL_DOCS_SUBFOLDER}/{rel} (download failed: {e})")
            continue
        add_docs_arcnames.append((local, f"{ADDITIONAL_DOCS_SUBFOLDER}/{rel}"))
    debug["additional_docs_packaged"] = [a for _, a in add_docs_arcnames]

    # ---- per-regulator: build a zip each, named <ApplicationID>.zip ----
    ts = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    export_rel = f"{app_base_rel}/{EXPORT_FOLDER_PREFIX}_{ts}"
    ensure_folder(sp, export_rel)

    packages: List[Dict[str, Any]] = []
    no_report: List[str] = []
    # Regulators held back from the SFTP because the CSV or XML check failed.
    check_blocked_regulators: List[str] = []

    # Open one SFTP session for regulator delivery (best-effort: SharePoint
    # upload still succeeds even if SFTP is unreachable). On a dry run, no SFTP
    # session is opened, so the per-regulator delivery block is skipped entirely.
    sftp = transport = None
    sftp_error: Optional[str] = None
    if dry_run:
        sftp_error = "dry run: SFTP push skipped"
    else:
        try:
            sftp, transport = sftp_connect()
        except Exception as e:
            sftp_error = f"SFTP connect failed: {e}"

    applicant_copies: List[Dict[str, Any]] = []

    for rb_name in rb_names:
        abbrev = abbrev_for(rb_name)
        if not abbrev:
            skipped.append(f"{rb_name} (no abbreviation mapping)")
            continue

        # reports for this regulator, keyed by designation (may be 0, 1, or many)
        by_des = reports_by_abbrev.get(abbrev, {})
        if not by_des:
            no_report.append(f"{rb_name} ({abbrev})")

        rb_work = os.path.join(workdir, f"rb_{abbrev}")
        rb_dirs = {cat: os.path.join(rb_work, path) for cat, path in CATEGORY_FOLDERS.items()}
        for d in rb_dirs.values():
            os.makedirs(d, exist_ok=True)
        # copy shared images in (identical across designations)
        for cat in ("identity", "registration", "other"):
            for fn in os.listdir(shared_dirs[cat]):
                shutil.copyfile(os.path.join(shared_dirs[cat], fn),
                                os.path.join(rb_dirs[cat], fn))
        # applicant application-summary form -> images/other/ (same in every package)
        if applicant_form_local and applicant_form_packaged:
            shutil.copyfile(applicant_form_local,
                            os.path.join(rb_dirs["other"], applicant_form_packaged))

        # Build one record per designation. With no reports, fall back to a
        # single record using the RB-list-derived group (no report document).
        designations = list(by_des.items()) if by_des else [(None, None)]
        records: List[Tuple[Dict[str, Any], Dict[str, List[Dict[str, str]]]]] = []
        report_names: List[str] = []
        rep_idx = 0

        for desig, rep in designations:
            rep_idx += 1
            docs = {k: list(v) for k, v in shared_docs.items()}
            docs["report"] = []
            if rep:
                rlocal = os.path.join(dl_dir, f"{abbrev}_{rep['name']}")
                try:
                    download_file(sp, rep["url"], rlocal)
                    ext = os.path.splitext(rep["name"])[1] or ".pdf"
                    # keep the designation visible in the packaged report name
                    rep_name = sftp_name(doc_id, app_id_num, "report",
                                         rep["_doctype"], rep_idx, ext)
                    shutil.copyfile(rlocal, os.path.join(rb_dirs["report"], rep_name))
                    rblock = doc_metadata_block(rep.get("props", {}), rlocal,
                                                rep.get("file"), rep["_doctype"])
                    rblock["Name"] = rep_name
                    rblock["Path"] = CATEGORY_FOLDERS["report"] + "/"
                    docs["report"].append(rblock)
                    report_names.append(rep["name"])

                    # ---- applicant copy: strip the Assessment Notes block ----
                    try:
                        ac_local = os.path.join(dl_dir, applicant_copy_name(rep["name"]))
                        removed = strip_assessment_notes(rlocal, ac_local)
                        # upload to <regulator report folder>/Applicant Copy/
                        rep_parent = os.path.dirname(rep["url"])
                        ac_folder = f"{rep_parent}/{APPLICANT_COPY_SUBFOLDER}"
                        ac_dest = upload_file(sp, ac_local, ac_folder,
                                              applicant_copy_name(rep["name"]))
                        applicant_copies.append({
                            "abbrev": abbrev, "designation": desig,
                            "source": rep["name"],
                            "applicant_copy": ac_dest,
                            "assessment_notes_removed": removed,
                        })
                    except Exception as e:
                        skipped.append(f"applicant copy failed for {rep['name']}: {e}")
                except Exception as e:
                    skipped.append(f"{rep['name']} (download failed: {e})")

            # model for this designation: NursingGroup follows the report's
            # designation when present, else the RB-list-derived group.
            model = build_applicant_model(form, payload, doc_id, app_id_num,
                                          rb_name=rb_name)
            if desig:
                model["NursingGroup"] = desig
            records.append((model, docs))

        # serialize CSV (one row per designation) + XML (one Application each)
        csv_text = build_csv(records)
        xml_text = build_xml(records)

        # ---- CSV + XML output checks (run on dry runs too) ----
        csv_check = verify_csv_output(csv_text, records, app_id_num)
        warnings.extend(f"{abbrev.upper()}: {w}" for w in csv_check["warnings"])

        xml_check = verify_xml_output(xml_text, records)
        warnings.extend(f"{abbrev.upper()}: {w}" for w in xml_check["warnings"])

        blocked_by: List[str] = []
        if CSV_CHECK_BLOCKS_DELIVERY and not csv_check["ok"]:
            blocked_by.append("CSV")
        if XML_CHECK_BLOCKS_DELIVERY and not xml_check["ok"]:
            blocked_by.append("XML")
        check_blocked = bool(blocked_by)
        if check_blocked:
            check_blocked_regulators.append(abbrev)

        csv_name = f"{doc_id}_{app_id_num}.csv"
        xml_name = f"{doc_id}_{app_id_num}.xml"
        csv_path = os.path.join(rb_work, csv_name)
        xml_path = os.path.join(rb_work, xml_name)
        with open(csv_path, "w", newline="", encoding="utf-8-sig") as fh:
            fh.write(csv_text)
        with open(xml_path, "w", encoding="utf-8") as fh:
            fh.write(xml_text)

        # one zip per regulator, containing all designations' reports.
        zip_name = f"{app_id_num}.zip"
        zip_path = os.path.join(workdir, f"{abbrev}_{app_id_num}.zip")  # local: keep distinct
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.write(csv_path, arcname=csv_name)
            zf.write(xml_path, arcname=xml_name)
            for cat, path in CATEGORY_FOLDERS.items():
                d = rb_dirs[cat]
                for fn in sorted(os.listdir(d)):
                    zf.write(os.path.join(d, fn), arcname=f"{path}/{fn}")
            # include the "Additional documents" folder verbatim
            for local_path, arcname in add_docs_arcnames:
                zf.write(local_path, arcname=arcname)

        # SharePoint export: since every zip is "{ApplicationID}.zip", give
        # each regulator its own subfolder under the export folder to avoid a
        # name collision.
        sp_export_dir = f"{export_rel}/{abbrev.upper()}"
        uploaded = upload_file(sp, zip_path, sp_export_dir, zip_name)

        # ---- loose CSV + XML beside the zip (links only in the run output) ----
        csv_url: Optional[str] = None
        xml_url: Optional[str] = None
        if EXPORT_LOOSE_INTERCHANGE:
            try:
                csv_url = upload_file(sp, csv_path, sp_export_dir, csv_name)
                xml_url = upload_file(sp, xml_path, sp_export_dir, xml_name)
            except Exception as e:
                skipped.append(f"{abbrev}: loose CSV/XML upload failed: {e}")

        # ---- regulator SFTP delivery ----
        # zip  -> /REGULATORY/<RB>/<ApplicationID>/<ApplicationID>.zip
        # index-> /REGULATORY/<RB>/<RB>_NNAS_index_<DDMMYYYY>.csv (one row/designation)
        # Held back when the CSV or XML check failed and the matching
        # *_CHECK_BLOCKS_DELIVERY flag is on.
        sftp_dest = None
        sftp_index_paths: List[str] = []
        rb_sftp_error = None
        if check_blocked and sftp is not None:
            rb_sftp_error = (f"{abbrev}: SFTP delivery held back because the "
                             f"{' and '.join(blocked_by)} check failed")
            skipped.append(rb_sftp_error)
        elif sftp is not None:
            try:
                regulator_dir = sftp_regulator_dir(sftp, abbrev)
                app_dir = sftp_app_dir(sftp, regulator_dir, app_id_num)
                sftp_dest = sftp_upload_zip(sftp, zip_path, app_dir, zip_name)
                idx_name = index_filename(abbrev)
                for model, _docs in records:
                    sftp_index_paths.append(
                        sftp_append_index(sftp, regulator_dir, idx_name,
                                          build_index_row(model)))
            except Exception as e:
                rb_sftp_error = f"{abbrev}: {e}"
                skipped.append(f"SFTP delivery failed for {abbrev}: {e}")

        packages.append({
            "regulator": rb_name, "abbrev": abbrev,
            "zip_name": zip_name, "uploaded_zip": uploaded,
            "csv_file": csv_url,
            "xml_file": xml_url,
            "reports": report_names,
            "applicant_form": applicant_form_packaged if applicant_form_local else None,
            "translations": [t["packaged_as"] for t in translations_packaged],
            "designations": [m["NursingGroup"] for m, _ in records],
            "doc_counts": {k: len(v) for k, v in records[0][1].items()},
            "report_count": len(report_names),
            "csv_check": csv_check,
            "xml_check": xml_check,
            "blocked_by": blocked_by,
            "sftp_dest": sftp_dest,
            "sftp_index": sftp_index_paths,
            "sftp_error": rb_sftp_error,
        })

    # close SFTP session
    if sftp is not None:
        try:
            sftp.close()
        except Exception:
            pass
    if transport is not None:
        try:
            transport.close()
        except Exception:
            pass

    if not packages:
        return {"status": "no_packages",
                "message": "No regulator packages produced (no regulators resolved/abbreviated).",
                "warnings": warnings,
                "debug": debug}

    debug["shared_doc_counts"] = {k: len(v) for k, v in shared_docs.items()}
    debug["regulators_without_report"] = no_report
    debug["check_blocked"] = check_blocked_regulators
    debug["sftp_connect_error"] = sftp_error
    debug["dry_run"] = dry_run
    debug["skipped"] = skipped
    return {
        "status": "ok",
        "reference_id": folder_app_id,
        "id": doc_id,
        "application_id": app_id_num,
        "export_folder": export_rel,
        "sftp_host": f"{SFTP_HOST}:{SFTP_PORT}",
        "dry_run": dry_run,
        "sftp_delivered": (not dry_run) and sftp_error is None and not check_blocked_regulators,
        "packages": packages,
        "package_count": len(packages),
        "applicant_copies": applicant_copies,
        "warnings": warnings,
        "debug": debug,
    }