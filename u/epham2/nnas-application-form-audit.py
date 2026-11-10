# py: 3.12
# requirements:
# wmill
# requests

"""
NNAS Applicant Form Audit  (READ-ONLY Windmill script)
======================================================

Finds applications whose most recent SFTP package was sent WITHOUT the
Applicant Form. Nothing is written, moved or sent: the script only reads
SharePoint and returns a report.

For every application folder under each provider folder (ces/, tm/):

  1. Looks in the application root for stamped application summaries
     ('<ReferenceID>_<YYYYMMDD>_<HHMMSS>.<ext>'), the same files the package
     builder packages as the Applicant Form.
  2. Finds the most recent 'SFTP_Export_<YYYYMMDD_HHMMSS>' folder. Only the
     latest export matters, because each repost replaces the regulator's zip.
  3. Opens the zip(s) in that export and reads ONLY the zip's file list
     (HTTP range requests against the central directory, so whole zips are
     not downloaded). It looks for an 'Applicant Form' file under
     images/other/.
  4. Classifies the application:
        OK               the latest zip already contains the Applicant Form
        REPOST           form missing from the zip, but a stamped form is on
                         file in SharePoint (fixable with the repost script)
        NO FORM ON FILE  form missing and no stamped form in the root
                         (needs manual follow-up)
        NOT EXPORTED     no SFTP_Export folder found
        NO ZIP IN EXPORT latest export folder contains no zip
        ERROR            the folder or zip could not be read (see note)

Both export layouts are handled:
    SFTP_Export_<ts>/<RB>/<ApplicationID>.zip        (current)
    SFTP_Export_<ts>/<ApplicationID>-<abbrev>.zip     (older)

By default only ONE zip per export is opened: every regulator's zip in a run
is built from the same Applicant Form copy, so one zip answers the question
for the whole export. Set check_all_regulators=True to open every zip.

The report contains application IDs and flags only, no applicant names or
personal details.

Large libraries: use max_apps and start_after to run the sweep in batches.
"""

import io
import re
import csv
import time
import zipfile
from io import StringIO
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import quote

import wmill
import requests


# =========================================================
# CONFIG
# =========================================================

TENANT_ID     = wmill.get_variable("u/epham2/AZURE_TENANT_ID")
CLIENT_ID     = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_ID")
CLIENT_SECRET = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_SECRET")

GRAPH_BASE            = "https://graph.microsoft.com/v1.0"
SHAREPOINT_HOSTNAME   = "nnasca.sharepoint.com"
SITE_PATH             = "/sites/NNAS-Portal-Sharepoint"
DOCUMENT_LIBRARY_NAME = "Applications"

DEFAULT_PROVIDERS = ["ces", "tm"]

# Application folders are named by reference ID (e.g. NNAS-C-0000000481).
# Anything else under a provider folder is ignored and counted.
APP_FOLDER_RE = re.compile(r"^NNAS-", re.IGNORECASE)

EXPORT_FOLDER_RE = re.compile(r"^SFTP_Export_(\d{8}_\d{6})$", re.IGNORECASE)

# What the package builder writes into images/other/:
#   '<ReferenceID>-Applicant Form - <YYYYMMDD_HHMMSS>.<ext>'
FORM_ZIP_FOLDER = "images/other/"
FORM_ZIP_MARKER = "applicant form"
STAMP_RE        = re.compile(r"(\d{8}_\d{6})")

MAX_RETRIES   = 6
RANGE_BUFFER  = 64 * 1024      # bytes per range request

ACTIONS = ("OK", "REPOST", "NO FORM ON FILE", "NOT EXPORTED",
           "NO ZIP IN EXPORT", "ERROR")

_GRAPH_TOKEN: Optional[str] = None
_DRIVE_ID: Optional[str] = None
_SESSION = requests.Session()
_STATS = {"graph_calls": 0, "range_requests": 0, "full_downloads": 0, "retries": 0}


# =========================================================
# GRAPH HELPERS  (read-only)
# =========================================================

def _get_graph_token(force: bool = False) -> str:
    global _GRAPH_TOKEN
    if _GRAPH_TOKEN and not force:
        return _GRAPH_TOKEN
    for name, val in (("AZURE_TENANT_ID", TENANT_ID),
                      ("SHAREPOINT_CLIENT_ID", CLIENT_ID),
                      ("SHAREPOINT_CLIENT_SECRET", CLIENT_SECRET)):
        if not val:
            raise RuntimeError(f"Missing Windmill variable: u/epham2/{name}")
    resp = requests.post(
        f"https://login.microsoftonline.com/{TENANT_ID.strip()}/oauth2/v2.0/token",
        data={
            "client_id": CLIENT_ID.strip(),
            "client_secret": CLIENT_SECRET.strip(),
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        },
        timeout=60,
    )
    if resp.status_code != 200:
        raise RuntimeError(f"Graph token request failed ({resp.status_code}): {resp.text[:1000]}")
    _GRAPH_TOKEN = resp.json()["access_token"]
    return _GRAPH_TOKEN


def _send(method: str, url: str, headers: Optional[Dict[str, str]] = None,
          timeout: int = 180) -> requests.Response:
    """HTTP with retry on throttling / transient errors (429, 502, 503, 504)."""
    resp = None
    for attempt in range(MAX_RETRIES):
        resp = _SESSION.request(method, url, headers=headers or {}, timeout=timeout)
        if resp.status_code in (429, 502, 503, 504):
            _STATS["retries"] += 1
            try:
                wait = int(resp.headers.get("Retry-After", "0"))
            except ValueError:
                wait = 0
            time.sleep(wait or min(2 ** attempt * 2, 60))
            continue
        return resp
    return resp


def _graph_get(endpoint: str, allow_404: bool = False) -> Optional[Dict[str, Any]]:
    """GET a Graph endpoint (or a full nextLink URL). Refreshes the token once on 401."""
    url = endpoint if endpoint.startswith("http") else f"{GRAPH_BASE}{endpoint}"
    for refresh in (False, True):
        _STATS["graph_calls"] += 1
        resp = _send("GET", url, headers={"Authorization": f"Bearer {_get_graph_token(force=refresh)}"})
        if resp.status_code == 401 and not refresh:
            continue  # token expired during a long sweep
        if resp.status_code == 404 and allow_404:
            return None
        if resp.status_code != 200:
            raise RuntimeError(f"Graph GET failed ({resp.status_code}) {endpoint}: {resp.text[:1000]}")
        return resp.json()
    return None


def _init_drive() -> str:
    global _DRIVE_ID
    if _DRIVE_ID:
        return _DRIVE_ID
    site = _graph_get(f"/sites/{SHAREPOINT_HOSTNAME}:{SITE_PATH}")
    drives = _graph_get(f"/sites/{site['id']}/drives")["value"]
    drive = next((d for d in drives if d.get("name") == DOCUMENT_LIBRARY_NAME), None)
    if not drive:
        raise RuntimeError(f"Drive '{DOCUMENT_LIBRARY_NAME}' not found. "
                           f"Available: {[d.get('name') for d in drives]}")
    _DRIVE_ID = drive["id"]
    return _DRIVE_ID


def get_item_by_path(path: str) -> Optional[Dict[str, Any]]:
    drive = _init_drive()
    return _graph_get(f"/drives/{drive}/root:/{quote(path.strip('/'), safe='/')}", allow_404=True)


def list_children(item_id: str) -> List[Dict[str, Any]]:
    drive = _init_drive()
    out: List[Dict[str, Any]] = []
    endpoint: Optional[str] = f"/drives/{drive}/items/{item_id}/children?$top=999"
    while endpoint:
        data = _graph_get(endpoint)
        out.extend(data.get("value", []))
        endpoint = data.get("@odata.nextLink")
    return out


def is_folder(item: Dict[str, Any]) -> bool:
    return item.get("folder") is not None


def is_file(item: Dict[str, Any]) -> bool:
    return item.get("file") is not None


# =========================================================
# ZIP LISTING WITHOUT DOWNLOADING THE WHOLE ZIP
# =========================================================

def _range_fetch(url: str, start: int, end: int) -> Tuple[int, bytes]:
    """Fetch bytes [start, end] of a pre-authenticated download URL."""
    _STATS["range_requests"] += 1
    resp = _send("GET", url, headers={"Range": f"bytes={start}-{end}"}, timeout=300)
    if resp.status_code not in (200, 206):
        raise RuntimeError(f"Range read failed ({resp.status_code})")
    return resp.status_code, resp.content


class HttpRangeReader(io.RawIOBase):
    """
    Seekable, read-only file over HTTP Range requests. zipfile only needs the
    end-of-central-directory record and the central directory, so listing a
    zip costs a few small requests. If the server ignores Range (returns 200),
    the full body is kept in memory and served from there.
    """

    def __init__(self, url: str, size: int):
        super().__init__()
        self._url = url
        self._size = size
        self._pos = 0
        self._full: Optional[bytes] = None

    def readable(self) -> bool:
        return True

    def seekable(self) -> bool:
        return True

    def tell(self) -> int:
        return self._pos

    def seek(self, offset: int, whence: int = io.SEEK_SET) -> int:
        if whence == io.SEEK_SET:
            self._pos = offset
        elif whence == io.SEEK_CUR:
            self._pos += offset
        elif whence == io.SEEK_END:
            self._pos = self._size + offset
        else:
            raise ValueError(f"invalid whence: {whence}")
        self._pos = max(0, self._pos)
        return self._pos

    def readinto(self, b) -> int:
        if self._pos >= self._size:
            return 0
        end = min(self._pos + len(b), self._size) - 1
        if self._full is not None:
            data = self._full[self._pos:end + 1]
        else:
            status, data = _range_fetch(self._url, self._pos, end)
            if status == 200:                 # Range not honored: whole file came back
                _STATS["full_downloads"] += 1
                self._full = data
                data = self._full[self._pos:end + 1]
        n = len(data)
        b[:n] = data
        self._pos += n
        return n


def zip_namelist(zip_item: Dict[str, Any]) -> List[str]:
    """File names inside a SharePoint zip, read via range requests."""
    drive = _init_drive()
    # Re-fetch the item: download URLs are short-lived.
    item = _graph_get(f"/drives/{drive}/items/{zip_item['id']}")
    url = item.get("@microsoft.graph.downloadUrl")
    size = int(item.get("size") or 0)
    if url and size:
        raw = HttpRangeReader(url, size)
        with zipfile.ZipFile(io.BufferedReader(raw, buffer_size=RANGE_BUFFER)) as zf:
            return zf.namelist()
    # Fallback: full download through the authenticated content endpoint.
    _STATS["full_downloads"] += 1
    resp = _send("GET", f"{GRAPH_BASE}/drives/{drive}/items/{zip_item['id']}/content",
                 headers={"Authorization": f"Bearer {_get_graph_token()}"}, timeout=300)
    if resp.status_code != 200:
        raise RuntimeError(f"Zip download failed ({resp.status_code})")
    with zipfile.ZipFile(io.BytesIO(resp.content)) as zf:
        return zf.namelist()


# =========================================================
# CLASSIFICATION HELPERS
# =========================================================

def applicant_form_stamp(file_name: str, app_id: str) -> Optional[str]:
    """'YYYYMMDD_HHMMSS' if file_name is exactly '<app_id>_<date>_<time>.<ext>'
    (same rule as the package builder)."""
    stem = file_name.rsplit(".", 1)[0] if "." in file_name else file_name
    m = re.match(rf"^{re.escape(app_id)}_(\d{{8}})_(\d{{6}})$", stem, re.IGNORECASE)
    return f"{m.group(1)}_{m.group(2)}" if m else None


def form_entries_in_zip(names: List[str]) -> List[str]:
    return [n for n in names
            if n.lower().startswith(FORM_ZIP_FOLDER) and FORM_ZIP_MARKER in n.lower()]


def regulator_from_zip_name(zip_name: str, app_id: str) -> str:
    """Older layout '<ApplicationID>-<abbrev>.zip' -> 'ABBREV'; else ''."""
    stem = zip_name.rsplit(".", 1)[0]
    prefix = f"{app_id}-"
    if stem.lower().startswith(prefix.lower()) and len(stem) > len(prefix):
        return stem[len(prefix):].upper()
    return ""


def collect_export_zips(export_item: Dict[str, Any], app_id: str) -> List[Tuple[str, Dict[str, Any]]]:
    """(regulator, zip_item) pairs in an export folder, for both layouts."""
    zips: List[Tuple[str, Dict[str, Any]]] = []
    for child in list_children(export_item["id"]):
        name = child.get("name", "")
        if is_folder(child):
            for z in list_children(child["id"]):
                if is_file(z) and z.get("name", "").lower().endswith(".zip"):
                    zips.append((name.upper(), z))
        elif is_file(child) and name.lower().endswith(".zip"):
            zips.append((regulator_from_zip_name(name, app_id), child))
    zips.sort(key=lambda t: (t[0], t[1].get("name", "")))
    return zips


# =========================================================
# PER-APPLICATION AUDIT
# =========================================================

def _row(provider: str, app_id: str, **kw: Any) -> Dict[str, Any]:
    row = {
        "application_id": app_id,
        "provider": provider.upper(),
        "export_timestamp": "",
        "export_count": 0,
        "regulators_in_export": "",
        "regulator_checked": "",
        "zip_name": "",
        "form_in_zip": "",
        "zip_form_stamp": "",
        "stamped_form_on_file": "",
        "latest_stamp_on_file": "",
        "action": "",
        "note": "",
    }
    row.update(kw)
    return row


def audit_application(provider: str, app_folder: Dict[str, Any],
                      check_all_regulators: bool) -> List[Dict[str, Any]]:
    app_id = app_folder["name"]
    children = list_children(app_folder["id"])

    stamps = sorted(s for s in (applicant_form_stamp(c.get("name", ""), app_id)
                                for c in children if is_file(c)) if s)
    latest_stamp = stamps[-1] if stamps else ""

    exports = sorted(((m.group(1), c) for c in children if is_folder(c)
                      for m in [EXPORT_FOLDER_RE.match(c.get("name", ""))] if m),
                     key=lambda t: t[0])

    base = dict(stamped_form_on_file="Y" if latest_stamp else "N",
                latest_stamp_on_file=latest_stamp,
                export_count=len(exports))

    if not exports:
        return [_row(provider, app_id, action="NOT EXPORTED", **base)]

    export_ts, export_item = exports[-1]
    base["export_timestamp"] = export_ts

    zips = collect_export_zips(export_item, app_id)
    base["regulators_in_export"] = ";".join(r or "?" for r, _ in zips)
    if not zips:
        return [_row(provider, app_id, action="NO ZIP IN EXPORT",
                     note=f"{export_item['name']} has no zip", **base)]

    rows: List[Dict[str, Any]] = []
    read_errors: List[str] = []
    for regulator, z in zips:
        try:
            names = zip_namelist(z)
        except Exception as e:
            read_errors.append(f"{regulator or z.get('name')}: {e}")
            if check_all_regulators:
                rows.append(_row(provider, app_id, regulator_checked=regulator,
                                 zip_name=z.get("name", ""), action="ERROR",
                                 note=f"zip unreadable: {e}"[:300], **base))
            continue  # try the next zip

        entries = form_entries_in_zip(names)
        row_kw = dict(regulator_checked=regulator, zip_name=z.get("name", ""), **base)
        if entries:
            m = STAMP_RE.search(entries[0])
            zip_stamp = m.group(1) if m else ""
            note = ""
            if latest_stamp and zip_stamp and latest_stamp > zip_stamp:
                note = f"newer form on file ({latest_stamp}) than the one sent ({zip_stamp})"
            rows.append(_row(provider, app_id, form_in_zip="Y", zip_form_stamp=zip_stamp,
                             action="OK", note=note, **row_kw))
        elif latest_stamp:
            rows.append(_row(provider, app_id, form_in_zip="N", action="REPOST", **row_kw))
        else:
            rows.append(_row(provider, app_id, form_in_zip="N", action="NO FORM ON FILE",
                             note="no '<ReferenceID>_<YYYYMMDD>_<HHMMSS>' file in the application root",
                             **row_kw))

        if not check_all_regulators:
            break   # one readable zip answers the question for this export

    if not rows:
        rows.append(_row(provider, app_id, action="ERROR",
                         note=("no zip in the latest export could be read: "
                               + " | ".join(read_errors))[:500], **base))
    return rows


# =========================================================
# ENTRYPOINT
# =========================================================

CSV_COLUMNS = ["application_id", "provider", "export_timestamp", "export_count",
               "regulators_in_export", "regulator_checked", "zip_name", "form_in_zip",
               "zip_form_stamp", "stamped_form_on_file", "latest_stamp_on_file",
               "action", "note"]


def main(providers: Optional[List[str]] = None,
         app_ids: Optional[List[str]] = None,
         start_after: Optional[str] = None,
         max_apps: int = 0,
         check_all_regulators: bool = False):
    """
    providers            provider folders to scan (default: ces, tm)
    app_ids              audit only these application IDs (optional)
    start_after          skip application IDs up to and including this one
                         (alphabetical), to resume a batched sweep
    max_apps             stop after this many applications (0 = no limit)
    check_all_regulators open every regulator zip instead of one per export

    Read-only: nothing in SharePoint is created, changed or deleted.
    """
    started = time.time()
    providers = [p.strip().lower() for p in (providers or DEFAULT_PROVIDERS) if p and p.strip()]
    wanted = {a.strip().lower() for a in (app_ids or []) if a and a.strip()}

    rows: List[Dict[str, Any]] = []
    provider_errors: List[str] = []
    non_app_folders = 0
    apps_scanned = 0
    last_app_id = ""
    stopped_early = False

    for provider in providers:
        root = get_item_by_path(provider)
        if not root or not is_folder(root):
            provider_errors.append(f"provider folder '{provider}' not found")
            continue

        app_folders = []
        for c in list_children(root["id"]):
            if not is_folder(c):
                continue
            if not APP_FOLDER_RE.match(c.get("name", "")):
                non_app_folders += 1
                continue
            app_folders.append(c)
        app_folders.sort(key=lambda c: c["name"].lower())

        for folder in app_folders:
            name = folder["name"]
            if wanted and name.lower() not in wanted:
                continue
            if start_after and name.lower() <= start_after.strip().lower():
                continue
            if max_apps and apps_scanned >= max_apps:
                stopped_early = True
                break
            try:
                rows.extend(audit_application(provider, folder, check_all_regulators))
            except Exception as e:
                rows.append(_row(provider, name, action="ERROR",
                                 note=f"application folder unreadable: {e}"[:500]))
            apps_scanned += 1
            last_app_id = name
        if stopped_early:
            break

    # ---- summaries (by application, worst action wins when several rows) ----
    severity = {a: i for i, a in enumerate(ACTIONS)}  # OK lowest, ERROR highest
    per_app: Dict[Tuple[str, str], str] = {}
    for r in rows:
        key = (r["provider"], r["application_id"])
        cur = per_app.get(key)
        if cur is None or severity[r["action"]] > severity[cur]:
            per_app[key] = r["action"]

    summary = {a: 0 for a in ACTIONS}
    for action in per_app.values():
        summary[action] += 1

    def ids_with(action: str) -> List[str]:
        return sorted(app for (_, app), a in per_app.items() if a == action)

    buf = StringIO()
    writer = csv.DictWriter(buf, fieldnames=CSV_COLUMNS, lineterminator="\n")
    writer.writeheader()
    for r in rows:
        writer.writerow(r)

    return {
        "status": "ok" if not provider_errors else "partial",
        "read_only": True,
        "providers": providers,
        "apps_scanned": apps_scanned,
        "summary_by_application": summary,
        "repost_app_ids": ids_with("REPOST"),
        "no_form_on_file_app_ids": ids_with("NO FORM ON FILE"),
        "not_exported_app_ids": ids_with("NOT EXPORTED"),
        "error_app_ids": ids_with("ERROR"),
        "stopped_early": stopped_early,
        "resume_with_start_after": last_app_id if stopped_early else None,
        "provider_errors": provider_errors,
        "non_app_folders_ignored": non_app_folders,
        "stats": {**_STATS, "elapsed_seconds": round(time.time() - started, 1)},
        "generated_at_utc": datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S"),
        "csv": buf.getvalue(),
    }