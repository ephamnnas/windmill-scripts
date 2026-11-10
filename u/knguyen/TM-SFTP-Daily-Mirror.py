# py: 3.12
# requirements:
# wmill
# paramiko

"""
NNAS Regulator Mirror  (scheduled Windmill script -- run every 6 hours)
=======================================================================

Mirrors applicant packages from the TruMerit source SFTP to the NNAS regulator
delivery SFTP, driven by each regulator's daily tracking index.

Direction:
    SOURCE  ftp.nnas.ca:22            (vars trumerit_FTP_username / trumerit_ftp_pw)
    DEST    rb-ftp-connect.nnas.ca:2022 (same creds as the package builder)

Per run, for every regulator (RB) folder under the source root:
  1. For each of the last DAYS_BACK calendar days, look for the daily tracking
     file  <RB>_NNAS_index_<DDMMYYYY>.csv.
  2. Read its rows. Each row is the 9-column NNAS index:
        LASTNAME,FIRSTNAME,id,applicationID,DOB,RN|LPN,x,N,x
     The applicant folder is named by the applicationID (column 4).
  3. Reconcile that tracking file into the destination RB folder. The
     DESTINATION is the authoritative accumulator:
        - if no dated CSV exists there yet, the source file is written verbatim
          (real rows OR the "No Applicants" sentinel);
        - if a dated CSV already exists there, ONLY the source rows not already
          present are appended -- a "No Applicants" source never overwrites the
          real rows already delivered, and real source rows replace a stale
          "No Applicants" sentinel rather than being appended after it.
  4. Mirror each referenced applicant folder (recursively, with all subfolders)
     to the destination RB folder. Files already present with the same size are
     skipped (idempotent across the 6-hour cadence).

Both endpoints use SFTP (paramiko). Destination RB folders are matched by name
first (direct), falling back to RB_FALLBACK_MAP for renamed/merged regulators;
destination RB folders are never created.
"""

import io
import os
import csv
import stat
import tempfile
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional, Tuple

import wmill
import paramiko


# =========================================================
# CONFIG
# =========================================================

# --- CREDENTIAL OVERRIDES (debug / fallback) ---------------------------------
# Leave these BLANK to use the Windmill variables (the secure default).
# If set (non-empty), they take precedence over the variables for that field.
# WARNING: values here are plaintext in the script and its history. Use only to
# diagnose an auth failure, then clear them and revert to the variables.
SRC_USERNAME_OVERRIDE  = ""
SRC_PASSWORD_OVERRIDE  = ""
DEST_USERNAME_OVERRIDE = ""
DEST_PASSWORD_OVERRIDE = ""

# --- SOURCE (read from) ---
SRC_HOST = "ftp.nnas.ca"
SRC_PORT = 22
SRC_USERNAME_VAR = "u/epham2/trumerit_FTP_username"
SRC_PASSWORD_VAR = "u/epham2/trumerit_ftp_pw"
SRC_ROOT = "/"                      # RB folders live directly under this root

# --- DESTINATION (write to) ---
DEST_HOST = "rb-ftp-connect.nnas.ca"
DEST_PORT = 2022
DEST_USERNAME_VAR = "u/epham2/NNAS_SFTP_USERNAME"
DEST_PASSWORD_VAR = "u/epham2/NNAS_SFTP_PASSWORD"
DEST_ROOT = "/REGULATORY"           # RB folders live under /REGULATORY/ on the destination

# Source (FTP long code) -> destination (/REGULATORY/ short) RB folder.
# Consulted ONLY as a FALLBACK when a direct same-name match is not found on
# the destination (see resolve_dest_rb). Derived from the regulator
# amalgamation/rename table; several source folders may map to one destination
# folder (merged colleges), which is fine -- they converge into the same dest
# folder and the index dedupe handles overlaps.
RB_FALLBACK_MAP: Dict[str, str] = {
    # --- renames (old source name -> new dest name) ---
    "CARNA":  "CRNA",      # Alberta RN
    "SRNA":   "CRNS",      # Saskatchewan RN
    "NANB":   "CNNB",      # New Brunswick RN
    "SALPN":  "CLPNS",     # Saskatchewan LPN
    "RPNAS":  "CRPNS",     # Saskatchewan RPN
    "CRPA":   "CRPNA",     # Alberta RPN (source folder is usually CRPNA -> direct)
    # --- BC merge -> BCCNM ---
    "CRNBC":  "BCCNM",
    "CLPNBC": "BCCNM",
    "CRPNBC": "BCCNM",
    # --- Newfoundland & Labrador merge -> NLCN ---
    "ARNNL":  "NLCN",
    "CLPNNL": "NLCN",
    # --- Nova Scotia merge -> NSCN ---
    "CRNNS":  "NSCN",
    "CLPNNS": "NSCN",
    # --- PEI merge -> PEICNM ---
    "ARNPEI": "PEICNM",
    "PEILPN": "PEICNM",
    # --- Alberta LPN (confirmed against source FTP listing) ---
    "CLPNA":  "CLHA",
    # Same-name on both sides (ANBLPN, CNO, CRNM, CLPNM, CRPNM, CRPNA) are
    # handled by the DIRECT match and need no entry here.
}

# Folders under SRC_ROOT that are NOT regulators (skipped).
SKIP_FOLDERS = {"CGFNS_TEMP", "NNASDEV", "RDX", "RDXTEST", "TO_TRUMERIT", "Forms"}

# Daily tracking filename pattern (matches the package builder + dest).
INDEX_TEMPLATE = "{rb}_NNAS_index_{ddmmyyyy}.csv"

# How many calendar days back (incl. today) to process each run.
DAYS_BACK = 5

# 0-based column in the index row that holds the applicant folder name.
#   LASTNAME(0) FIRSTNAME(1) id(2) applicationID(3) DOB(4) RN|LPN(5) x(6) N(7) x(8)
APP_ID_COLUMN = 3

# An index file containing only this sentinel text means there are no applicants
# for that regulator/day. The sentinel is only written to the destination when
# the destination has no dated CSV yet -- it NEVER overwrites real rows already
# delivered (the destination is the authoritative accumulator).
NO_APPLICANTS_SENTINEL = "No Applicants"

# Skip re-uploading a file whose destination copy already has the same size.
SKIP_UNCHANGED = True


# =========================================================
# SFTP CONNECT  (password, then keyboard-interactive fallback)
# =========================================================

def _connect(host: str, port: int, user_var: str, pw_var: str,
             user_override: str = "", pw_override: str = ""
             ) -> Tuple[paramiko.SFTPClient, paramiko.Transport]:
    # Override (if set) wins over the Windmill variable. Strip stray
    # whitespace/newlines that commonly break pasted secrets.
    username = (user_override or wmill.get_variable(user_var) or "").strip()
    password = (pw_override or wmill.get_variable(pw_var) or "")
    # password: strip only trailing newline/CR, keep any intentional spaces
    password = password.rstrip("\r\n")
    if not username or not password:
        missing = []
        if not username:
            missing.append(f"username ({'override' if user_override else user_var})")
        if not password:
            missing.append(f"password ({'override' if pw_override else pw_var})")
        raise RuntimeError(f"Credentials not resolved for {host}: {', '.join(missing)}")

    transport = paramiko.Transport((host, port))
    transport.connect(hostkey=None)
    try:
        transport.auth_password(username=username, password=password)
    except paramiko.AuthenticationException:
        try:
            transport.auth_interactive(username=username,
                                       handler=lambda t, i, prompts: [password for _ in prompts])
        except Exception:
            transport.close()
            raise RuntimeError(f"Authentication failed for {host} (user '{username}').")
    if not transport.is_authenticated():
        transport.close()
        raise RuntimeError(f"Authentication failed for {host} (user '{username}').")
    return paramiko.SFTPClient.from_transport(transport), transport


# =========================================================
# SFTP HELPERS
# =========================================================

def _isdir(sftp: paramiko.SFTPClient, path: str) -> bool:
    try:
        return stat.S_ISDIR(sftp.stat(path).st_mode)
    except IOError:
        return False


def _exists(sftp: paramiko.SFTPClient, path: str) -> bool:
    try:
        sftp.stat(path)
        return True
    except IOError:
        return False


def _ensure_dir_under(sftp: paramiko.SFTPClient, base: str, path: str) -> None:
    """
    Create directory components of `path` that live UNDER `base`, which must
    already exist. The base (the regulator folder and everything above it) is
    NEVER created -- if it is missing, that is an error the caller handles.
    Only per-applicant subfolders below the RB folder are created here.
    """
    if not _isdir(sftp, base):
        raise RuntimeError(f"destination base folder does not exist: {base}")
    base = "/" + "/".join(p for p in base.split("/") if p)
    full = "/" + "/".join(p for p in path.split("/") if p)
    if full == base:
        return
    if not full.startswith(base + "/"):
        raise RuntimeError(f"refusing to create '{full}' outside base '{base}'")
    rel_parts = full[len(base) + 1:].split("/")
    cur = base
    for p in rel_parts:
        cur = f"{cur}/{p}"
        if not _isdir(sftp, cur):
            try:
                sftp.mkdir(cur)
            except IOError:
                if not _isdir(sftp, cur):
                    raise


def _list_dirs(sftp: paramiko.SFTPClient, path: str) -> List[str]:
    try:
        return [e.filename for e in sftp.listdir_attr(path) if stat.S_ISDIR(e.st_mode)]
    except IOError:
        return []


def resolve_dest_rb(dest: paramiko.SFTPClient, src_rb: str) -> Tuple[Optional[str], str]:
    """
    Resolve the destination RB folder name for a source RB folder.

    1. DIRECT match: if a folder of the same name exists under DEST_ROOT, use it.
    2. FALLBACK:     otherwise consult RB_FALLBACK_MAP (source code -> dest name).
    3. UNMAPPED:     otherwise return (None, "unmapped") -- caller skips + reports.

    Returns (dest_rb_name | None, mode).
    """
    direct = f"{DEST_ROOT.rstrip('/')}/{src_rb}"
    if _isdir(dest, direct):
        return src_rb, "direct"
    mapped = RB_FALLBACK_MAP.get(src_rb.upper())
    if mapped:
        return mapped, "fallback"
    return None, "unmapped"


def _read_text(sftp: paramiko.SFTPClient, path: str) -> Optional[str]:
    try:
        with sftp.open(path, "r") as fh:
            return fh.read().decode("utf-8", errors="replace")
    except IOError:
        return None


def _write_text(sftp: paramiko.SFTPClient, path: str, text: str) -> None:
    with sftp.open(path, "w") as fh:
        fh.write(text)


# =========================================================
# TRACKING INDEX
# =========================================================

def index_filename(rb: str, day: datetime) -> str:
    return INDEX_TEMPLATE.format(rb=rb.upper(), ddmmyyyy=day.strftime("%d%m%Y"))


def _row_key(row: List[str]) -> Tuple[str, ...]:
    """Normalized identity of an index row, for dedupe."""
    return tuple(c.strip() for c in row)


def is_no_applicants(text: Optional[str]) -> bool:
    """True if a tracking file is the 'No Applicants' sentinel (no data rows)."""
    if text is None:
        return False
    return text.strip().lower() == NO_APPLICANTS_SENTINEL.strip().lower()


def parse_index(text: str) -> List[List[str]]:
    # The sentinel is not applicant data -> no rows.
    if is_no_applicants(text):
        return []
    rows: List[List[str]] = []
    for raw in io.StringIO(text):
        line = raw.rstrip("\n").rstrip("\r")
        if not line.strip():
            continue
        rows.append(next(csv.reader([line])))
    return rows


def _rows_to_csv(rows: List[List[str]]) -> str:
    out = io.StringIO()
    w = csv.writer(out, lineterminator="\n")
    for r in rows:
        w.writerow(r)
    return out.getvalue()


def reconcile_index(dest: paramiko.SFTPClient, dest_path: str,
                    src_text: str, write: bool) -> Tuple[int, str]:
    """
    Reconcile a source daily tracking file into the destination dated CSV.

    The DESTINATION is the authoritative accumulator. The reconciliation always
    checks the destination FIRST and only ever ADDS rows when a dated CSV is
    already present there -- it never overwrites delivered data with a stale
    "No Applicants" source.

    Behaviour:
      * No dated CSV on the destination yet
            -> write the source verbatim (real rows OR the sentinel).
      * Destination dated CSV already exists, source is "No Applicants"
            -> do nothing (keep the real rows already delivered).
      * Destination dated CSV already exists and is just the sentinel,
        source has real rows
            -> replace the sentinel with the source rows.
      * Destination dated CSV already exists with real rows, source has rows
            -> append ONLY the source rows not already present (dedupe).

    When `write` is False nothing is written (dry-run estimate). The destination
    RB folder is never created here; the caller gates `write` on its existence.

    Returns (rows_added, mode).
    """
    src_no_applicants = is_no_applicants(src_text)
    src_rows = [] if src_no_applicants else parse_index(src_text)

    existing_text = _read_text(dest, dest_path)

    # --- destination has no dated CSV yet: seed it verbatim ---------------
    if existing_text is None:
        if write:
            if src_no_applicants:
                _write_text(dest, dest_path,
                            src_text if src_text.endswith("\n") else src_text + "\n")
            else:
                _write_text(dest, dest_path, _rows_to_csv(src_rows))
        if src_no_applicants:
            return 0, "seeded_no_applicants"
        return len(src_rows), "seeded"

    # --- destination already has a dated CSV: APPEND-ONLY -----------------
    # Source says "No Applicants": nothing to add, and we must never clobber
    # the real rows already on the destination. THIS is the case the rewrite
    # fixes -- previously the sentinel overwrote delivered data.
    if src_no_applicants:
        return 0, "kept_existing"

    # Existing destination file is just the sentinel -> replace it with the
    # real rows rather than appending after a "No Applicants" line.
    if is_no_applicants(existing_text):
        if write:
            _write_text(dest, dest_path, _rows_to_csv(src_rows))
        return len(src_rows), "replaced_sentinel"

    # Existing destination file has real rows -> append only the new ones.
    existing_rows = parse_index(existing_text)
    seen = {_row_key(r) for r in existing_rows}
    to_add = [r for r in src_rows if _row_key(r) not in seen]
    if not to_add:
        return 0, "no_change"

    if write:
        text = existing_text
        if text and not text.endswith("\n"):
            text += "\n"
        _write_text(dest, dest_path, text + _rows_to_csv(to_add))
    return len(to_add), "appended"


# =========================================================
# RECURSIVE FOLDER MIRROR
# =========================================================

def mirror_tree(src: paramiko.SFTPClient, dest: paramiko.SFTPClient,
                src_dir: str, dest_dir: str, dest_base: str, tmpdir: str,
                stats: Dict[str, Any]) -> None:
    """
    Recursively copy src_dir -> dest_dir (files + subfolders). Subfolders are
    created only UNDER dest_base (the regulator folder, which must already
    exist); dest_base itself is never created.
    """
    try:
        entries = src.listdir_attr(src_dir)
    except IOError:
        stats["missing_app_folders"].append(src_dir)
        return
    _ensure_dir_under(dest, dest_base, dest_dir)
    for e in entries:
        sp = f"{src_dir}/{e.filename}"
        dp = f"{dest_dir}/{e.filename}"
        if stat.S_ISDIR(e.st_mode):
            mirror_tree(src, dest, sp, dp, dest_base, tmpdir, stats)
        else:
            if SKIP_UNCHANGED:
                try:
                    if dest.stat(dp).st_size == e.st_size:
                        stats["files_skipped"] += 1
                        continue
                except IOError:
                    pass
            local = os.path.join(tmpdir, "xfer.bin")
            src.get(sp, local)
            dest.put(local, dp)
            stats["files_copied"] += 1


# =========================================================
# ENTRYPOINT
# =========================================================

def main(days_back: Optional[int] = None, dry_run: bool = False):
    """
    Scheduled mirror (run every 6 hours). days_back overrides DAYS_BACK;
    dry_run reads + reports what would be mirrored without writing.
    """
    days = days_back or DAYS_BACK
    today = datetime.utcnow()
    window = [today - timedelta(days=i) for i in range(days)]

    src, src_t = _connect(SRC_HOST, SRC_PORT, SRC_USERNAME_VAR, SRC_PASSWORD_VAR,
                          SRC_USERNAME_OVERRIDE, SRC_PASSWORD_OVERRIDE)
    dest, dest_t = _connect(DEST_HOST, DEST_PORT, DEST_USERNAME_VAR, DEST_PASSWORD_VAR,
                            DEST_USERNAME_OVERRIDE, DEST_PASSWORD_OVERRIDE)

    summary: Dict[str, Any] = {
        "dry_run": dry_run,
        "days_window": [d.strftime("%d%m%Y") for d in window],
        "regulators": {},
        "totals": {"files_copied": 0, "files_skipped": 0,
                   "index_rows_added": 0, "app_folders_mirrored": 0},
    }

    try:
        rb_folders = [d for d in _list_dirs(src, SRC_ROOT) if d not in SKIP_FOLDERS]

        with tempfile.TemporaryDirectory() as tmpdir:
            for rb in rb_folders:
                # Resolve destination RB: direct same-name match first, then the
                # fallback map. Unmapped -> skip (no guessing, no folder created).
                dest_rb, resolve_mode = resolve_dest_rb(dest, rb)
                src_rb_dir = f"{SRC_ROOT.rstrip('/')}/{rb}"
                dest_rb_dir = (f"{DEST_ROOT.rstrip('/')}/{dest_rb}"
                               if dest_rb else None)

                rb_stats: Dict[str, Any] = {
                    "dest_folder": dest_rb_dir,
                    "resolve_mode": resolve_mode,
                    "tracking_files": [],
                    "app_ids": [],
                    "files_copied": 0, "files_skipped": 0,
                    "index_rows_added": 0,
                    "missing_app_folders": [],
                }

                if dest_rb is None:
                    # No direct match and no fallback entry. Only report it as an
                    # issue if this source folder actually has tracking data in
                    # the window (otherwise it is just an inactive/unknown folder).
                    has_tracking = any(
                        _read_text(src, f"{src_rb_dir}/{index_filename(rb, d)}") is not None
                        for d in window)
                    if has_tracking:
                        rb_stats["error"] = (f"unmapped source RB '{rb}': no direct "
                                             f"match on destination and no RB_FALLBACK_MAP "
                                             f"entry. Add a mapping to mirror it.")
                        summary["regulators"][rb] = rb_stats
                    continue

                # The destination RB folder MUST already exist -- this script
                # never creates regulator folders. If absent, skip with an error
                # so a stray folder is never created.
                dest_rb_exists = _isdir(dest, dest_rb_dir)
                if not dest_rb_exists:
                    rb_stats["error"] = (f"destination RB folder missing: {dest_rb_dir} "
                                         f"(not created). Check RB_FALLBACK_MAP / DEST_ROOT.")

                # gather application IDs from the last N daily tracking files
                app_ids: List[str] = []
                seen_ids = set()
                for day in window:
                    idx_name = index_filename(rb, day)
                    src_idx = f"{src_rb_dir}/{idx_name}"
                    text = _read_text(src, src_idx)
                    if text is None:
                        continue
                    dest_idx = f"{dest_rb_dir}/{index_filename(dest_rb, day)}"

                    no_applicants = is_no_applicants(text)
                    rows = [] if no_applicants else parse_index(text)
                    rb_stats["tracking_files"].append(
                        {"file": idx_name, "rows": len(rows),
                         "no_applicants": no_applicants})

                    # Reconcile the dated tracking file into the destination.
                    # The destination is the authoritative accumulator: it is
                    # checked FIRST, a "No Applicants" source never overwrites
                    # real rows already delivered, and real rows are appended
                    # (deduped) or replace a stale sentinel. Writes happen only
                    # when the RB folder exists and this is not a dry run.
                    write = (not dry_run) and dest_rb_exists
                    added, mode = reconcile_index(dest, dest_idx, text, write)
                    rb_stats["index_rows_added"] += added

                    # collect application IDs only from real rows
                    for r in rows:
                        if len(r) > APP_ID_COLUMN:
                            aid = r[APP_ID_COLUMN].strip()
                            if aid and aid not in seen_ids:
                                seen_ids.add(aid)
                                app_ids.append(aid)

                if not rb_stats["tracking_files"]:
                    continue  # not an active RB folder this window

                rb_stats["app_ids"] = app_ids

                # mirror each referenced applicant folder (only if RB exists)
                for aid in app_ids:
                    src_app = f"{src_rb_dir}/{aid}"
                    dest_app = f"{dest_rb_dir}/{aid}"
                    if not _isdir(src, src_app):
                        rb_stats["missing_app_folders"].append(src_app)
                        continue
                    if dry_run or not dest_rb_exists:
                        continue
                    tree_stats = {"files_copied": 0, "files_skipped": 0,
                                  "missing_app_folders": []}
                    mirror_tree(src, dest, src_app, dest_app, dest_rb_dir, tmpdir, tree_stats)
                    rb_stats["files_copied"] += tree_stats["files_copied"]
                    rb_stats["files_skipped"] += tree_stats["files_skipped"]
                    rb_stats["missing_app_folders"].extend(tree_stats["missing_app_folders"])

                summary["regulators"][rb] = rb_stats
                summary["totals"]["files_copied"] += rb_stats["files_copied"]
                summary["totals"]["files_skipped"] += rb_stats["files_skipped"]
                summary["totals"]["index_rows_added"] += rb_stats["index_rows_added"]
                summary["totals"]["app_folders_mirrored"] += len(
                    [a for a in app_ids if f"{src_rb_dir}/{a}" not in rb_stats["missing_app_folders"]])

        summary["status"] = "ok"
        summary["regulator_count"] = len(summary["regulators"])
        return summary
    finally:
        for c in (src, dest):
            try:
                c.close()
            except Exception:
                pass
        for t in (src_t, dest_t):
            try:
                t.close()
            except Exception:
                pass