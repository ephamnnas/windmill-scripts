# py: 3.12
# requirements:
# wmill
# requests

"""
NNAS RB Portal Record Upsert (Windmill)
=======================================

Lifecycle-stage-callable step. On submission it creates/updates ONE RB Portal
record PER REGULATORY BODY, all sharing the same application_id (e.g.
NNAS-T-0000000030), distinguished by regulatory_body. This is intentionally
different from the scraper (which collapses to one record); here we want a
row per RB.

Per-record fields set this stage:
    status                = "Waiting For Required Items"  (constant)
    dob                   = SMApply tasks "Date of birth" field
    order_date            = date of THIS run (today)
    order_expiration_date = order_date + 2 years
    country               = additional_content.personal_info
                            .sa_initial_education_country
    regulatory_body= the RB name (distinguishes the row + composite key)
Keys:
    application_id = title (same on every row)   nnas_id = user_id

REGULATORS: union of additional_content.application_info
.regulatory_body_rn and
regulatory_body_lpn (comma-split, trimmed, de-duped, order preserved) -- exactly
what the SFTP package builder uses.

MODE:
    "per_rb_upsert" (default) -- match each row on application_id+regulatory_body
        so re-runs at later stages UPDATE the right row (no duplicate explosion).
        Every row still shares the same application_id.
    "append" -- literally create a new row every call (raw true-duplicates;
        WILL multiply rows on every re-run -- use only if you never re-run).

Requires RBP Upload API plugin v1.5.0 (composite lookup + opt-in flags).
Worker egress IP must be on the plugin allowlist (the scraper already posts here).
"""

import re
import json
from datetime import date
from typing import Any, Dict, List, Optional
from urllib.parse import quote

import wmill
import requests


# ====================================
# CONFIG  -- point these at your existing values
# ====================================
SMAPPLY_API_BASE      = "https://portal.nnas.ca/api"
SMAPPLY_API_TOKEN_VAR = "u/epham2/SMAPPLY_ACCESS_TOKEN"

RBP_API_BASE     = "https://rbportal.nnas.ca/wp-json/rbp/v1"
RBP_WP_USER_VAR  = "u/knguyen/RBP_WP_USER"
RBP_APP_PASS_VAR = "u/knguyen/RBP_APP_PASS"

STAGE_STATUS = "Waiting For Required Items"
HTTP_TIMEOUT = 30


# ====================================
# HELPERS
# ====================================
def parse_additional_content(ac: Optional[Any]) -> Dict[str, Any]:
    if not ac:
        return {}
    if isinstance(ac, dict):
        return ac
    if not isinstance(ac, str):
        return {}
    s = ac.strip()
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
            except json.JSONDecodeError:
                return {}
        else:
            return {}
    return raw if isinstance(raw, dict) else {}


def resolve_additional_content(named: Optional[Any], kwargs: Dict[str, Any]) -> Optional[Any]:
    """Use the named param if present, else recover it from **kwargs.
    The dry-run blank-country bug was Windmill not binding the named param;
    pulling it from kwargs as a fallback makes the step resilient to that."""
    if named:
        return named
    for k in ("additional_content", "additionalContent", "additional_content_blob"):
        if kwargs.get(k):
            return kwargs[k]
    return named


def normalize_date(v: Any) -> str:
    v = ("" if v is None else str(v)).strip()
    if not v:
        return ""
    m = re.match(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})", v)
    if m:
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    m = re.match(r"^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$", v)
    if m:
        return f"{m.group(3)}-{int(m.group(2)):02d}-{int(m.group(1)):02d}"
    return v


def clean_name(s: Any) -> str:
    """Collapse whitespace and trim for a display name."""
    return re.sub(r"\s+", " ", str(s or "")).strip()


def smapply_get_all_tasks(internal_id: str, token: str) -> List[Dict[str, Any]]:
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    url = f"{SMAPPLY_API_BASE}/applications/{internal_id}/tasks/"
    out: List[Dict[str, Any]] = []
    for _ in range(20):
        r = requests.get(url, headers=headers, timeout=HTTP_TIMEOUT)
        r.raise_for_status()
        data = r.json()
        out.extend(data.get("results", []))
        nxt = data.get("next")
        if not nxt:
            break
        url = nxt.replace("http://", "https://", 1)
    return out


def smapply_get_application(internal_id: str, token: str) -> Dict[str, Any]:
    """Best-effort fetch of the application detail (for the applicant name).
    Returns {} on any failure so a missing name never breaks the run."""
    try:
        headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
        url = f"{SMAPPLY_API_BASE}/applications/{internal_id}/"
        r = requests.get(url, headers=headers, timeout=HTTP_TIMEOUT)
        if not r.ok:
            return {}
        data = r.json()
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def ien_from_applicant(detail: Dict[str, Any]) -> str:
    """first_name + last_name from an application detail's applicant object."""
    applicant = detail.get("applicant") or {}
    first = clean_name(applicant.get("first_name"))
    last  = clean_name(applicant.get("last_name"))
    return clean_name(f"{first} {last}")


def dob_from_tasks(tasks: List[Dict[str, Any]]) -> str:
    def scan(task: Dict[str, Any]) -> str:
        for v in (task.get("data") or {}).values():
            if isinstance(v, dict):
                if str(v.get("label", "")).strip().lower().startswith("date of birth"):
                    return normalize_date(v.get("response"))
        return ""
    for t in tasks:
        if str(t.get("name", "")).strip().lower() == "personal information":
            d = scan(t)
            if d:
                return d
    for t in tasks:
        d = scan(t)
        if d:
            return d
    return ""


def plus_two_years(d: date) -> date:
    """d + 2 years, with a Feb-29 -> Feb-28 guard."""
    try:
        return d.replace(year=d.year + 2)
    except ValueError:
        return d.replace(year=d.year + 2, day=28)


# ------------------------------------
# Regulatory-body alias -> (portal tag, canonical name)
# ------------------------------------
# SMApply sends display aliases (e.g. "NL College of Nurses"). The portal
# surfaces a record to an RB user only when BOTH hold:
#   (a) the record's rb_recipient term (from all_recipients_raw) intersects the
#       user's tags, AND
#   (b) the record's regulatory_body matches one of the user's mapped names
#       (rb-portal.php tag_to_regulatory_bodies()).
# So each per-RB record must carry the correct TAG (recipient) AND a canonical
# regulatory_body NAME that the portal maps that tag to. Every (tag, name) below
# was verified present in tag_to_regulatory_bodies().
#
# Shared-name colleges (NL, PEI, NS, BC) resolve by DISCIPLINE -- the list the
# RB came from (regulatory_body_rn vs regulatory_body_lpn) -- giving a distinct
# tag and/or name per discipline. That also keeps the composite key
# (application_id + regulatory_body) unique if an applicant lists the same
# college under both disciplines.
#
# Single-discipline entries ignore the discipline argument.

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
    "pei college of nursing & midwifery": {
        "rn":  ("CRNPEI",  "College Of Registered Nurses Of Prince Edward Island"),
        "lpn": ("CLPNPEI", "College of Licensed Practical Nurses of Prince Edward Island"),
    },
    "pei college of nursing and midwifery": {
        "rn":  ("CRNPEI",  "College Of Registered Nurses Of Prince Edward Island"),
        "lpn": ("CLPNPEI", "College of Licensed Practical Nurses of Prince Edward Island"),
    },
    "nova scotia college of nursing": {
        "rn":  ("NSCN", "Nova Scotia College of Nursing-RN"),
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


def resolve_rb(name: str, discipline: str):
    """(alias, discipline) -> (tag, canonical_name) or (None, name) if unmapped."""
    key = (name or "").strip().lower()
    key = re.sub(r"\s+", " ", key)
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


def regulators_from_ac(ac: Dict[str, Any]) -> List[tuple]:
    """SFTP-style union as (name, discipline), de-duped on (name, discipline)."""
    ai = ac.get("application_info") or {}
    seen, out = set(), []
    for field, disc in (("regulatory_body_rn", "rn"), ("regulatory_body_lpn", "lpn")):
        for rb in str(ai.get(field) or "").split(","):
            rb = rb.strip()
            if rb and (rb.lower(), disc) not in seen:
                seen.add((rb.lower(), disc))
                out.append((rb, disc))
    return out


def post_rb_record(application_id: str, body: Dict[str, Any],
                   wp_user: str, app_pass: str) -> Dict[str, Any]:
    url = f"{RBP_API_BASE}/record/{quote(application_id, safe='')}"
    r = requests.post(url, json=body, auth=(wp_user, app_pass), timeout=HTTP_TIMEOUT)
    try:
        detail = r.json()
    except ValueError:
        detail = r.text[:500]
    if not r.ok:
        raise RuntimeError(f"RB portal upsert failed [{r.status_code}]: {detail}")
    return detail if isinstance(detail, dict) else {"raw": detail}


# ====================================
# ENTRYPOINT
# ====================================
def main(
    id: Any = None,                          # SMApply INTERNAL id (35367271)
    title: Optional[str] = None,             # ApplicationID (NNAS-T-0000000030)
    user_id: Any = None,                     # NNAS ID (131701251)
    additional_content: Optional[Any] = None,
    reference_id: Optional[str] = None,
    ien_name: Optional[str] = None,          # applicant name (optional override)
    mode: str = "per_rb_upsert",             # "per_rb_upsert" | "append"
    dry_run: bool = False,
    **kwargs: Any,
):
    application_id = (str(title) if title else "") or (str(reference_id) if reference_id else "")
    nnas_id        = str(user_id) if user_id is not None else ""
    internal_id    = str(id) if id is not None else (str(kwargs.get("id")) if kwargs.get("id") is not None else "")

    if not application_id:
        return {"status": "error", "message": "Missing title/reference_id (ApplicationID)."}
    if not internal_id:
        return {"status": "error", "message": "Missing internal SMApply id for the tasks call."}
    if mode not in ("per_rb_upsert", "append"):
        return {"status": "error", "message": f"Invalid mode '{mode}'."}

    smapply_token = wmill.get_variable(SMAPPLY_API_TOKEN_VAR)
    wp_user       = wmill.get_variable(RBP_WP_USER_VAR)
    app_pass      = wmill.get_variable(RBP_APP_PASS_VAR)
    missing = [n for n, v in ((SMAPPLY_API_TOKEN_VAR, smapply_token),
                              (RBP_WP_USER_VAR, wp_user),
                              (RBP_APP_PASS_VAR, app_pass)) if not v]
    if missing:
        return {"status": "error", "message": "Missing Windmill variables: " + ", ".join(missing)}

    # ---- gather shared data ----
    ac_raw = resolve_additional_content(additional_content, kwargs)
    ac     = parse_additional_content(ac_raw)
    # If the WHOLE webhook envelope was passed as additional_content (a common
    # mapping mistake), unwrap the inner blob. The envelope has an
    # "additional_content" key but no application_info/personal_info at top level.
    for _ in range(3):
        if isinstance(ac, dict) and "additional_content" in ac \
           and not (ac.get("application_info") or ac.get("personal_info")):
            ac = parse_additional_content(ac.get("additional_content"))
        else:
            break
    pi     = ac.get("personal_info") or {}

    tasks      = smapply_get_all_tasks(internal_id, smapply_token)
    dob        = dob_from_tasks(tasks) or normalize_date(pi.get("sa_dob"))
    run_date   = date.today()
    order_date = run_date.isoformat()
    order_expiration_date = plus_two_years(run_date).isoformat()
    country    = str(pi.get("sa_initial_education_country") or "").strip()

    # ---- IEN name ----
    # Webhook/payload value wins (saves an API call); else applicant first+last
    # from personal_info; else fetched from the SMApply application detail. The
    # detail fetch is best-effort so a missing name never breaks the run.
    ien = clean_name(
        ien_name
        or kwargs.get("ien_name")
        or kwargs.get("applicant_name")
        or kwargs.get("name")
    )
    ien_source = "input" if ien else ""
    if not ien:
        f = clean_name(kwargs.get("first_name") or pi.get("sa_first_name"))
        l = clean_name(kwargs.get("last_name") or pi.get("sa_last_name"))
        ien = clean_name(f"{f} {l}")
        if ien:
            ien_source = "payload_name"
    if not ien:
        ien = ien_from_applicant(smapply_get_application(internal_id, smapply_token))
        ien_source = "smapply" if ien else "unresolved"

    # Resolve each (alias, discipline) -> (tag, canonical regulatory_body).
    # De-dup on the canonical regulatory_body so one row per distinct college/
    # discipline. Unmapped names still get a record (admin-visible) and are
    # surfaced in debug so they can be added to the alias map.
    resolved, seen_rb, unmapped = [], set(), []
    for name, disc in regulators_from_ac(ac):
        tag, canonical = resolve_rb(name, disc)
        if canonical.lower() in seen_rb:
            continue
        seen_rb.add(canonical.lower())
        resolved.append({"alias": name, "discipline": disc,
                         "tag": tag, "regulatory_body": canonical})
        if tag is None:
            unmapped.append(name)

    # Full recipient set for the application (every resolved tag), listed on
    # EVERY record -- matches the scraper's "All Recipients" behavior. Per-user
    # visibility is still correct: the portal also requires the strict
    # regulatory_body match, so each user only sees the row whose regulatory_body
    # maps to their tag, even though all rows share the same recipient list.
    all_tags = " ".join(dict.fromkeys(r["tag"] for r in resolved if r["tag"]))

    shared = {
        "nnas_id":               nnas_id,
        "ien_name":              ien,
        "status":                STAGE_STATUS,
        "order_date":            order_date,
        "order_expiration_date": order_expiration_date,
        "dob":                   dob,
        "country":               country,
    }
    shared = {k: v for k, v in shared.items() if v != ""}

    debug = {
        "internal_id": internal_id,
        "task_count": len(tasks),
        "ac_present": bool(ac),                       # <-- shows if additional_content arrived
        "ien_name": ien,
        "ien_name_source": ien_source,
        "resolved_regulators": resolved,
        "unmapped_regulators": unmapped,             # add these to the alias map if non-empty
        "all_recipients": all_tags,
        "resolved": {"order_date": order_date,
                     "order_expiration_date": order_expiration_date,
                     "dob": dob, "country": country},
        "mode": mode,
    }

    # ---- build one body per resolved regulator (fall back to RB-less record) ----
    def build_body(rb: Optional[Dict[str, Any]]) -> Dict[str, Any]:
        b = {"application_id": application_id, **shared}
        if all_tags:
            b["all_recipients_raw"] = all_tags          # full recipient set on every row
        if rb:
            b["regulatory_body"] = rb["regulatory_body"]   # canonical -> strict match routes visibility
            disc_label = discipline_label(rb.get("discipline"), rb["regulatory_body"])
            if disc_label:
                b["discipline"] = disc_label               # per-RB: RN / LPN / Psychiatric
            if mode == "per_rb_upsert":
                b["_match_by_regulatory_body"] = True       # composite addressing (plugin v1.5.0)
        if mode == "append":
            b["_force_create"] = True                       # always insert a new row
        return b

    targets = resolved if resolved else [None]
    bodies = [build_body(rb) for rb in targets]

    if dry_run:
        return {"status": "dry_run", "application_id": application_id,
                "record_count": len(bodies), "would_post": bodies, "debug": debug}

    results, errors = [], []
    for rb, body in zip(targets, bodies):
        try:
            rb_name = rb["regulatory_body"] if rb else None
            results.append({"regulatory_body": rb_name, "response": post_rb_record(application_id, body, wp_user, app_pass)})
        except Exception as e:
            rb_name = rb["regulatory_body"] if rb else None
            errors.append({"regulatory_body": rb_name, "error": str(e)})

    return {
        "status": "ok" if not errors else "partial",
        "application_id": application_id,
        "records_posted": len(results),
        "records_failed": len(errors),
        "results": results,
        "errors": errors,
        "debug": debug,
    }