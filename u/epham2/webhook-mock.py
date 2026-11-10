# py: 3.12
# requirements:
# (none; standard library only)

"""
NNAS Webhook Mock  (self-contained Windmill script)
===================================================

A safe receiver for the SMApply webhook. Point the SMApply webhook at this
script instead of the real package builder to see exactly what arrives and
how the real script would interpret it.

It has NO side effects:
  - no SharePoint calls, no SFTP connection, no file writes
  - no Windmill variables or secrets are read

It returns a summary of:
  - which top-level webhook fields arrived
  - the ApplicationID and NNAS ID the real script would use
  - whether additional_content parsed, and which sections it contains
  - the regulators, abbreviations, provinces and nursing groups resolved
  - which applicant fields are present or missing (values are masked)
  - warnings for anything that would cause the real run to skip or fail

PII note: values are masked in the result by default. Windmill still stores
the full job input (the raw webhook body) in the run history, which is where
you can copy the complete additional_content from if you need to replay it
into the real script. Keep INCLUDE_RAW_PAYLOAD = False unless you are
debugging in a restricted workspace.

The RB_ABBREV / RB_PROVINCE_MAP tables below are copied from
nnas_sftp_repush.py. Keep them in sync if you edit either script.
"""

import json
import re
from typing import Any, Dict, List, Optional, Tuple

INCLUDE_RAW_PAYLOAD = False

KNOWN_PROVIDERS = ("tm", "ces")
REPORT_DESIGNATIONS = ("RN", "LPN", "RPN")


# =========================================================
# REGULATOR TABLES  (copied from nnas_sftp_repush.py)
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
    "NL College of Nurses":                                              "Newfoundland and Labrador",
    "BC College of Nurses and Midwives":                                 "British Columbia",
    "Nova Scotia College of Nursing":                                    "Nova Scotia",
    "PEI College of Nursing & Midwifery":                                "Prince Edward Island",
    "College of Licensed Practical Nurses and Health Care Aides of Alberta": "Alberta",
}

RB_ABBREV = {
    "British Columbia College of Nurses and Midwives":                   "bccnm",
    "College Of Nurses Of Ontario":                                      "cno",
    "College of Registered Nurses Of Alberta":                           "crna",
    "College Of Registered Nurses Of Manitoba":                          "crnm",
    "College of Registered Nurses of Saskatchewan":                      "crns",
    "College of Nursing of New Brunswick":                               "cnnb",
    "College of Registered Nurses of Newfoundland and Labrador":         "nlcn",
    "College of Registered Nurses and Midwives of Prince Edward Island": "peicnm",
    "Nova Scotia College of Nursing - RN":                               "nscn",
    "Association Of New Brunswick Licensed Practical Nurses":            "anblpn",
    "College Of Licensed Practical Nurses Of Alberta":                   "clha",
    "College Of Licensed Practical Nurses Of Manitoba":                  "clpnm",
    "College of Licensed Practical Nurses of Saskatchewan":              "clpns",
    "College Of Licensed Practical Nurses Of Newfoundland And Labrador": "nlcn",
    "College of Licensed Practical Nurses of Prince Edward Island":      "peicnm",
    "Nova Scotia College of Nursing - LPN":                              "nscn",
    "College of Registered Psychiatric Nurses of Alberta":               "crpna",
    "College of Registered Psychiatric Nurses of Manitoba":              "crpnm",
    "College of Registered Psychiatric Nurses of Saskatchewan":          "crpns",
    "NL College of Nurses":                                              "nlcn",
    "BC College of Nurses and Midwives":                                 "bccnm",
    "Nova Scotia College of Nursing":                                    "nscn",
    "PEI College of Nursing & Midwifery":                                "peicnm",
    "College of Licensed Practical Nurses and Health Care Aides of Alberta": "clha",
}

_KNOWN_ABBREVS = sorted(set(RB_ABBREV.values()), key=len, reverse=True)


def _lookup(table: Dict[str, str], name: str) -> str:
    if name in table:
        return table[name]
    low = {k.lower(): v for k, v in table.items()}
    return low.get((name or "").strip().lower(), "")


def canonical_rb_for_abbrev(abbrev: str) -> str:
    ab = (abbrev or "").strip().lower()
    return next((n for n, a in RB_ABBREV.items() if a == ab), "")


def group_for_rb_name(rb_name: str) -> str:
    low = (rb_name or "").lower()
    if "psychiatric" in low:
        return "RPN"
    if "practical" in low:
        return "LPN"
    return "RN"


# =========================================================
# VALUE HELPERS  (same behavior as the real script)
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


def first_nonempty(*vals: str) -> str:
    for v in vals:
        if v and str(v).strip():
            return str(v).strip()
    return ""


def mask(value: Any) -> str:
    """Report presence only, never the value."""
    s = safe_str(value)
    return f"present ({len(s)} chars)" if s else "missing"


def parse_additional_content(additional_content: Any) -> Tuple[Dict[str, Any], Dict[str, Any]]:
    dbg: Dict[str, Any] = {"present": bool(additional_content),
                           "type": type(additional_content).__name__}
    if not additional_content:
        return {}, {**dbg, "parsed": False}
    if isinstance(additional_content, dict):
        return additional_content, {**dbg, "parsed": True, "note": "arrived as an object"}
    if not isinstance(additional_content, str):
        return {}, {**dbg, "parsed": False, "error": "unexpected type"}
    s = additional_content.strip()
    try:
        raw = json.loads(s)
        if isinstance(raw, str):
            raw = json.loads(raw)
            dbg["note"] = "double-encoded JSON string"
    except json.JSONDecodeError as e:
        if (s.startswith('"') and s.endswith('"')) or (s.startswith("'") and s.endswith("'")):
            try:
                raw = json.loads(s[1:-1])
                if isinstance(raw, str):
                    raw = json.loads(raw)
                dbg["note"] = "wrapped in extra quotes"
            except json.JSONDecodeError as e2:
                return {}, {**dbg, "parsed": False, "error": f"json decode failed: {e2}"}
        else:
            return {}, {**dbg, "parsed": False, "error": f"json decode failed: {e}"}
    if not isinstance(raw, dict):
        return {}, {**dbg, "parsed": False, "error": f"not an object: {type(raw).__name__}"}
    return raw, {**dbg, "parsed": True}


# =========================================================
# ENTRYPOINT  (same signature as nnas_sftp_repush.main)
# =========================================================

def main(reference_id: Optional[str] = None,
         additional_content: Optional[str] = None,
         user_id: Any = None,
         regulators: Optional[List[str]] = None,
         **kwargs: Any):
    warnings: List[str] = []

    # ---- what arrived ----
    received = {
        "reference_id": mask(reference_id) if reference_id else "missing",
        "user_id": "present" if user_id not in (None, "") else "missing",
        "additional_content": "present" if additional_content else "missing",
        "regulators": regulators if regulators else "not sent",
        "other_top_level_keys": sorted(kwargs.keys()),
    }

    payload, ac_dbg = parse_additional_content(additional_content)
    if not ac_dbg.get("parsed"):
        warnings.append("additional_content did not parse; the real script would fall "
                        "back to the reports folder / PDF for regulators.")

    # ---- identifiers (the real script's precedence) ----
    application_id = first_nonempty(
        reference_id, kwargs.get("title"), kwargs.get("application_title"),
        safe_str(payload.get("reference_id")),
        str(kwargs["appid"]) if kwargs.get("appid") is not None else "",
    )
    application_id_source = (
        "reference_id" if reference_id else
        "title" if kwargs.get("title") else
        "application_title" if kwargs.get("application_title") else
        "additional_content.reference_id" if safe_str(payload.get("reference_id")) else
        "appid" if kwargs.get("appid") is not None else "none"
    )
    nnas_id = first_nonempty(
        str(user_id) if user_id is not None else "",
        str(kwargs["nnas_id"]) if kwargs.get("nnas_id") is not None else "",
        safe_str(payload.get("nnas_id")),
    )
    if not application_id:
        warnings.append("No ApplicationID found; the real script would stop.")
    elif not re.match(r"^NNAS-", application_id, re.IGNORECASE):
        warnings.append(f"ApplicationID '{application_id}' does not start with 'NNAS-'; "
                        "check this matches the SharePoint folder name.")
    if not nnas_id:
        warnings.append("No NNAS ID (user_id) found; the real script would fall back "
                        "to the ApplicationID for file names.")

    # ---- application_info ----
    ai = payload.get("application_info") or {}
    provider = safe_str(ai.get("provider")).lower()
    if provider and provider not in KNOWN_PROVIDERS:
        warnings.append(f"Provider '{provider}' is not tm/ces (informational only; the "
                        "real script detects the provider from SharePoint).")

    # ---- regulators ----
    rn_list  = [x.strip() for x in safe_str(ai.get("regulatory_body_rn")).split(",") if x.strip()]
    lpn_list = [x.strip() for x in safe_str(ai.get("regulatory_body_lpn")).split(",") if x.strip()]

    resolved: List[Dict[str, Any]] = []
    seen = set()

    def add_rb(name: str, group: str, source: str):
        if name.lower() in seen:
            return
        seen.add(name.lower())
        abbrev = _lookup(RB_ABBREV, name)
        province = _lookup(RB_PROVINCE_MAP, name)
        resolved.append({
            "name": name, "source": source, "group": group,
            "abbrev": abbrev or None,
            "sftp_folder": f"/REGULATORY/{abbrev.upper()}/{application_id}/" if abbrev else None,
            "province": province or None,
        })
        if not abbrev:
            warnings.append(f"Regulator '{name}' has no abbreviation mapping; the real "
                            "script would SKIP it. Add it to RB_ABBREV.")
        elif not province:
            warnings.append(f"Regulator '{name}' has no province mapping; RBProvince "
                            "would be blank. Add it to RB_PROVINCE_MAP.")

    for name in rn_list:
        add_rb(name, "RN", "additional_content (rn)")
    for name in lpn_list:
        add_rb(name, "LPN", "additional_content (lpn)")

    if not resolved and regulators:
        for raw in regulators:
            txt, group = str(raw or "").strip(), ""
            if ":" in txt and txt.rsplit(":", 1)[1].strip().upper() in REPORT_DESIGNATIONS:
                txt, group = (p.strip() for p in txt.rsplit(":", 1))
                group = group.upper()
            if not txt:
                continue
            name = canonical_rb_for_abbrev(txt) if txt.lower() in _KNOWN_ABBREVS else txt
            add_rb(name, group or group_for_rb_name(name), "regulators input")

    if not resolved:
        warnings.append("No regulators resolved from the webhook; the real script would "
                        "derive them from the '_latest' files in the credential reports folder.")

    # ---- applicant field coverage (masked) ----
    pi = payload.get("personal_info") or {}
    reg = payload.get("registration") or {}
    e1 = reg.get("entry_1") or {}
    personal_fields = ["sa_first_name", "sa_middle_name", "sa_last_name", "sa_other_names",
                       "sa_dob", "sa_email", "sa_phone", "sa_alt_phone",
                       "sa_current_address_line1", "sa_current_city", "sa_current_province",
                       "sa_current_country", "sa_current_postal_code",
                       "sa_mailing_same_as_current_flag", "sa_initial_education_country",
                       "sa_first_language"]
    mailing_fields = ["sa_mailing_address_line1", "sa_mailing_city", "sa_mailing_province",
                      "sa_mailing_country", "sa_mailing_postal_code"]
    registration_fields = ["sa_registration_title_1", "sa_registration_number_1",
                           "sa_registration_status_1", "sa_registration_issued_date_1",
                           "sa_registration_expiry_date_1", "sa_registrant_name_issued_1",
                           "sa_jurisdiction_name_1", "sa_jurisdiction_address_1",
                           "sa_jurisdiction_country_1", "sa_registration_method_1",
                           "sa_registration_lifetime_flag_1", "sa_registration_discipline_flag_1"]

    mailing_same = safe_str(pi.get("sa_mailing_same_as_current_flag")).lower().startswith("y")
    coverage = {
        "personal_info": {f: mask(pi.get(f)) for f in personal_fields},
        "mailing_address": ("same as current (not needed)" if mailing_same
                            else {f: mask(pi.get(f)) for f in mailing_fields}),
        "registration_flags": {
            f: mask(reg.get(f)) for f in (
                "sa_eligible_to_practice_in_initial_edu_jurisdiction_flag",
                "sa_diploma_grants_right_to_practice_flag")},
        "registration_entry_1": {f: mask(e1.get(f)) for f in registration_fields},
    }

    result: Dict[str, Any] = {
        "status": "ok" if not warnings else "ok_with_warnings",
        "mock": True,
        "note": "No SharePoint or SFTP calls were made.",
        "received": received,
        "resolved": {
            "application_id": application_id or None,
            "application_id_source": application_id_source,
            "nnas_id": nnas_id or None,
            "payload_provider": provider or None,
            "nursing_group_selection": safe_str(ai.get("nursing_group_selection")) or None,
            "regulators": resolved,
        },
        "additional_content": {
            **ac_dbg,
            "sections": sorted(payload.keys()) if payload else [],
            "application_info_keys": sorted(ai.keys()) if isinstance(ai, dict) else [],
        },
        "field_coverage": coverage,
        "warnings": warnings,
    }
    if INCLUDE_RAW_PAYLOAD:
        result["raw"] = {"reference_id": reference_id, "user_id": user_id,
                         "additional_content": payload, "other": kwargs}
    return result