# requirements:
# requests
# wmill

"""
Windmill script - NNAS SMApply onboarding (program 295185)

Drop-in replacement for the onboard_applicant runnable used by the
Invite-NB-Applicant app. Same arguments the app sends, same result keys the
app reads.

Flow:
  1) POST /api/users/ with the applicant role and two applicant custom
     fields: Pilot Program (153504, "NB-E"/"NB-F" from language) and
     applicant_email (165959). An existing address is reused instead.
  2) Read the user back and confirm both custom fields persisted.
  3) PUT /api/users/{id}/eligibility/ with the six-question form.
  4) POST /api/applications/ (program 295185, stage 1361209), reusing an
     existing application on the program if there is one.
  5) PATCH the "NB Pilot" label onto the application, confirmed on readback.
  6) PATCH task 2048019 with profession and regulatory body, then a
     status-only PATCH to mark it complete.
  7) If send_invite is set, POST the applicant details to the Power Automate
     instant cloud flow, which sends the invitation email. One call, no
     polling.

Removed in this version because they did not work or only added time:
the signed_up/email_verified and email_subscribed fields on create (both
rejected by this tenant, costing a failed POST each run), the 10 second
settle delay, every guessed invite route, and the invited_at polling.

Runs on a self-hosted Windmill instance. SMApply calls use the bearer token in
u/epham2/SMAPPLY_ACCESS_TOKEN. Errors surface the response body explicitly.
"""

import wmill
import requests
import time
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional, Tuple, Literal

# =========================================================
# CONSTANTS
# =========================================================
SMAPPLY_BASE = "https://nnas.smapply.io/api"
SMAPPLY_TOKEN_VAR = "u/epham2/SMAPPLY_ACCESS_TOKEN"

# ---------------------------------------------------------
# Invite flow (Power Automate, "When an HTTP request is received").
#
# Trigger is set to "Anyone", so the URL's own signature lets a request in.
# On top of that, every call carries a shared secret in the X-NNAS-Flow-Key
# header, and the flow's first step rejects any request without the right
# value. A leaked URL alone therefore cannot send an email.
#
# Both values live in Windmill secret variables, never in this file.
# ---------------------------------------------------------
INVITE_FLOW_URL_VAR = "u/knguyen/NB_INVITE_FLOW_URL"
INVITE_FLOW_KEY_VAR = "u/knguyen/NB_INVITE_FLOW_KEY"
INVITE_FLOW_KEY_HEADER = "X-NNAS-Flow-Key"
INVITE_FLOW_TIMEOUT = 60  # seconds; the flow responds after sending

PROGRAM_ID = 295185
INITIAL_STAGE_ID = 1361209      # "Application Stage"
DISCIPLINE_TASK_ID = 2048019    # "Select Nursing Discipline"

PILOT_PROGRAM_CUSTOM_FIELD_ID = 153504     # "Pilot Program"
APPLICANT_EMAIL_CUSTOM_FIELD_ID = 165959   # "applicant_email"

SMAPPLY_STATUS_COMPLETED = 3

NB_PILOT_LABEL_ID = 183165
NB_PILOT_LABEL_NAME = "NB Pilot"

# Only used when the application POST says "not eligible"; the normal path
# never waits.
ELIGIBILITY_RETRIES = 5
ELIGIBILITY_RETRY_DELAY = 3  # seconds

Q_PROFESSION = "7D4FvVGWuh"   # RN -> 0, LPN -> 1
Q_RB_RN      = "faO2qkS4PX"
Q_RB_LPN     = "IWgvFrNZHX"

PROFESSION_RESPONSE = {"RN": 0, "LPN": 1}
PILOT_PROGRAM_BY_LANGUAGE = {"English": "NB-E", "French": "NB-F"}

# ---------------------------------------------------------
# Eligibility form (PUT /api/users/{id}/eligibility/). Values adopted from
# the 295185 reference account. GET returns {"label","response"} wrappers;
# PUT takes flat {question_id: value}.
# ---------------------------------------------------------
ELIG_Q_CONNECTED = "8LpmN6CBVM"
ELIG_Q_TERMS     = "A3fGSIXu9p"
ELIG_Q_FULL_NAME = "C1cK1aI9ni"
ELIG_Q_DOCS      = "Db2U6y1U0Q"
ELIG_Q_PRIOR     = "FFhD4gHD9J"
ELIG_Q_APPLY     = "qzzA39MQAV"

ELIG_CONNECTED_RESPONSE = 1
ELIG_DOCS_RESPONSE = None
ELIG_PRIOR_RESPONSE = 1
ELIG_APPLY_RESPONSE = 0

ELIG_TERMS_SIGNATURE = (
    "/protected/media-file/FkmBTc_C-GJ6iD28zx8OmFNuhZEUCtxmqgzNwaYIKWthTpR3JG5U"
    "DkAurInQBnd2u5FwYgql4crOJ9te2NzMZ_pTYbQaJmQErqQKHxKnNU6eTWkyGGIx19L5pP6c8A"
    "4Akm-_EJ1JGcecXsuKDHjE1WvMWnPY9ctiUjl0aTdNCCNf5lKEFuOZhSj-jXJFpRJPzPcvTsKT"
    "bZRsykgUvJeTmFkPTdoc2m97708s2xQDbjnLOftxs4-7OfUQMI-vfpfE__Qf0xmuratkGf24M5"
    "zjA9k9UOFl7znx2_LWBfJ6Kv6AlB1QsYEouRhZcqgNmi1uXfHg_WbPybWj6bRXBw1fVnKr0jCG"
    "UikqphrZ2L-1FHU=/2026-06-15-18-05-16-signature.svg"
)

# ---------------------------------------------------------
# Regulatory-body registry. Indices read from the live task definition and
# verified. LPN entries whose label also appears in the RN list carry a
# "|lpn" suffix; resolve_body handles it, so callers pass the plain name.
# ---------------------------------------------------------
RB_REGISTRY: Dict[str, Dict[str, Any]] = {
    # RN dropdown (faO2qkS4PX)
    "nl college of nurses": {"discipline": "RN", "province": "NL", "index": 0},
    "pei college of nursing & midwifery": {"discipline": "RN", "province": "PE", "index": 1},
    "college of nursing of new brunswick": {"discipline": "RN", "province": "NB", "index": 2},
    "nova scotia college of nursing": {"discipline": "RN", "province": "NS", "index": 3},
    "college of nurses of ontario": {"discipline": "RN", "province": "ON", "index": 4},
    "college of registered nurses of manitoba": {"discipline": "RN", "province": "MB", "index": 5},
    "college of registered nurses of saskatchewan": {"discipline": "RN", "province": "SK", "index": 6},
    "college of registered nurses of alberta": {"discipline": "RN", "province": "AB", "index": 7},
    "bc college of nurses and midwives": {"discipline": "RN", "province": "BC", "index": 8},
    # LPN dropdown (IWgvFrNZHX)
    "nl college of nurses|lpn": {"discipline": "LPN", "province": "NL", "index": 0},
    "pei college of nursing & midwifery|lpn": {"discipline": "LPN", "province": "PE", "index": 1},
    "association of new brunswick licensed practical nurses": {"discipline": "LPN", "province": "NB", "index": 2},
    "nova scotia college of nursing|lpn": {"discipline": "LPN", "province": "NS", "index": 3},
    "college of nurses of ontario|lpn": {"discipline": "LPN", "province": "ON", "index": 4},
    "college of licensed practical nurses of manitoba": {"discipline": "LPN", "province": "MB", "index": 5},
    "college of licensed practical nurses of saskatchewan": {"discipline": "LPN", "province": "SK", "index": 6},
    "college of licensed practical nurses and health care aides of alberta": {"discipline": "LPN", "province": "AB", "index": 7},
    "bc college of nurses and midwives|lpn": {"discipline": "LPN", "province": "BC", "index": 8},
}


def _rb_key(name: str) -> str:
    return " ".join((name or "").lower().split())


def resolve_body(regulatory_body: str, regulator: str) -> Dict[str, Any]:
    base = _rb_key(regulatory_body)
    candidates = [base, f"{base}|lpn"] if regulator == "LPN" else [base]
    for key in candidates:
        entry = RB_REGISTRY.get(key)
        if entry and entry["discipline"] == regulator:
            return entry
    entry = RB_REGISTRY.get(base)
    if entry is None:
        raise ValueError(f"No SMApply mapping for regulatory body '{regulatory_body}'.")
    raise ValueError(
        f"Regulator '{regulator}' does not match '{regulatory_body}', which is "
        f"a {entry['discipline']} body."
    )


# =========================================================
# HTTP
# =========================================================
def _headers(automations: bool = False) -> Dict[str, str]:
    token = wmill.get_variable(SMAPPLY_TOKEN_VAR)
    h = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    if automations:
        h["apply-automations"] = "full"
    return h


def _request(method: str, url: str, **kwargs) -> Dict[str, Any]:
    resp = requests.request(method, url, timeout=60, **kwargs)
    if not resp.ok:
        try:
            body = resp.json()
        except Exception:
            body = {"text": resp.text}
        raise RuntimeError(f"{method} {url} -> {resp.status_code}: {body}")
    try:
        return resp.json()
    except Exception:
        return {"text": resp.text}


def _safe_get(url: str) -> Dict[str, Any]:
    try:
        resp = requests.get(url, headers=_headers(), timeout=30)
        try:
            body = resp.json()
        except Exception:
            body = {"text": resp.text}
        return {"url": url, "status": resp.status_code, "body": body}
    except Exception as e:
        return {"url": url, "error": str(e)}


def _id_of(value: Any) -> Any:
    return value.get("id") if isinstance(value, dict) else value


# =========================================================
# USERS
# =========================================================
def find_user_by_email(*, email: str) -> Optional[Dict[str, Any]]:
    """Look up an existing user by email. Re-matched on the address, since
    the filter parameter is not trusted. Never raises."""
    try:
        resp = requests.get(f"{SMAPPLY_BASE}/users/", headers=_headers(),
                            params={"email": email, "limit": 25}, timeout=30)
        if not resp.ok:
            return None
        body = resp.json()
    except Exception:
        return None
    rows = body if isinstance(body, list) else (body.get("results") or [])
    target = email.strip().lower()
    for row in rows:
        if isinstance(row, dict) and (row.get("email") or "").strip().lower() == target:
            return row
    return None


def create_user(*, email: str, first_name: str, last_name: str,
                pilot_program_value: str,
                reuse_existing_user: bool) -> Tuple[int, Dict[str, Any], str]:
    """Create the applicant, or reuse an existing account with this address.

    Returns (user_id, user_object, user_mode) where user_mode is "created" or
    "reused_existing". A reused account has both custom fields rewritten.
    """
    custom_fields = {
        "applicant": [
            {"id": PILOT_PROGRAM_CUSTOM_FIELD_ID, "value": pilot_program_value},
            {"id": APPLICANT_EMAIL_CUSTOM_FIELD_ID, "value": email},
        ],
    }
    payload = {
        "email": email,
        "first_name": first_name,
        "last_name": last_name,
        "roles": ["applicant"],
        "custom_fields": custom_fields,
    }

    try:
        data = _request("POST", f"{SMAPPLY_BASE}/users/",
                        headers=_headers(automations=True), json=payload)
    except RuntimeError as e:
        low = str(e).lower()
        duplicate = "already" in low or "exist" in low or "unique" in low
        if not (reuse_existing_user and duplicate):
            raise
        existing = find_user_by_email(email=email)
        if existing is None:
            raise RuntimeError(
                f"{e}\nThe address appears to be taken but no matching user "
                f"could be found, so nothing was reused."
            )
        uid = int(existing["id"])
        _request("PATCH", f"{SMAPPLY_BASE}/users/{uid}/",
                 headers=_headers(automations=True),
                 json={"custom_fields": custom_fields})
        return uid, existing, "reused_existing"

    user_id = data.get("id")
    if user_id is None:
        raise RuntimeError(f"User created but no 'id' in response: {data}")
    return int(user_id), data, "created"


def get_user(*, user_id: int) -> Dict[str, Any]:
    return _request("GET", f"{SMAPPLY_BASE}/users/{user_id}/", headers=_headers())


def find_custom_field_value(user_obj: Dict[str, Any], field_id: int) -> Any:
    found: List[Any] = []

    def scan(node: Any) -> None:
        if isinstance(node, list):
            for item in node:
                if isinstance(item, dict) and str(item.get("id")) == str(field_id):
                    found.append(item.get("value"))
                else:
                    scan(item)
        elif isinstance(node, dict):
            for value in node.values():
                scan(value)

    scan(user_obj.get("custom_fields"))
    return found[0] if found else None


def set_user_eligibility(*, user_id: int, full_name: str) -> Dict[str, Any]:
    payload = {
        "is_completed": True,
        "responses": {
            ELIG_Q_CONNECTED: ELIG_CONNECTED_RESPONSE,
            ELIG_Q_TERMS:     ELIG_TERMS_SIGNATURE,
            ELIG_Q_FULL_NAME: full_name,
            ELIG_Q_DOCS:      ELIG_DOCS_RESPONSE,
            ELIG_Q_PRIOR:     ELIG_PRIOR_RESPONSE,
            ELIG_Q_APPLY:     ELIG_APPLY_RESPONSE,
        },
    }
    return _request("PUT", f"{SMAPPLY_BASE}/users/{user_id}/eligibility/",
                    headers=_headers(), json=payload)


def diagnose_eligibility(*, user_id: int) -> Dict[str, Any]:
    return {
        "eligibility_readback": _safe_get(f"{SMAPPLY_BASE}/users/{user_id}/eligibility/"),
        "user": _safe_get(f"{SMAPPLY_BASE}/users/{user_id}/"),
        "program": _safe_get(f"{SMAPPLY_BASE}/programs/{PROGRAM_ID}/"),
    }


# =========================================================
# APPLICATIONS
# =========================================================
def find_application_for_user(*, user_id: int) -> Optional[Dict[str, Any]]:
    """One application per applicant per program; a second POST is rejected
    as "not eligible", so check first. Never raises."""
    try:
        resp = requests.get(f"{SMAPPLY_BASE}/applications/", headers=_headers(),
                            params={"applicant": user_id, "program": PROGRAM_ID,
                                    "limit": 25}, timeout=30)
        if not resp.ok:
            return None
        body = resp.json()
    except Exception:
        return None
    rows = body if isinstance(body, list) else (body.get("results") or [])
    for row in rows:
        if (isinstance(row, dict)
                and str(_id_of(row.get("applicant"))) == str(user_id)
                and str(_id_of(row.get("program"))) == str(PROGRAM_ID)):
            return row
    return None


def create_application(*, user_id: int) -> int:
    data = _request("POST", f"{SMAPPLY_BASE}/applications/",
                    headers=_headers(automations=True),
                    json={"applicant": user_id, "program": PROGRAM_ID,
                          "current_stage": INITIAL_STAGE_ID})
    app_id = data.get("id")
    if app_id is None:
        raise RuntimeError(f"Application created but no 'id' in response: {data}")
    return int(app_id)


def _label_ids(labels: Any) -> List[int]:
    out: List[int] = []
    for item in (labels or []):
        try:
            out.append(int(_id_of(item)))
        except (TypeError, ValueError):
            continue
    return out


def assign_label(*, app_id: int, label_id: int) -> List[int]:
    """Labels READ as objects but WRITE as a bare id list that replaces the
    set, so merge first. Confirmed on readback."""
    url = f"{SMAPPLY_BASE}/applications/{app_id}/"
    existing = _label_ids(_request("GET", url, headers=_headers()).get("labels"))
    if label_id in existing:
        return existing
    _request("PATCH", url, headers=_headers(automations=True),
             json={"labels": existing + [label_id]})
    after = _label_ids(_request("GET", url, headers=_headers()).get("labels"))
    if label_id not in after:
        raise RuntimeError(
            f"Label {label_id} not on application {app_id} after PATCH "
            f"(labels now {after})."
        )
    return after


def complete_discipline_task(*, app_id: int, discipline: str,
                             rb_index: int) -> Dict[str, Any]:
    """Writes take bare index lists; data and status in one PATCH 500s, so
    two PATCHes."""
    rb_field = Q_RB_RN if discipline == "RN" else Q_RB_LPN
    url = f"{SMAPPLY_BASE}/applications/{app_id}/tasks/{DISCIPLINE_TASK_ID}/"
    headers = _headers(automations=True)
    data_resp = _request("PATCH", url, headers=headers, json={"data": {
        Q_PROFESSION: [PROFESSION_RESPONSE[discipline]],
        rb_field:     [rb_index],
    }})
    status_resp = _request("PATCH", url, headers=headers,
                           json={"status": SMAPPLY_STATUS_COMPLETED})
    return {"data_patch": data_resp, "status_patch": status_resp}


# =========================================================
# INVITE (Power Automate instant cloud flow)
# =========================================================
def call_invite_flow(*, payload: Dict[str, Any]) -> Dict[str, Any]:
    """POST the applicant to the flow. One call, no polling.

    The flow is expected to send the email and then return 200 with a JSON
    body containing at least {"ok": true}. A 2xx with ok true is reported as
    sent. Never raises: onboarding is already complete at this point, so a
    failure here becomes a warning row in the app, not an error.
    """
    def result(sent: bool, status: str, message: str, error: Any = None,
               http_status: Any = None, body: Any = None) -> Dict[str, Any]:
        return {"sent": sent, "status": status, "message": message,
                "error": error, "http_status": http_status,
                "flow_response": body}

    try:
        url = wmill.get_variable(INVITE_FLOW_URL_VAR)
    except Exception:
        url = None
    if not url:
        return result(False, "flow_not_configured",
                      f"Windmill variable {INVITE_FLOW_URL_VAR} is not set; "
                      f"no invitation was sent.")

    try:
        key = wmill.get_variable(INVITE_FLOW_KEY_VAR)
    except Exception:
        key = None
    if not key:
        return result(False, "flow_key_missing",
                      f"Windmill variable {INVITE_FLOW_KEY_VAR} is not set; "
                      f"no invitation was sent.")
    headers = {"Content-Type": "application/json", INVITE_FLOW_KEY_HEADER: key}

    try:
        resp = requests.post(url, json=payload, headers=headers,
                             timeout=INVITE_FLOW_TIMEOUT)
    except requests.Timeout:
        return result(False, "flow_timeout",
                      f"The invite flow did not respond within "
                      f"{INVITE_FLOW_TIMEOUT}s. It may still have sent the "
                      f"email; check the flow's run history before retrying.",
                      error="timeout")
    except Exception as e:
        return result(False, "flow_unreachable",
                      "Could not reach the invite flow.", error=str(e))

    try:
        body: Any = resp.json()
    except Exception:
        body = {"text": resp.text[:500]}

    if resp.status_code in (401, 403):
        return result(False, "flow_unauthorized",
                      f"The invite flow rejected the request "
                      f"({resp.status_code}). Check that "
                      f"{INVITE_FLOW_KEY_VAR} matches the value in the flow.",
                      error=str(body)[:500], http_status=resp.status_code, body=body)

    flow_ok = resp.ok and (not isinstance(body, dict) or body.get("ok", True) is not False)
    if flow_ok:
        return result(True, "sent",
                      f"Invitation sent to {payload['email']} by the flow.",
                      http_status=resp.status_code, body=body)
    return result(False, "flow_error",
                  f"The invite flow returned {resp.status_code}.",
                  error=str(body)[:500], http_status=resp.status_code, body=body)


# =========================================================
# ENTRYPOINT
# =========================================================
def main(
    email: str,
    first_name: str,
    last_name: str,
    regulator: Literal["RN", "LPN"],
    regulatory_body: Literal[
        "College of Nursing of New Brunswick",
        "Association Of New Brunswick Licensed Practical Nurses",
        "NL College of Nurses",
        "PEI College of Nursing & Midwifery",
        "Nova Scotia College of Nursing",
        "College of Nurses of Ontario",
        "College of Registered Nurses of Manitoba",
        "College of Licensed Practical Nurses of Manitoba",
        "College of Registered Nurses of Saskatchewan",
        "College of Licensed Practical Nurses of Saskatchewan",
        "College of Registered Nurses of Alberta",
        "College of Licensed Practical Nurses and Health Care Aides of Alberta",
        "BC College of Nurses and Midwives",
    ],
    language: Literal["English", "French"],
    send_invite: bool = True,
    apply_nb_pilot_label: bool = True,
    reuse_existing_user: bool = True,
):
    email = email.strip()
    entry = resolve_body(regulatory_body, regulator)
    pilot_program_value = PILOT_PROGRAM_BY_LANGUAGE.get(language)
    if pilot_program_value is None:
        raise ValueError(f"Unsupported language '{language}'.")
    full_name = f"{first_name} {last_name}".strip()

    # 1) User.
    user_id, user_resp, user_mode = create_user(
        email=email, first_name=first_name, last_name=last_name,
        pilot_program_value=pilot_program_value,
        reuse_existing_user=reuse_existing_user,
    )

    # 2) Both custom fields must have landed; the invite and any automation
    # depend on them.
    user_after = get_user(user_id=user_id)
    pilot_on_user = find_custom_field_value(user_after, PILOT_PROGRAM_CUSTOM_FIELD_ID)
    email_cf_on_user = find_custom_field_value(user_after, APPLICANT_EMAIL_CUSTOM_FIELD_ID)
    if pilot_on_user != pilot_program_value:
        raise RuntimeError(
            f"Custom field {PILOT_PROGRAM_CUSTOM_FIELD_ID} did not persist as "
            f"{pilot_program_value!r} on user {user_id} (GET shows {pilot_on_user!r})."
        )
    if (email_cf_on_user or "").strip().lower() != email.lower():
        raise RuntimeError(
            f"Custom field {APPLICANT_EMAIL_CUSTOM_FIELD_ID} did not persist as "
            f"{email!r} on user {user_id} (GET shows {email_cf_on_user!r})."
        )

    # 3) Eligibility.
    eligibility_resp = set_user_eligibility(user_id=user_id, full_name=full_name)

    # 4) Application.
    existing_app = find_application_for_user(user_id=user_id)
    if existing_app is not None:
        app_id = int(existing_app["id"])
        application_mode = "reused_existing"
    else:
        application_mode = "created"
        app_id = None
        for attempt in range(ELIGIBILITY_RETRIES):
            try:
                app_id = create_application(user_id=user_id)
                break
            except RuntimeError as e:
                if "not eligible" not in str(e).lower():
                    raise
                late = find_application_for_user(user_id=user_id)
                if late is not None:
                    app_id = int(late["id"])
                    application_mode = "reused_existing"
                    break
                if attempt < ELIGIBILITY_RETRIES - 1:
                    time.sleep(ELIGIBILITY_RETRY_DELAY)
                    continue
                diag = diagnose_eligibility(user_id=user_id)
                elig_body = diag["eligibility_readback"].get("body")
                raise RuntimeError(
                    f"{e}\n--- eligibility diagnostics ---\n"
                    f"user {user_id} ({user_mode}) has no application on "
                    f"program {PROGRAM_ID}, so this is a genuine eligibility "
                    f"failure. eligibility is_completed: "
                    f"{elig_body.get('is_completed') if isinstance(elig_body, dict) else elig_body}\n"
                    f"user: {diag['user']}"
                )

    # 5) Label.
    label_ids = assign_label(app_id=app_id, label_id=NB_PILOT_LABEL_ID) \
        if apply_nb_pilot_label else None

    # 6) Discipline task.
    task_resp = complete_discipline_task(
        app_id=app_id, discipline=regulator, rb_index=entry["index"],
    )
    task_status = (task_resp.get("status_patch") or {}).get("status")

    # 7) Invite via the flow.
    flow_payload = {
        "user_id": user_id,
        "application_id": app_id,
        "email": email,
        "first_name": first_name,
        "last_name": last_name,
        "full_name": full_name,
        "language": language,
        "pilot_program": pilot_program_value,
        "regulator": regulator,
        "regulatory_body": regulatory_body,
        "province": entry["province"],
        "program_id": PROGRAM_ID,
        "user_mode": user_mode,
        "application_mode": application_mode,
        "submitted_at": datetime.now(timezone.utc).replace(microsecond=0).isoformat(),
    }
    if send_invite:
        invite = call_invite_flow(payload=flow_payload)
    else:
        invite = {"sent": False, "status": "not_requested",
                  "message": "Invitation not requested.", "error": None,
                  "http_status": None, "flow_response": None}

    overall_status = (
        "completed" if (not send_invite or invite["sent"])
        else "completed_invitation_failed"
    )

    return {
        "ok": True,
        "overall_status": overall_status,
        "onboarding_completed": True,

        "user_id": user_id,
        "user_mode": user_mode,
        "application_id": app_id,
        "application_mode": application_mode,
        "email": email,
        "full_name": full_name,

        "regulator": regulator,
        "regulatory_body": regulatory_body,
        "rb_index": entry["index"],
        "language": language,
        "province": entry["province"],

        "pilot_program_value": pilot_program_value,
        "pilot_program_on_user": pilot_on_user,
        "province_on_user": pilot_on_user,
        "applicant_email_on_user": email_cf_on_user,

        "eligibility_completed": eligibility_resp.get("is_completed"),
        "task_status": task_status,
        "label_applied": NB_PILOT_LABEL_NAME if label_ids else None,
        "label_ids_on_application": label_ids,

        "invite_requested": bool(send_invite),
        "invite_sent": invite["sent"],
        "invite_confirmed": invite["sent"],
        "invite_status": invite["status"],
        "invite_message": invite["message"],
        "invite_error": invite["error"],
        "invite_http_status": invite["http_status"],
        "invite_flow_response": invite["flow_response"],

        "program_id": PROGRAM_ID,
        "task_response": task_resp,
    }