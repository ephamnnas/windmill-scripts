# requirements:
# requests
# wmill

"""
Windmill script - NNAS SMApply onboarding

Flow:
  1) Create (invite) an applicant user via POST /api/users/, setting custom
     field 153504 to the province ("NB" for the currently mapped bodies).
  2) Create an application for that user via POST /api/applications/
     (program 295865, current_stage 1369714).
  3) Assign the "NB Pilot" label to the new application.
  4) PATCH the "Select Nursing Discipline" task (2053929) with the profession +
     regulatory-body responses and mark it complete (status 3).

Runs on a self-hosted Windmill instance. All SMApply calls use the bearer token
in the Windmill variable u/epham/SMAPPLY_ACCESS_TOKEN. Errors surface the
response body explicitly rather than a bare raise_for_status().
"""

import wmill
import requests
import time
from typing import Any, Dict, List, Tuple, Literal

# =========================================================
# CONSTANTS
# =========================================================
SMAPPLY_BASE = "https://nnas.smapply.io/api"
SMAPPLY_TOKEN_VAR = "u/epham2/SMAPPLY_ACCESS_TOKEN"

PROGRAM_ID = 295865
INITIAL_STAGE_ID = 1369714
DISCIPLINE_TASK_ID = 2053929
PROVINCE_CUSTOM_FIELD_ID = 153504

SMAPPLY_STATUS_COMPLETED = 3

# Label applied to every application created by this script.
# SMApply exposes no labels collection endpoint on this instance, so the id is
# fixed here. Harvested from live applications; write shape confirmed by probe.
NB_PILOT_LABEL_ID = 183165
NB_PILOT_LABEL_NAME = "NB Pilot"

# Program-eligibility evaluation can lag the eligibility PUT, so retry the
# application POST a few times before treating "not eligible" as final.
ELIGIBILITY_RETRIES = 5
ELIGIBILITY_RETRY_DELAY = 3  # seconds

# Task 2053929 question keys (from the live task object).
Q_PROFESSION = "4JqOe5pS3h"   # nursing profession select: RN -> 0, LPN -> 1
Q_RB_RN      = "faO2qkS4PX"   # regulatory-body select shown for the RN discipline
Q_RB_LPN     = "IWgvFrNZHX"   # regulatory-body select shown for the LPN discipline

PROFESSION_RESPONSE = {"RN": 0, "LPN": 1}

# ---------------------------------------------------------
# Eligibility form (PUT /api/users/{id}/eligibility/) - program 295865 gates on
# this. Question keys + answers mirror a confirmed-eligible submission; the
# full-name answer is filled per applicant. All six questions are required;
# omitting any (as an earlier partial attempt did) leaves the user "not eligible".
# ---------------------------------------------------------
ELIG_Q_CONNECTED = "8LpmN6CBVM"   # "Is any part of your application connected to..." -> 2
ELIG_Q_TERMS     = "A3fGSIXu9p"   # "I HEREBY AGREE TO THE NNAS TERMS AND CONDITIONS" -> signature
ELIG_Q_FULL_NAME = "C1cK1aI9ni"   # "Please type your full name" -> applicant full name
ELIG_Q_DOCS      = "Db2U6y1U0Q"   # "Are any of the documents, credentials, institutions..." -> null
ELIG_Q_PRIOR     = "FFhD4gHD9J"   # "Have you completed and received a prior NNAS Expedited..." -> 1
ELIG_Q_APPLY     = "qzzA39MQAV"   # "Eligibility to Apply to NNAS" -> 0

ELIG_CONNECTED_RESPONSE = 2
ELIG_DOCS_RESPONSE = None
ELIG_PRIOR_RESPONSE = 1
ELIG_APPLY_RESPONSE = 0
# Per-user signature upload; SMApply re-hosts it under a fresh media path on PUT,
# so this shared reference is accepted for any applicant.
ELIG_TERMS_SIGNATURE = (
    "/protected/media-file/FkmBTc_C-GJ6iD28zx8OmFNuhZEUCtxmqgzNwaYIKWthTpR3JG5U"
    "DkAurInQBnd2u5FwYgql4crOJ9te2NzMZ_pTYbQaJmQErqQKHxKnNU4eYurjY182p4_V9NUYDu"
    "oPP2UUYBbbjflR6XRQgSHZrGvMWnPY9ctiUjl0aTdNCCNf5lKEFuOZhSj-jXJFpRJPzPcvTsKT"
    "bZRsykgUvJeTmFkPTdoc2m97708s2xQDbjnLOftxs4-7OfUQMI-vfpfE__Qf0xmuratkGf24M5"
    "zjA9k9UOFl7znx2_LWBfJ6Kv6ahl922bHhHqtDTqxTWnuWNBJnZM6oVv_ZKe4kfzcWVXKr0jC"
    "GUikqphrZ2L-1FHU=/2026-07-15-18-27-20-signature.svg"
)

# ---------------------------------------------------------
# Regulatory-body registry - the ONE place to extend.
#
# Each entry maps an RB to:
#   discipline : "RN" or "LPN"  -> selects Q_PROFESSION response and which
#                                  RB field (Q_RB_RN vs Q_RB_LPN) is populated
#   province   : value written to custom field 153504
#   index      : 0-based position of this RB *within that discipline's* SMApply
#                dropdown (this is the number that goes in the `response` array)
#
# Keys are normalized (lowercased, whitespace-collapsed) so dropdown-label
# casing/whitespace quirks don't matter.
#
# ONLY the two New Brunswick bodies below are confirmed. To add others you need
# the response index from the live task's per-discipline dropdown ordering
# (do NOT assume it matches the combined input dropdown order).
# ---------------------------------------------------------
RB_REGISTRY: Dict[str, Dict[str, Any]] = {
    "college of nursing of new brunswick": {
        "discipline": "RN", "province": "NB", "index": 2,
    },
    "association of new brunswick licensed practical nurses": {
        "discipline": "LPN", "province": "NB", "index": 2,
    },
}


def _rb_key(name: str) -> str:
    return " ".join((name or "").lower().split())


# =========================================================
# HTTP HELPERS
# =========================================================
def _headers(extra: Dict[str, str] = None) -> Dict[str, str]:
    token = wmill.get_variable(SMAPPLY_TOKEN_VAR)
    h = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    if extra:
        h.update(extra)
    return h


def _request(method: str, url: str, **kwargs) -> Dict[str, Any]:
    """Make a request, surfacing the response body on failure."""
    resp = requests.request(method, url, **kwargs)
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


# =========================================================
# SMAPPLY OPERATIONS
# =========================================================
def create_user(*, email: str, first_name: str, last_name: str,
                province: str, send_invite: bool = True) -> Tuple[int, Dict[str, Any]]:
    """POST /api/users/ - creates/invites an applicant with the province custom field.

    SMApply automatically sends its account-activation (invitation) email when an
    applicant is created without a completed signup, so send_invite=True just uses
    the minimal payload. email_subscribed is NOT set here: it governs ongoing
    notifications and is user-only ("only user can set value" -> 403 for admins).
    send_invite=False marks the user as already onboarded to suppress the invite.
    """
    payload = {
        "email": email,
        "first_name": first_name,
        "last_name": last_name,
        "roles": ["applicant"],
        "custom_fields": {
            "applicant": [
                {"id": PROVINCE_CUSTOM_FIELD_ID, "value": province},
            ],
        },
    }
    if not send_invite:
        # Best-effort: mark as already signed up/verified so the automatic
        # activation email isn't sent.
        payload["signed_up"] = True
        payload["email_verified"] = True
    data = _request("POST", f"{SMAPPLY_BASE}/users/", headers=_headers(), json=payload)
    user_id = data.get("id")
    if user_id is None:
        raise RuntimeError(f"User created but no 'id' in response: {data}")
    return int(user_id), data


def _safe_get(url: str) -> Dict[str, Any]:
    """Best-effort GET for diagnostics; never raises, always returns a dict."""
    try:
        resp = requests.get(url, headers=_headers())
        try:
            body = resp.json()
        except Exception:
            body = {"text": resp.text}
        return {"url": url, "status": resp.status_code, "body": body}
    except Exception as e:
        return {"url": url, "error": str(e)}


def diagnose_eligibility(*, user_id: int) -> Dict[str, Any]:
    """Pull the pieces needed to see why an applicant is 'not eligible'.

    Surfaces the applicant's stored eligibility (is_completed + responses), their
    custom fields, the definition of field 153504, and the program config.
    Endpoints are best-effort; if one 404s the others still return.
    """
    return {
        "eligibility_readback": _safe_get(f"{SMAPPLY_BASE}/users/{user_id}/eligibility/"),
        "user": _safe_get(f"{SMAPPLY_BASE}/users/{user_id}/"),
        "custom_field_153504": _safe_get(
            f"{SMAPPLY_BASE}/customfields/{PROVINCE_CUSTOM_FIELD_ID}/"
        ),
        "program": _safe_get(f"{SMAPPLY_BASE}/programs/{PROGRAM_ID}/"),
    }


def get_user(*, user_id: int) -> Dict[str, Any]:
    """GET /api/users/{id}/ - read the user back to inspect persisted custom fields."""
    return _request("GET", f"{SMAPPLY_BASE}/users/{user_id}/", headers=_headers())


def find_custom_field_value(user_obj: Dict[str, Any], field_id: int) -> Any:
    """Locate a custom field's value in a user object, tolerant of shape.

    custom_fields may come back keyed by role ({"applicant": [...]}) or as a
    flat list; this scans either and matches on id (compared as strings).
    Returns the value if found, else None.
    """
    found = []

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
    """PUT /api/users/{id}/eligibility/ - submit the completed eligibility form.

    Program 295865 requires completed eligibility before an application can be
    created. Responses use the documented flat {question_id: value} shape.
    Requires the token to have the 'Assume role' permission.
    """
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
    url = f"{SMAPPLY_BASE}/users/{user_id}/eligibility/"
    return _request("PUT", url, headers=_headers(), json=payload)


def create_application(*, user_id: int) -> Tuple[int, Dict[str, Any]]:
    """POST /api/applications/ - program 295865, current_stage 1369714."""
    payload = {
        "applicant": user_id,
        "program": PROGRAM_ID,
        "current_stage": INITIAL_STAGE_ID,
    }
    data = _request("POST", f"{SMAPPLY_BASE}/applications/", headers=_headers(), json=payload)
    app_id = data.get("id")
    if app_id is None:
        raise RuntimeError(f"Application created but no 'id' in response: {data}")
    return int(app_id), data


# ---------------------------------------------------------
# LABELS
# ---------------------------------------------------------
def _label_ids(labels: Any) -> List[int]:
    """Extract ids from an application's `labels` value.

    GET serializes labels as [{"id": int, "name": str}]; tolerate bare ids too.
    """
    out: List[int] = []
    for item in (labels or []):
        raw = item.get("id") if isinstance(item, dict) else item
        try:
            out.append(int(raw))
        except (TypeError, ValueError):
            continue
    return out


def assign_label(*, app_id: int, label_id: int) -> Dict[str, Any]:
    """PATCH /api/applications/{id}/ with the merged label id list.

    Confirmed by probe: the serializer READS labels as nested objects but
    WRITES them as a bare id list, e.g. {"labels": [183165]}. The list replaces
    the full set, so existing labels are read first and merged in to avoid
    dropping anything an automation applied at creation. Success is confirmed
    on readback, not on status code alone.
    """
    url = f"{SMAPPLY_BASE}/applications/{app_id}/"

    current = _request("GET", url, headers=_headers())
    existing = _label_ids(current.get("labels"))
    if label_id in existing:
        return {"already_present": True, "label_ids": existing}

    merged = existing + [label_id]
    patch_resp = _request("PATCH", url, headers=_headers(), json={"labels": merged})

    after = _label_ids(
        _request("GET", url, headers=_headers()).get("labels")
    )
    if label_id not in after:
        raise RuntimeError(
            f"PATCH {url} labels={merged} returned 2xx but label {label_id} is "
            f"not on application {app_id} after readback (labels now {after}). "
            f"PATCH response labels: {patch_resp.get('labels')!r}"
        )
    return {"already_present": False, "label_ids": after}


def complete_discipline_task(*, app_id: int, discipline: str,
                             rb_index: int) -> Dict[str, Any]:
    """Write the discipline task responses, then mark it complete.

    Two PATCHes: the task GET echoes each field as {"label", "response"}, but
    WRITES only accept bare option-index lists keyed by field id
    (e.g. {"data": {"4JqOe5pS3h": [0]}}). The {"label", "response"} wrapper 500s
    server-side (confirmed by probe). Only the fields in play are sent; the server
    auto-populates the hidden conditional RB field as empty. Writing data and
    flipping to status 3 in one request also 500s, so data is saved first, then
    the status-only mark-complete PATCH runs.
    """
    rb_field = Q_RB_RN if discipline == "RN" else Q_RB_LPN

    url = f"{SMAPPLY_BASE}/applications/{app_id}/tasks/{DISCIPLINE_TASK_ID}/"
    headers = _headers({"apply-automations": "full"})

    # 1) Write the responses (raw lists, only the fields in play).
    data_payload = {
        "data": {
            Q_PROFESSION: [PROFESSION_RESPONSE[discipline]],  # RN -> [0], LPN -> [1]
            rb_field:     [rb_index],
        },
    }
    data_resp = _request("PATCH", url, headers=headers, json=data_payload)

    # 2) Mark complete (status-only; completed_at is server-set / read-only).
    status_resp = _request(
        "PATCH", url, headers=headers, json={"status": SMAPPLY_STATUS_COMPLETED}
    )
    return {"data_patch": data_resp, "status_patch": status_resp}


# =========================================================
# ENTRYPOINT
# =========================================================
def main(
    email: str,
    first_name: str,
    last_name: str,
    regulator: Literal["RN", "LPN"],
    regulatory_body: Literal[
        "Association Of New Brunswick Licensed Practical Nurses",
        "College Of Licensed Practical Nurses Of British Columbia",
        "College Of Licensed Practical Nurses Of Manitoba",
        "College Of Nurses Of Ontario",
        "College Of Registered Nurses Of British Columbia",
        "College Of Registered Nurses Of Manitoba",
        "College Of Registered Nurses Of Prince Edward Island",
        "College of Licensed Practical Nurses and Health Care Aides of Alberta",
        "College of Licensed Practical Nurses of Prince Edward Island",
        "College of Licensed Practical Nurses of Saskatchewan",
        "College of Nursing of New Brunswick",
        "College of Registered Nurses Of Alberta",
        "College of Registered Nurses of Saskatchewan",
        "Newfoundland and Labrador College of Nurses",
        "Nova Scotia College of Nursing-LPN",
        "Nova Scotia College of Nursing-RN",
    ],
    send_invite: bool = True,
    apply_nb_pilot_label: bool = True,
):
    # --- Resolve & validate the regulatory body ---
    entry = RB_REGISTRY.get(_rb_key(regulatory_body))
    if entry is None:
        raise ValueError(
            f"No SMApply mapping configured for regulatory body "
            f"'{regulatory_body}'. Add it to RB_REGISTRY (discipline, province, "
            f"and its response index within that discipline's SMApply dropdown)."
        )
    if entry["discipline"] != regulator:
        raise ValueError(
            f"Regulator '{regulator}' does not match '{regulatory_body}', which "
            f"is a {entry['discipline']} body."
        )

    province = entry["province"]
    rb_index = entry["index"]

    # 1) Create/invite the applicant user.
    user_id, user_resp = create_user(
        email=email, first_name=first_name, last_name=last_name,
        province=province, send_invite=send_invite,
    )

    # 1a) Read the user back and confirm the province custom field actually
    # persisted. Program 295865's eligibility gate reads this field, so if it
    # didn't land the application POST 403s with "not eligible for this program".
    # Failing here (before the app call) makes the root cause unambiguous.
    user_after = get_user(user_id=user_id)
    province_on_user = find_custom_field_value(user_after, PROVINCE_CUSTOM_FIELD_ID)
    if province_on_user != province:
        raise RuntimeError(
            f"Custom field {PROVINCE_CUSTOM_FIELD_ID} did not persist as {province!r} "
            f"on user {user_id} (GET shows {province_on_user!r}). The program "
            f"eligibility gate reads this field, so create_application will 403. "
            f"POST-create custom_fields echo: {user_resp.get('custom_fields')!r}"
        )

    # 1b) Submit the eligibility form. Program 295865 gates the application on
    # completed eligibility, so this must run before create_application.
    full_name = f"{first_name} {last_name}".strip()
    eligibility_resp = set_user_eligibility(user_id=user_id, full_name=full_name)

    # 2) Create the application. Retry on "not eligible" in case the
    # eligibility->program evaluation lags the PUT above.
    app_id = None
    app_resp = None
    for attempt in range(ELIGIBILITY_RETRIES):
        try:
            app_id, app_resp = create_application(user_id=user_id)
            break
        except RuntimeError as e:
            if "not eligible" not in str(e).lower():
                raise
            if attempt < ELIGIBILITY_RETRIES - 1:
                time.sleep(ELIGIBILITY_RETRY_DELAY)
                continue
            # Final attempt failed; surface everything useful.
            diag = diagnose_eligibility(user_id=user_id)
            raise RuntimeError(
                f"{e}\n--- eligibility diagnostics "
                f"(after {ELIGIBILITY_RETRIES} attempts) ---\n"
                f"eligibility readback is_completed: "
                f"{diag['eligibility_readback'].get('body', {}).get('is_completed') if isinstance(diag['eligibility_readback'].get('body'), dict) else diag['eligibility_readback']}\n"
                f"full user object: {diag['user']}\n"
                f"program access={diag['program'].get('body', {}).get('access') if isinstance(diag['program'].get('body'), dict) else '?'} "
                f"applicant_type={diag['program'].get('body', {}).get('applicant_type') if isinstance(diag['program'].get('body'), dict) else '?'}"
            )

    # 2a) Label the application. Done before the task write so the label is on
    # the record regardless of what the automations fired by step 3 do.
    label_resp = None
    if apply_nb_pilot_label:
        label_resp = assign_label(app_id=app_id, label_id=NB_PILOT_LABEL_ID)

    # 3) Populate + complete the discipline-select task.
    task_resp = complete_discipline_task(
        app_id=app_id, discipline=regulator, rb_index=rb_index,
    )

    return {
        "user_id": user_id,
        "application_id": app_id,
        "regulator": regulator,
        "regulatory_body": regulatory_body,
        "province": province,
        "province_on_user": province_on_user,
        "eligibility_completed": eligibility_resp.get("is_completed"),
        "rb_index": rb_index,
        "label_applied": NB_PILOT_LABEL_NAME if label_resp else None,
        "label_ids_on_application": (label_resp or {}).get("label_ids"),
        "task_status": (task_resp.get("status_patch") or {}).get("status"),
        "task_response": task_resp,
    }